"""
Core mail tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - list_emails
  - read_email
  - search_emails
  - filter_emails
  - send_email
  - reply_email
  - forward_email
  - mark_as_read
  - mark_as_unread
  - move_email
  - trash_email
  - delete_email
  - bulk_move_emails
  - bulk_trash_emails
  - bulk_delete_emails
"""

from __future__ import annotations

import re
from typing import Any, Callable, Literal, Optional

from fastmcp.server.context import Context
from pydantic import BaseModel, Field

from mcp_microsoft.common.mail_utils import (
    format_mail_datetime,
    parse_recipients,
    recipient_values,
)
from mcp_microsoft.models import (
    AttachmentInfo,
    BulkDeleteEmailsResponse,
    BulkEmailFailure,
    BulkMovedEmail,
    BulkMoveEmailsResponse,
    BulkTrashEmailsResponse,
    DeleteEmailResponse,
    DisplayAddress,
    ForwardEmailResponse,
    ListEmailsResponse,
    MarkEmailReadResponse,
    MessageSummary,
    MoveEmailResponse,
    ReadEmailResponse,
    ReadEmailSummaryResponse,
    ReplyEmailResponse,
    SearchEmailsResponse,
    SendEmailResponse,
    TrashEmailResponse,
)
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.text import strip_html
from mcp_microsoft.common.tooling import (
    DESTRUCTIVE_TOOL,
    IDEMPOTENT_WRITE_TOOL,
    READ_ONLY_TOOL,
    WRITE_TOOL,
    register_tool,
)
from mcp_microsoft.feature_flags import is_deletion_disabled
from mcp_microsoft.graph_types import (
    GraphAttachment,
    GraphItemBody,
    GraphMessage,
    GraphRecipient,
    GraphSender,
    parse_graph_collection,
)
from mcp_microsoft.graph import get_graph

# ---------------------------------------------------------------------------
# Elicitation helpers
# ---------------------------------------------------------------------------


class _Confirmation(BaseModel):
    confirmed: bool


class _BatchRequestEntry(BaseModel):
    id: str
    method: str
    url: str
    body: dict[str, Any] | None = None
    headers: dict[str, str] | None = None


class _BatchOperation(BaseModel):
    method: str
    url_template: str
    body: dict[str, Any] | None = None


class _BatchRequestPlan(BaseModel):
    requests: list[_BatchRequestEntry] = Field(default_factory=list)
    message_ids_by_batch_id: dict[str, str] = Field(default_factory=dict)


class _BatchErrorBody(BaseModel):
    code: str | None = None
    message: str | None = None


class _BatchResponseBody(BaseModel):
    error: _BatchErrorBody | None = None


class _BatchResponseEntry(BaseModel):
    id: str = ""
    status: int = 0
    body: dict[str, Any] | None = None


class _BatchExecutionResult(BaseModel):
    succeeded_message_ids: list[str] = Field(default_factory=list)
    failures: list[BulkEmailFailure] = Field(default_factory=list)
    response_messages: dict[str, GraphMessage] = Field(default_factory=dict)


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

BodyType = Literal["text", "html"]


class ListEmailsInput(ToolRequestModel):
    folder: str = "inbox"
    max_results: int = 10
    unread_only: bool = False
    sort_order: Literal["newest", "oldest"] = "newest"
    skip_token: str | None = None
    profile: str | None = None


class ReadEmailInput(ToolRequestModel):
    message_id: str
    summary_mode: bool = False
    profile: str | None = None


class SearchEmailsInput(ToolRequestModel):
    query: str
    max_results: int = 10
    folder: str | None = None
    profile: str | None = None


class FilterEmailsInput(ToolRequestModel):
    from_address: str | None = None
    to_address: str | None = None
    subject_contains: str | None = None
    received_after: str | None = None
    received_before: str | None = None
    has_attachments: bool | None = None
    importance: Literal["low", "normal", "high"] | None = None
    folder: str = "inbox"
    max_results: int = 50
    sort_order: Literal["newest", "oldest"] = "newest"
    skip_token: str | None = None
    profile: str | None = None


class SendEmailInput(ToolRequestModel):
    to: str | list[str]
    subject: str
    body: str
    cc: str | list[str] | None = None
    bcc: str | list[str] | None = None
    body_type: BodyType = "text"
    save_to_sent: bool = True
    reply_to: str | list[str] | None = None
    profile: str | None = None
    confirm: bool = False


class ReplyEmailInput(ToolRequestModel):
    message_id: str
    body: str
    reply_all: bool = False
    body_type: BodyType = "text"
    profile: str | None = None


class ForwardEmailInput(ToolRequestModel):
    message_id: str
    to: str | list[str]
    comment: str | None = None
    profile: str | None = None


class MarkAsReadInput(ToolRequestModel):
    message_id: str
    profile: str | None = None


class MarkAsUnreadInput(ToolRequestModel):
    message_id: str
    profile: str | None = None


class MoveEmailInput(ToolRequestModel):
    message_id: str
    destination_folder: str
    profile: str | None = None


class TrashEmailInput(ToolRequestModel):
    message_id: str
    profile: str | None = None


class DeleteEmailInput(ToolRequestModel):
    message_id: str
    profile: str | None = None
    confirm: bool = False


class BulkMoveEmailsInput(ToolRequestModel):
    message_ids: list[str] | None = None
    destination_folder: str = ""
    source_folder: str | None = None
    profile: str | None = None


class BulkTrashEmailsInput(ToolRequestModel):
    message_ids: list[str] | None = None
    folder: str | None = None
    profile: str | None = None
    confirm: bool = False


class BulkDeleteEmailsInput(ToolRequestModel):
    message_ids: list[str] | None = None
    folder: str | None = None
    profile: str | None = None
    confirm: bool = False


def _body_text(body: GraphItemBody) -> str:
    """Convert a Graph message body into readable plain text when needed."""
    if body.content_type.lower() == "html":
        return strip_html(body.content)
    return body.content


def _fmt_date(iso: Optional[str]) -> str:
    """Format an ISO 8601 date string to a human-readable form."""
    return format_mail_datetime(iso)


def _fmt_sender(sender_obj: GraphSender | None) -> str:
    """Format a Graph sender/from object as 'Name <email>' or just 'email'."""
    if not sender_obj:
        return "unknown"
    ea = sender_obj.email_address
    name = ea.name
    addr = ea.address or "unknown"
    return f"{name} <{addr}>" if name else addr


def _fmt_recipients(recipients: list[GraphRecipient]) -> str:
    """Format a list of Graph recipient objects as a comma-separated string."""
    parts = []
    for r in recipients or []:
        ea = r.email_address
        name = ea.name
        addr = ea.address
        parts.append(f"{name} <{addr}>" if name else addr)
    return ", ".join(parts)


def _display_address_from_sender(sender_obj: GraphSender | None) -> DisplayAddress:
    """Normalize a Graph sender into a typed address model."""
    email_address = sender_obj.email_address if sender_obj else None
    return DisplayAddress(
        display=_fmt_sender(sender_obj),
        name=email_address.name if email_address else "",
        address=email_address.address if email_address else "",
    )


def _message_summary(msg: GraphMessage) -> MessageSummary:
    """Normalize a Graph message into a summary payload."""
    return MessageSummary(
        id=msg.id,
        subject=msg.subject or "(no subject)",
        from_=_display_address_from_sender(msg.from_),
        received_at=msg.received_date_time,
        received_at_display=_fmt_date(msg.received_date_time),
        is_read=msg.is_read,
        has_attachments=msg.has_attachments,
        importance=msg.importance,
        preview=msg.body_preview.replace("\n", " ")[:120],
    )


def _attachment_info(att: GraphAttachment) -> AttachmentInfo:
    size_kb = att.size // 1024
    return AttachmentInfo(
        id=att.id,
        name=att.name or "unnamed",
        content_type=att.content_type,
        size_bytes=att.size,
        size_kb=size_kb,
    )


def _graph_message_from_payload(payload: dict[str, Any] | None) -> GraphMessage | None:
    """Parse a Graph message response body when the operation returns one."""
    if not isinstance(payload, dict):
        return None
    return GraphMessage.model_validate(payload)


# ---------------------------------------------------------------------------
# list_emails
# ---------------------------------------------------------------------------


async def list_emails(
    params: ListEmailsInput,
) -> ListEmailsResponse:
    """
    List emails from a mail folder.

    Args:
        folder: Well-known folder name or folder ID.
                Well-known names: inbox, sentitems, drafts, deleteditems,
                junkemail, archive. Defaults to 'inbox'.
        max_results: Maximum number of messages to return (1-100). Defaults to 10.
        unread_only: When True, return only unread messages. Defaults to False.
        sort_order: 'newest' (default) or 'oldest' first.
        skip_token: Opaque pagination cursor returned as next_page_token from a
                    previous call. Omit for the first page.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured message summaries with pagination metadata. When has_more is
        True, pass next_page_token as skip_token to retrieve the next page.
    """
    from urllib.parse import parse_qs, urlparse

    g = get_graph(params.profile)
    order = "receivedDateTime asc" if params.sort_order == "oldest" else "receivedDateTime desc"
    query: dict[str, Any] = {
        "$top": params.max_results,
        "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments,importance",
        "$orderby": order,
    }
    if params.unread_only:
        # Graph requires an $orderby property to appear first in $filter.
        query["$filter"] = (
            "receivedDateTime ge 1900-01-01T00:00:00Z and isRead eq false"
        )
    if params.skip_token is not None:
        query["$skiptoken"] = params.skip_token

    result = await g.get(f"/me/mailFolders/{params.folder}/messages", params=query)

    messages = parse_graph_collection(result, GraphMessage)
    next_link = result.get("@odata.nextLink", "")

    next_page_token: str | None = None
    if next_link:
        qs = parse_qs(urlparse(next_link).query)
        next_page_token = qs.get("$skiptoken", [None])[0]

    return ListEmailsResponse(
        folder=params.folder,
        count=len(messages),
        messages=[_message_summary(msg) for msg in messages],
        next_page_token=next_page_token,
        has_more=(next_page_token is not None),
    )


# ---------------------------------------------------------------------------
# read_email
# ---------------------------------------------------------------------------


async def read_email(
    params: ReadEmailInput,
) -> ReadEmailResponse | ReadEmailSummaryResponse:
    """
    Fetch a full email message by ID.

    Args:
        message_id: The Graph message ID.
        summary_mode: When True, return only subject, from, date, and body preview
                      instead of the full body.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured message details.
    """
    g = get_graph(params.profile)
    query = {
        "$select": (
            "id,subject,from,toRecipients,ccRecipients,receivedDateTime,"
            "body,bodyPreview,attachments,isRead,conversationId,importance"
        ),
        "$expand": "attachments($select=id,name,size,contentType)",
    }

    msg = GraphMessage.model_validate(await g.get(f"/me/messages/{params.message_id}", params=query))

    subject = msg.subject or "(no subject)"
    to_str = _fmt_recipients(msg.to_recipients)
    cc_str = _fmt_recipients(msg.cc_recipients)
    date = _fmt_date(msg.received_date_time)
    is_read = msg.is_read
    conv_id = msg.conversation_id

    if params.summary_mode:
        return ReadEmailSummaryResponse(
            id=params.message_id,
            subject=subject,
            from_=_display_address_from_sender(msg.from_),
            received_at=msg.received_date_time,
            received_at_display=date,
            is_read=is_read,
            preview=msg.body_preview.replace("\n", " "),
        )

    return ReadEmailResponse(
        id=params.message_id,
        subject=subject,
        from_=_display_address_from_sender(msg.from_),
        to=recipient_values(msg.to_recipients),
        to_display=to_str,
        cc=recipient_values(msg.cc_recipients),
        cc_display=cc_str,
        received_at=msg.received_date_time,
        received_at_display=date,
        is_read=is_read,
        conversation_id=conv_id,
        importance=msg.importance,
        body=_body_text(msg.body),
        body_content_type=msg.body.content_type.lower(),
        attachments=[_attachment_info(att) for att in msg.attachments],
    )


# ---------------------------------------------------------------------------
# search_emails
# ---------------------------------------------------------------------------


_KQL_OPERATORS = frozenset({"AND", "OR", "NOT"})
_MAIL_KQL_PROPERTIES = frozenset(
    {
        "attachment",
        "bcc",
        "body",
        "cc",
        "from",
        "hasattachment",
        "importance",
        "isread",
        "kind",
        "participants",
        "received",
        "sent",
        "subject",
        "to",
    }
)
_SUBJECT_SEARCH_PATTERN = re.compile(
    r'^subject\s*:\s*"(?P<subject>(?:\\.|[^"\\])*)"$',
    flags=re.IGNORECASE,
)


def _normalize_email_search_query(query: str) -> str:
    """Preserve structured KQL while quoting punctuation-bearing operands."""
    query = query.strip()
    is_structured = bool(
        '"' in query
        or "(" in query
        or ")" in query
        or re.search(r"\b(?:AND|OR|NOT)\b", query, flags=re.IGNORECASE)
        or re.search(r"\b[A-Za-z][A-Za-z0-9]*:", query)
    )
    if not is_structured:
        return f'"{query}"'

    parts = re.split(r'("(?:\\.|[^"\\])*"|[()\s]+)', query)
    normalized: list[str] = []
    for part in parts:
        if (
            not part
            or part.isspace()
            or part in {"(", ")"}
            or (part.startswith('"') and part.endswith('"'))
            or part.upper() in _KQL_OPERATORS
        ):
            normalized.append(part)
            continue

        property_name, separator, value = part.partition(":")
        if separator and property_name.casefold() in _MAIL_KQL_PROPERTIES:
            if value and any(character in value for character in ".@"):
                part = f'{property_name}:"{value}"'
        elif any(character in part for character in ".@"):
            part = f'"{part}"'
        normalized.append(part)

    return "".join(normalized)


def _subject_search_filter(query: str) -> str | None:
    """Convert a standalone subject phrase to Graph's supported OData filter."""
    match = _SUBJECT_SEARCH_PATTERN.fullmatch(query.strip())
    if not match:
        return None
    subject = match.group("subject").replace(r'\"', '"').replace("'", "''")
    return f"contains(subject, '{subject}')"


async def search_emails(
    params: SearchEmailsInput,
) -> SearchEmailsResponse:
    """
    Search messages using Graph KQL $search syntax.

    Note: Graph $search and $filter cannot be combined in the same request.
    The Graph API caps $search results at 25 regardless of the value requested.

    Args:
        query: KQL search string, e.g. 'from:alice@example.com' or 'project update'.
        max_results: Maximum number of results (1-25). Values above 25 are
            silently capped by the Graph API. Defaults to 10.
        folder: Optional well-known folder name or folder ID to restrict the search.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured search results.
    """
    g = get_graph(params.profile)
    query_params: dict[str, Any] = {
        "$top": min(params.max_results, 25),
        "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments",
    }
    subject_filter = _subject_search_filter(params.query)
    if subject_filter:
        query_params["$filter"] = subject_filter
    else:
        # Preserve structured KQL instead of wrapping it in another pair of
        # quotes, and quote punctuation-bearing operands Graph rejects bare.
        query_params["$search"] = _normalize_email_search_query(params.query)

    if params.folder:
        path = f"/me/mailFolders/{params.folder}/messages"
    else:
        path = "/me/messages"

    result = await g.get(path, params=query_params)
    messages = parse_graph_collection(result, GraphMessage)

    return SearchEmailsResponse(
        query=params.query,
        folder=params.folder,
        count=len(messages),
        messages=[_message_summary(msg) for msg in messages],
    )


# ---------------------------------------------------------------------------
# filter_emails
# ---------------------------------------------------------------------------


async def filter_emails(
    params: FilterEmailsInput,
) -> ListEmailsResponse:
    """
    Find emails matching specific criteria using OData $filter.

    Unlike search_emails (which is limited to 25 results), this tool
    supports up to 100 results per page with full pagination — ideal for
    finding all emails from a sender, within a date range, or matching
    a subject.

    All filter parameters are combined with AND logic. Omitted parameters
    are not filtered on.

    Args:
        from_address: Filter by sender email address (exact match).
        to_address: Filter by recipient email address (exact match).
        subject_contains: Filter by subject containing this text (case-insensitive).
        received_after: Only messages received on or after this date.
            ISO 8601 format: '2026-01-01' or '2026-01-01T00:00:00Z'.
        received_before: Only messages received before this date.
            ISO 8601 format: '2026-03-31' or '2026-03-31T23:59:59Z'.
        has_attachments: When True, only messages with attachments.
            When False, only messages without.
        importance: Filter by importance level: 'low', 'normal', or 'high'.
        folder: Well-known folder name or folder ID. Defaults to 'inbox'.
        max_results: Maximum number of messages to return (1-100). Defaults to 50.
        sort_order: 'newest' (default) or 'oldest' first.
        skip_token: Opaque pagination cursor returned as next_page_token from a
                    previous call. Omit for the first page.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured message summaries with pagination metadata.
    """
    g = get_graph(params.profile)
    order = "receivedDateTime asc" if params.sort_order == "oldest" else "receivedDateTime desc"
    query: dict[str, Any] = {
        "$top": min(max(1, params.max_results), 100),
        "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments,importance",
        "$orderby": order,
    }

    # Graph requires the $orderby property to appear first in $filter.
    date_clauses: list[str] = []
    if params.received_after:
        ts = params.received_after if "T" in params.received_after else f"{params.received_after}T00:00:00Z"
        date_clauses.append(f"receivedDateTime ge {ts}")
    if params.received_before:
        ts = params.received_before if "T" in params.received_before else f"{params.received_before}T23:59:59Z"
        date_clauses.append(f"receivedDateTime lt {ts}")

    clauses: list[str] = []
    if params.from_address:
        safe = params.from_address.replace("'", "''")
        clauses.append(f"from/emailAddress/address eq '{safe}'")
    if params.to_address:
        safe = params.to_address.replace("'", "''")
        clauses.append(f"toRecipients/any(r:r/emailAddress/address eq '{safe}')")
    if params.subject_contains:
        safe = params.subject_contains.replace("'", "''")
        clauses.append(f"contains(subject, '{safe}')")
    if params.has_attachments is not None:
        clauses.append(f"hasAttachments eq {str(params.has_attachments).lower()}")
    if params.importance:
        clauses.append(f"importance eq '{params.importance}'")

    if clauses or date_clauses:
        if not date_clauses:
            date_clauses.append("receivedDateTime ge 1900-01-01T00:00:00Z")
        query["$filter"] = " and ".join(date_clauses + clauses)

    if params.skip_token is not None:
        query["$skiptoken"] = params.skip_token

    result = await g.get(f"/me/mailFolders/{params.folder}/messages", params=query)

    messages = parse_graph_collection(result, GraphMessage)
    next_link = result.get("@odata.nextLink")

    from urllib.parse import parse_qs, urlparse
    next_page_token: str | None = None
    if next_link:
        qs = parse_qs(urlparse(next_link).query)
        next_page_token = qs.get("$skiptoken", [None])[0]

    return ListEmailsResponse(
        folder=params.folder,
        count=len(messages),
        messages=[_message_summary(msg) for msg in messages],
        next_page_token=next_page_token,
        has_more=(next_page_token is not None),
    )


# ---------------------------------------------------------------------------
# send_email
# ---------------------------------------------------------------------------


async def send_email(
    params: SendEmailInput,
    ctx: Context | None = None,
) -> SendEmailResponse:
    """
    Send a new email message.

    Args:
        to: Recipient address(es). Comma-separated string or list.
        subject: Email subject line.
        body: Message body content.
        cc: Optional CC address(es). Comma-separated string or list.
        bcc: Optional BCC address(es). Comma-separated string or list.
        body_type: 'text' or 'html'. Defaults to 'text'.
        save_to_sent: When True (default), save a copy in Sent Items.
        reply_to: Optional reply-to address(es).
        profile: Microsoft 365 profile to use. Omit to use the default profile.
        confirm: When True, prompt the user to confirm before sending. Defaults to False.

    Returns:
        Structured send confirmation.
    """
    if params.confirm and ctx:
        to_display = params.to if isinstance(params.to, str) else ", ".join(params.to)
        preview = (
            f"To: {to_display}\n"
            f"Subject: {params.subject}\n\n"
            f"{params.body[:200]}{'...' if len(params.body) > 200 else ''}"
        )
        result = await ctx.elicit(
            f"Send this email?\n\n{preview}",
            response_type=_Confirmation,
        )
        if result.action != "accept" or not result.data.confirmed:
            return SendEmailResponse(success=False, action="send_email", error="Cancelled by user.")

    g = get_graph(params.profile)
    message: dict = {
        "subject": params.subject,
        "body": {
            "contentType": "HTML" if params.body_type.lower() == "html" else "Text",
            "content": params.body,
        },
        "toRecipients": parse_recipients(params.to),
    }

    if params.cc:
        message["ccRecipients"] = parse_recipients(params.cc)
    if params.bcc:
        message["bccRecipients"] = parse_recipients(params.bcc)
    if params.reply_to:
        message["replyTo"] = parse_recipients(params.reply_to)

    payload = {
        "message": message,
        "saveToSentItems": params.save_to_sent,
    }

    await g.post("/me/sendMail", json=payload)

    return SendEmailResponse(
        success=True,
        action="send_email",
        to=[addr.get("emailAddress", {}).get("address", "") for addr in message["toRecipients"]],
        cc=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("ccRecipients", [])],
        bcc=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("bccRecipients", [])],
        reply_to=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("replyTo", [])],
        subject=params.subject,
        body_type=params.body_type,
        saved_to_sent_items=params.save_to_sent,
    )


# ---------------------------------------------------------------------------
# reply_email
# ---------------------------------------------------------------------------


async def reply_email(
    params: ReplyEmailInput,
) -> ReplyEmailResponse:
    """
    Reply to an existing email message.

    Args:
        message_id: The Graph message ID to reply to.
        body: Reply body text or HTML.
        reply_all: When True, reply to all recipients. Defaults to False.
        body_type: 'text' or 'html'. Defaults to 'text'.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured reply confirmation.
    """
    g = get_graph(params.profile)
    endpoint = "replyAll" if params.reply_all else "reply"
    if params.body_type.lower() == "html":
        payload = {
            "message": {
                "body": {
                    "contentType": "HTML",
                    "content": params.body,
                }
            }
        }
    else:
        payload = {
            "message": {},
            "comment": params.body,
        }

    await g.post(f"/me/messages/{params.message_id}/{endpoint}", json=payload)

    return ReplyEmailResponse(
        success=True,
        action="reply_all" if params.reply_all else "reply",
        message_id=params.message_id,
        body_type=params.body_type,
    )


# ---------------------------------------------------------------------------
# forward_email
# ---------------------------------------------------------------------------


async def forward_email(
    params: ForwardEmailInput,
) -> ForwardEmailResponse:
    """
    Forward an email message to one or more recipients.

    Args:
        message_id: The Graph message ID to forward.
        to: Recipient address(es). Comma-separated string or list.
        comment: Optional comment to prepend to the forwarded message.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured forward confirmation.
    """
    g = get_graph(params.profile)
    payload: dict = {
        "toRecipients": parse_recipients(params.to),
        "comment": params.comment or "",
    }

    await g.post(f"/me/messages/{params.message_id}/forward", json=payload)

    return ForwardEmailResponse(
        success=True,
        action="forward",
        message_id=params.message_id,
        to=[addr.get("emailAddress", {}).get("address", "") for addr in payload["toRecipients"]],
        comment=params.comment or "",
    )


# ---------------------------------------------------------------------------
# mark_as_read / mark_as_unread
# ---------------------------------------------------------------------------


async def mark_as_read(params: MarkAsReadInput) -> MarkEmailReadResponse:
    """
    Mark a message as read.

    Args:
        message_id: The Graph message ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured update confirmation.
    """
    g = get_graph(params.profile)
    await g.patch(f"/me/messages/{params.message_id}", json={"isRead": True})
    return MarkEmailReadResponse(success=True, action="mark_as_read", message_id=params.message_id, is_read=True)


async def mark_as_unread(params: MarkAsUnreadInput) -> MarkEmailReadResponse:
    """
    Mark a message as unread.

    Args:
        message_id: The Graph message ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured update confirmation.
    """
    g = get_graph(params.profile)
    await g.patch(f"/me/messages/{params.message_id}", json={"isRead": False})
    return MarkEmailReadResponse(success=True, action="mark_as_unread", message_id=params.message_id, is_read=False)


# ---------------------------------------------------------------------------
# move_email
# ---------------------------------------------------------------------------


async def move_email(params: MoveEmailInput) -> MoveEmailResponse:
    """
    Move a message to a different mail folder.

    Args:
        message_id: The Graph message ID to move.
        destination_folder: Target folder — well-known name (e.g. 'archive',
            'inbox', 'junkemail', 'deleteditems') or opaque folder ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured move confirmation.
    """
    g = get_graph(params.profile)
    result = await g.post(
        f"/me/messages/{params.message_id}/move",
        json={"destinationId": params.destination_folder},
    )
    moved_message = _graph_message_from_payload(result if isinstance(result, dict) else None)
    new_id = moved_message.id if moved_message and moved_message.id else params.message_id
    return MoveEmailResponse(
        success=True,
        action="move",
        message_id=params.message_id,
        new_message_id=new_id,
        destination_folder=params.destination_folder,
    )


# ---------------------------------------------------------------------------
# trash_email
# ---------------------------------------------------------------------------


async def trash_email(params: TrashEmailInput) -> TrashEmailResponse:
    """
    Move a message to the Deleted Items folder (soft delete / recoverable).

    To permanently delete without recovery, use delete_email instead.

    Args:
        message_id: The Graph message ID to trash.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured trash confirmation.
    """
    g = get_graph(params.profile)
    result = await g.post(
        f"/me/messages/{params.message_id}/move",
        json={"destinationId": "deleteditems"},
    )
    moved_message = _graph_message_from_payload(result if isinstance(result, dict) else None)
    new_id = moved_message.id if moved_message and moved_message.id else params.message_id
    return TrashEmailResponse(
        success=True,
        action="trash",
        message_id=params.message_id,
        new_message_id=new_id,
        destination_folder="deleteditems",
        soft_delete=True,
        profile=params.profile,
    )


# ---------------------------------------------------------------------------
# delete_email
# ---------------------------------------------------------------------------


async def delete_email(
    params: DeleteEmailInput,
    ctx: Context | None = None,
) -> DeleteEmailResponse:
    """
    Permanently delete a message from the mailbox. This action is IRREVERSIBLE.

    The message will be hard-deleted and cannot be recovered from Deleted Items.
    For a recoverable soft delete, use trash_email instead.

    Args:
        message_id: The Graph message ID to permanently delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
        confirm: When True, prompt the user to confirm before deleting. Defaults to False.

    Returns:
        Structured delete confirmation.
    """
    if params.confirm:
        if ctx is None:
            return DeleteEmailResponse(
                success=False,
                action="permanent_delete",
                message_id=params.message_id,
                error=(
                    "confirm=True requires an MCP host that supports elicitation. "
                    "Set confirm=False to bypass the prompt or upgrade your client."
                ),
                irreversible=True,
            )
        result = await ctx.elicit(
            f"Permanently delete this email? This action is IRREVERSIBLE.\n\nMessage ID: {params.message_id}",
            response_type=_Confirmation,
        )
        if result.action != "accept" or not result.data.confirmed:
            return DeleteEmailResponse(success=False, action="permanent_delete", message_id=params.message_id, error="Cancelled by user.", irreversible=True)

    g = get_graph(params.profile)
    await g.post(f"/me/messages/{params.message_id}/permanentDelete")
    return DeleteEmailResponse(
        success=True,
        action="permanent_delete",
        message_id=params.message_id,
        irreversible=True,
    )


# ---------------------------------------------------------------------------
# Batch helper (Graph $batch endpoint, max 20 requests per batch)
# ---------------------------------------------------------------------------


def _build_batch_requests(
    message_ids: list[str],
    operation: _BatchOperation,
) -> _BatchRequestPlan:
    """
    Build Graph batch request entries with synthetic IDs.

    Returns a typed plan that preserves the mapping between synthetic batch IDs
    and the original message IDs.
    """
    plan = _BatchRequestPlan()
    requests: list[_BatchRequestEntry] = []
    for idx, mid in enumerate(message_ids):
        batch_id = str(idx + 1)
        plan.message_ids_by_batch_id[batch_id] = mid
        requests.append(
            _BatchRequestEntry(
                id=batch_id,
                method=operation.method,
                url=operation.url_template.format(mid=mid),
                body=operation.body,
                headers={"Content-Type": "application/json"} if operation.body is not None else None,
            )
        )
    plan.requests = requests
    return plan


def _parse_batch_error(resp: _BatchResponseEntry) -> tuple[str | None, str]:
    """Extract error code and message from a batch response item."""
    status = resp.status
    if not isinstance(resp.body, dict):
        return None, f"HTTP {status}"
    body = _BatchResponseBody.model_validate(resp.body)
    if body.error is None:
        return None, f"HTTP {status}"
    return body.error.code, body.error.message or f"HTTP {status}"


async def _execute_batch(
    g: Any,
    plan: _BatchRequestPlan,
    success_parser: Callable[[dict[str, Any] | None], GraphMessage | None] | None = None,
) -> _BatchExecutionResult:
    """
    Send requests via GraphClient.batch() (which handles $batch chunking).

    Returns typed success and failure details keyed by the original message IDs.
    """
    result = _BatchExecutionResult()

    # Build a set of all expected batch IDs so we can detect missing responses.
    expected_ids = {request.id for request in plan.requests}

    try:
        responses = await g.batch([request.model_dump(exclude_none=True) for request in plan.requests])
    except Exception as exc:
        # The entire batch call failed — record a failure for every request.
        for req in plan.requests:
            mid = plan.message_ids_by_batch_id.get(req.id, req.id)
            result.failures.append(BulkEmailFailure(
                message_id=mid, status=0, error=str(exc),
            ))
        return result

    seen_ids: set[str] = set()
    for raw_response in responses:
        resp = _BatchResponseEntry.model_validate(raw_response)
        batch_id = resp.id
        seen_ids.add(batch_id)
        mid = plan.message_ids_by_batch_id.get(batch_id, batch_id)
        status = resp.status

        if 200 <= status < 300:
            result.succeeded_message_ids.append(mid)
            if success_parser is not None:
                parsed_message = success_parser(resp.body)
                if parsed_message is not None:
                    result.response_messages[mid] = parsed_message
        else:
            code, error_msg = _parse_batch_error(resp)
            result.failures.append(BulkEmailFailure(
                message_id=mid, status=status, code=code, error=error_msg,
            ))

    # Detect missing responses (Graph returned fewer items than we sent).
    missing = expected_ids - seen_ids
    for batch_id in missing:
        mid = plan.message_ids_by_batch_id.get(batch_id, batch_id)
        result.failures.append(BulkEmailFailure(
            message_id=mid, status=0, error="No response from batch",
        ))

    return result


# ---------------------------------------------------------------------------
# bulk_move_emails
# ---------------------------------------------------------------------------


async def _collect_folder_message_ids(g: Any, folder: str) -> list[str]:
    """Fetch all message IDs from a mail folder, paginating as needed."""
    ids: list[str] = []
    path = f"/me/mailFolders/{folder}/messages"
    params: dict[str, Any] = {"$top": 100, "$select": "id"}
    while path:
        result = await g.get(path, params=params)
        for msg in result.get("value", []):
            parsed = GraphMessage.model_validate(msg)
            if parsed.id:
                ids.append(parsed.id)
        next_link = result.get("@odata.nextLink")
        if next_link:
            # nextLink is a full URL; extract relative path + query
            import re as _re
            m = _re.search(r"v1\.0(/.+)", next_link)
            path = m.group(1) if m else ""
            params = {}  # params are embedded in nextLink
        else:
            path = ""
    return ids


async def _bulk_confirmation_error(
    *,
    confirm: bool,
    ctx: Context | None,
    message_count: int,
    folder: str | None,
    permanent: bool,
) -> str | None:
    """Return an error unless a required bulk-operation confirmation succeeds."""
    folder_mode = folder is not None
    if folder_mode and not confirm:
        action = "permanent deletion" if permanent else "trashing"
        return (
            f"Folder-wide {action} requires confirm=True so the message count "
            "can be shown to the user before continuing."
        )
    if not confirm:
        return None
    if ctx is None:
        return (
            "confirm=True requires an MCP host that supports elicitation. "
            "The bulk operation was not performed."
        )

    scope = f" from '{folder}'" if folder else ""
    if permanent:
        prompt = (
            f"Permanently delete {message_count} messages{scope}? "
            "This action is IRREVERSIBLE."
        )
    else:
        prompt = f"Move {message_count} messages{scope} to Deleted Items?"
    result = await ctx.elicit(prompt, response_type=_Confirmation)
    if result.action != "accept" or not result.data.confirmed:
        return "Cancelled by user."
    return None


async def bulk_move_emails(
    params: BulkMoveEmailsInput,
) -> BulkMoveEmailsResponse:
    """
    Move multiple messages to a destination folder in one operation.

    Uses the Graph batch API for efficiency (up to 20 per round-trip).

    Two modes:
    1. Pass message_ids explicitly.
    2. Pass source_folder to move ALL messages from that folder (e.g. 'junkemail').

    Args:
        message_ids: List of Graph message IDs to move. Optional if source_folder is set.
        destination_folder: Target folder — well-known name (e.g. 'archive',
            'inbox', 'junkemail', 'deleteditems') or opaque folder ID.
        source_folder: Move all messages from this folder instead of specifying IDs.
            Well-known names: inbox, sentitems, drafts, deleteditems, junkemail, archive.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Summary with success/failure counts, new message IDs, and failure details.
    """
    if not params.destination_folder:
        return BulkMoveEmailsResponse(success=False, action="bulk_move", error="destination_folder is required.")

    g = get_graph(params.profile)
    message_ids = params.message_ids

    if params.source_folder and not message_ids:
        message_ids = await _collect_folder_message_ids(g, params.source_folder)

    if not message_ids:
        return BulkMoveEmailsResponse(success=True, action="bulk_move", destination_folder=params.destination_folder, total=0, succeeded=0, failed=0)

    plan = _build_batch_requests(
        message_ids,
        _BatchOperation(
            method="POST",
            url_template="/me/messages/{mid}/move",
            body={"destinationId": params.destination_folder},
        ),
    )

    batch_result = await _execute_batch(g, plan, success_parser=_graph_message_from_payload)

    moved = [
        BulkMovedEmail(
            source_message_id=mid,
            new_message_id=message.id,
        )
        for mid, message in batch_result.response_messages.items()
    ]

    return BulkMoveEmailsResponse(
        success=len(batch_result.failures) == 0,
        action="bulk_move",
        destination_folder=params.destination_folder,
        total=len(message_ids),
        succeeded=len(batch_result.succeeded_message_ids),
        failed=len(batch_result.failures),
        moved=moved,
        failures=batch_result.failures,
    )


# ---------------------------------------------------------------------------
# bulk_trash_emails
# ---------------------------------------------------------------------------


async def bulk_trash_emails(
    params: BulkTrashEmailsInput,
    ctx: Context | None = None,
) -> BulkTrashEmailsResponse:
    """
    Move multiple messages to Deleted Items (soft delete / recoverable).

    Uses the Graph batch API for efficiency (up to 20 per round-trip).
    For permanent deletion, use bulk_delete_emails instead.

    Two modes:
    1. Pass message_ids explicitly.
    2. Pass folder to trash ALL messages from that folder (e.g. 'junkemail').

    Args:
        message_ids: List of Graph message IDs to trash. Optional if folder is set.
        folder: Trash all messages from this folder instead of specifying IDs.
            Well-known names: inbox, sentitems, drafts, junkemail, archive.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
        confirm: Required for folder-wide operations. Prompts with the target
            count before moving any messages.

    Returns:
        Summary with success/failure counts, new message IDs, and failure details.
    """
    g = get_graph(params.profile)
    message_ids = params.message_ids

    if params.folder and not message_ids:
        message_ids = await _collect_folder_message_ids(g, params.folder)

    if not message_ids:
        return BulkTrashEmailsResponse(success=True, action="bulk_trash", total=0, succeeded=0, failed=0)

    confirmation_error = await _bulk_confirmation_error(
        confirm=params.confirm,
        ctx=ctx,
        message_count=len(message_ids),
        folder=params.folder if params.folder and not params.message_ids else None,
        permanent=False,
    )
    if confirmation_error:
        return BulkTrashEmailsResponse(
            success=False,
            action="bulk_trash",
            total=len(message_ids),
            succeeded=0,
            failed=0,
            error=confirmation_error,
        )

    plan = _build_batch_requests(
        message_ids,
        _BatchOperation(
            method="POST",
            url_template="/me/messages/{mid}/move",
            body={"destinationId": "deleteditems"},
        ),
    )

    batch_result = await _execute_batch(g, plan, success_parser=_graph_message_from_payload)

    moved = [
        BulkMovedEmail(
            source_message_id=mid,
            new_message_id=message.id,
        )
        for mid, message in batch_result.response_messages.items()
    ]

    return BulkTrashEmailsResponse(
        success=len(batch_result.failures) == 0,
        action="bulk_trash",
        total=len(message_ids),
        succeeded=len(batch_result.succeeded_message_ids),
        failed=len(batch_result.failures),
        moved=moved,
        failures=batch_result.failures,
    )


# ---------------------------------------------------------------------------
# bulk_delete_emails
# ---------------------------------------------------------------------------


async def bulk_delete_emails(
    params: BulkDeleteEmailsInput,
    ctx: Context | None = None,
) -> BulkDeleteEmailsResponse:
    """
    Permanently delete multiple messages from the mailbox. This action is IRREVERSIBLE.

    Uses the Graph batch API for efficiency (up to 20 per round-trip).
    Messages will be hard-deleted and cannot be recovered from Deleted Items.
    For a recoverable soft delete, use bulk_trash_emails instead.

    Two modes:
    1. Pass message_ids explicitly.
    2. Pass folder to permanently delete ALL messages from that folder.

    Args:
        message_ids: List of Graph message IDs to permanently delete.
            Optional if folder is set.
        folder: Permanently delete all messages from this folder instead of
            specifying IDs. Well-known names: inbox, sentitems, drafts,
            deleteditems, junkemail, archive.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
        confirm: Required for folder-wide operations. Prompts with the target
            count and irreversible-action warning before deleting anything.

    Returns:
        Summary with success/failure counts and failure details.
    """
    g = get_graph(params.profile)
    message_ids = params.message_ids

    if params.folder and not message_ids:
        message_ids = await _collect_folder_message_ids(g, params.folder)

    if not message_ids:
        return BulkDeleteEmailsResponse(success=True, action="bulk_permanent_delete", total=0, succeeded=0, failed=0, irreversible=True)

    confirmation_error = await _bulk_confirmation_error(
        confirm=params.confirm,
        ctx=ctx,
        message_count=len(message_ids),
        folder=params.folder if params.folder and not params.message_ids else None,
        permanent=True,
    )
    if confirmation_error:
        return BulkDeleteEmailsResponse(
            success=False,
            action="bulk_permanent_delete",
            total=len(message_ids),
            succeeded=0,
            failed=0,
            irreversible=True,
            error=confirmation_error,
        )

    plan = _build_batch_requests(
        message_ids,
        _BatchOperation(
            method="POST",
            url_template="/me/messages/{mid}/permanentDelete",
        ),
    )

    batch_result = await _execute_batch(g, plan)

    return BulkDeleteEmailsResponse(
        success=len(batch_result.failures) == 0,
        action="bulk_permanent_delete",
        total=len(message_ids),
        succeeded=len(batch_result.succeeded_message_ids),
        failed=len(batch_result.failures),
        irreversible=True,
        failures=batch_result.failures,
    )


def register(server) -> None:
    """Register all mail tools with the given FastMCP server instance."""
    register_tool(server, list_emails, annotations=READ_ONLY_TOOL)
    register_tool(server, read_email, annotations=READ_ONLY_TOOL)
    register_tool(server, search_emails, annotations=READ_ONLY_TOOL)
    register_tool(server, filter_emails, annotations=READ_ONLY_TOOL)
    register_tool(server, send_email, annotations=WRITE_TOOL)
    register_tool(server, reply_email, annotations=WRITE_TOOL)
    register_tool(server, forward_email, annotations=WRITE_TOOL)
    register_tool(server, mark_as_read, annotations=IDEMPOTENT_WRITE_TOOL)
    register_tool(server, mark_as_unread, annotations=IDEMPOTENT_WRITE_TOOL)
    register_tool(server, move_email, annotations=WRITE_TOOL)
    register_tool(server, trash_email, annotations=WRITE_TOOL)
    register_tool(server, bulk_move_emails, annotations=WRITE_TOOL)
    register_tool(server, bulk_trash_emails, annotations=WRITE_TOOL)
    if not is_deletion_disabled():
        register_tool(server, delete_email, annotations=DESTRUCTIVE_TOOL)
        register_tool(server, bulk_delete_emails, annotations=DESTRUCTIVE_TOOL)
