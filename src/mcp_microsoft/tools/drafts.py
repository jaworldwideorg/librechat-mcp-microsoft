"""
Draft management tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - create_draft
  - create_reply_draft
  - list_drafts
  - get_draft
  - update_draft
  - send_draft
"""

from __future__ import annotations

import html
import re
from typing import Literal

import httpx
from pydantic import Field

from mcp_microsoft.common.mail_utils import format_mail_datetime, parse_recipients, recipient_values
from mcp_microsoft.common.text import strip_html
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.graph_types import GraphMessage, parse_graph_collection
from mcp_microsoft.models import (
    CreateDraftResponse,
    CreateReplyDraftResponse,
    DraftDetailResponse,
    DraftSummary,
    ListDraftsResponse,
    SendDraftResponse,
    UpdateDraftResponse,
)
from mcp_microsoft.graph import get_graph

# ---------------------------------------------------------------------------
# create_draft
# ---------------------------------------------------------------------------

BodyType = Literal["text", "html"]


class CreateDraftInput(ToolRequestModel):
    to: str | list[str]
    subject: str
    body: str
    cc: str | list[str] | None = None
    body_type: BodyType = "text"
    profile: str | None = None


class CreateReplyDraftInput(ToolRequestModel):
    message_id: str = Field(
        min_length=1,
        max_length=4096,
        description="Graph message ID to reply to.",
    )
    body: str | None = Field(
        default=None,
        description=(
            "Optional initial reply body, inserted above the quoted conversation "
            "history. Omit to create an empty reply draft."
        ),
    )
    reply_all: bool = Field(
        default=False,
        description="Create a reply-all draft when true; otherwise reply only to the sender.",
    )
    body_type: BodyType = Field(
        default="text",
        description="Format of body when supplied: text or html.",
    )
    profile: str | None = None


class ListDraftsInput(ToolRequestModel):
    max_results: int = 10
    profile: str | None = None


class GetDraftInput(ToolRequestModel):
    draft_id: str
    profile: str | None = None


class UpdateDraftInput(ToolRequestModel):
    draft_id: str
    subject: str | None = None
    body: str | None = None
    to: str | list[str] | None = None
    cc: str | list[str] | None = None
    body_type: BodyType = "text"
    preserve_history: bool = Field(
        default=False,
        description=(
            "When true and body is provided, prepend the new body to the draft's "
            "existing body so quoted reply history is retained. When false, body "
            "replaces the complete draft body."
        ),
    )
    profile: str | None = None


class SendDraftInput(ToolRequestModel):
    draft_id: str
    profile: str | None = None


async def create_draft(
    params: CreateDraftInput,
) -> CreateDraftResponse:
    """
    Create a new unsent email draft message in the Drafts folder.

    Use this when you want to compose an email now and send it later with
    send_draft. The result includes the draft message ID, which can also be
    used with get_draft or update_draft.

    Args:
        to: Recipient address(es). Comma-separated string or list.
        subject: Draft subject line.
        body: Draft body content.
        cc: Optional CC address(es). Comma-separated string or list.
        body_type: 'text' or 'html'. Defaults to 'text'.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured draft creation confirmation.
    """
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

    created = GraphMessage.model_validate(await g.post("/me/messages", json=message) or {})
    return CreateDraftResponse(
        success=True,
        action="create_draft",
        draft_id=created.id or "unknown",
        to=[addr.get("emailAddress", {}).get("address", "") for addr in message["toRecipients"]],
        cc=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("ccRecipients", [])],
        subject=params.subject,
        body_type=params.body_type,
    )


# ---------------------------------------------------------------------------
# create_reply_draft
# ---------------------------------------------------------------------------


async def create_reply_draft(
    params: CreateReplyDraftInput,
) -> CreateReplyDraftResponse:
    """
    Create an unsent reply draft linked to an existing email thread.

    The draft is saved in the Drafts folder and is not sent. Use get_draft or
    update_draft to review or edit it, then send_draft when ready. Requires the
    delegated Microsoft Graph ``Mail.ReadWrite`` permission.

    Args:
        message_id: Graph message ID to reply to.
        body: Optional initial reply body, inserted above the quoted conversation
            history. Omit to create an empty reply draft.
        reply_all: Create a reply-all draft when True. Defaults to False.
        body_type: 'text' or 'html'. Used only when body is supplied.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured confirmation containing the new draft ID.
    """
    g = get_graph(params.profile)
    endpoint = "createReplyAll" if params.reply_all else "createReply"
    payload: dict = {}
    if params.body is not None:
        # Graph composes ``comment`` above the quoted original. Supplying
        # ``message.body`` instead replaces the generated reply body and drops
        # the conversation history.
        payload["comment"] = params.body

    raw_created = await g.post(
        f"/me/messages/{params.message_id}/{endpoint}",
        json=payload,
    )
    created = GraphMessage.model_validate(raw_created or {})
    if not created.id or not created.is_draft:
        return CreateReplyDraftResponse(
            success=False,
            action="create_reply_draft",
            error="Microsoft Graph did not return a valid reply draft.",
            draft_id=created.id,
            original_message_id=params.message_id,
            reply_all=params.reply_all,
            body_type=params.body_type,
        )

    return CreateReplyDraftResponse(
        success=True,
        action="create_reply_draft",
        draft_id=created.id,
        original_message_id=params.message_id,
        reply_all=params.reply_all,
        body_type=params.body_type,
    )


# ---------------------------------------------------------------------------
# list_drafts
# ---------------------------------------------------------------------------


async def list_drafts(params: ListDraftsInput) -> ListDraftsResponse:
    """
    List unsent email drafts from the Drafts folder.

    Args:
        max_results: Maximum number of drafts to return (1-100). Defaults to 10.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured draft summaries.
    """
    g = get_graph(params.profile)
    query: dict = {
        "$top": params.max_results,
        "$select": "id,subject,toRecipients,lastModifiedDateTime,bodyPreview",
        "$orderby": "lastModifiedDateTime desc",
    }

    result = await g.get("/me/mailFolders/drafts/messages", params=query)
    drafts = parse_graph_collection(result, GraphMessage)

    items: list[DraftSummary] = []
    for draft in drafts:
        items.append(
            DraftSummary(
                id=draft.id,
                subject=draft.subject or "(no subject)",
                to=recipient_values(draft.to_recipients),
                last_modified_at=draft.last_modified_date_time,
                last_modified_at_display=format_mail_datetime(draft.last_modified_date_time),
                preview=(draft.body_preview or "").replace("\n", " ")[:100],
            )
        )

    return ListDraftsResponse(count=len(items), drafts=items)


# ---------------------------------------------------------------------------
# get_draft
# ---------------------------------------------------------------------------


async def get_draft(params: GetDraftInput) -> DraftDetailResponse:
    """
    Get a draft email by ID, including its full body, recipients, and metadata.

    Use this after list_drafts or create_draft when you need the complete draft
    content instead of just a preview.

    Args:
        draft_id: The Graph message ID of the draft.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured draft details, or a safe structured failure when the draft
        no longer exists. Other Graph failures are propagated normally.
    """
    g = get_graph(params.profile)
    query = {
        "$select": (
            "id,subject,from,toRecipients,ccRecipients,bccRecipients,"
            "lastModifiedDateTime,body,bodyPreview,isDraft"
        ),
    }

    try:
        raw_draft = await g.get(f"/me/messages/{params.draft_id}", params=query)
    except httpx.HTTPStatusError as exc:
        if exc.response.status_code != 404:
            raise
        return DraftDetailResponse(
            success=False,
            action="get_draft",
            error=(
                "Draft not found. It may have been deleted or sent, or the "
                "draft ID may no longer be valid."
            ),
            is_draft=False,
        )

    draft = GraphMessage.model_validate(raw_draft)

    subject = draft.subject or "(no subject)"
    modified = format_mail_datetime(draft.last_modified_date_time)

    content_type = (draft.body.content_type or "text").lower()
    raw_body = draft.body.content or ""

    if content_type == "html":
        body_text = strip_html(raw_body)
    else:
        body_text = raw_body

    return DraftDetailResponse(
        success=True,
        action="get_draft",
        id=params.draft_id,
        subject=subject,
        to=recipient_values(draft.to_recipients),
        cc=recipient_values(draft.cc_recipients),
        bcc=recipient_values(draft.bcc_recipients),
        last_modified_at=draft.last_modified_date_time,
        last_modified_at_display=modified,
        body=body_text,
        body_content_type=content_type,
        is_draft=draft.is_draft,
    )


# ---------------------------------------------------------------------------
# update_draft
# ---------------------------------------------------------------------------


_BODY_OPEN_TAG = re.compile(r"(<body(?:\s[^>]*)?>)", re.IGNORECASE)


def _text_to_html_fragment(value: str) -> str:
    """Convert caller-supplied plain text to a safe HTML fragment."""
    escaped = html.escape(value).replace("\r\n", "\n").replace("\r", "\n")
    return f"<div>{escaped.replace(chr(10), '<br>')}</div>"


def _prepend_draft_body(
    *,
    new_body: str,
    new_body_type: BodyType,
    existing_body: str,
    existing_body_type: str,
) -> str:
    """Prepend new content while keeping the complete existing draft body."""
    new_fragment = (
        new_body if new_body_type == "html" else _text_to_html_fragment(new_body)
    )
    existing_html = (
        existing_body
        if existing_body_type.lower() == "html"
        else _text_to_html_fragment(existing_body)
    )
    insertion = f"{new_fragment}<br><br>"
    if match := _BODY_OPEN_TAG.search(existing_html):
        return (
            existing_html[: match.end()]
            + insertion
            + existing_html[match.end() :]
        )
    return insertion + existing_html


async def update_draft(
    params: UpdateDraftInput,
) -> UpdateDraftResponse:
    """
    Update an existing email draft. Only provided fields are changed.

    By default, a supplied body replaces the complete draft body. Set
    preserve_history to prepend the new body while retaining the existing
    content, including quoted reply history. This tool does not send the
    draft. Requires the delegated Microsoft Graph ``Mail.ReadWrite``
    permission.

    Args:
        draft_id: The Graph message ID of the draft to update.
        subject: Replace subject line (omit to leave unchanged).
        body: Body content to replace or prepend (omit to leave unchanged).
        to: Replace recipient list (omit to leave unchanged).
        cc: Replace CC list (omit to leave unchanged).
        body_type: 'text' or 'html'. Only used when body is provided. Defaults to 'text'.
        preserve_history: When True, prepend body to the existing complete body
            so quoted reply history is retained. Defaults to False, which
            preserves the existing full-body replacement behavior.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured update confirmation.
    """
    g = get_graph(params.profile)
    patch: dict = {}

    if params.subject is not None:
        patch["subject"] = params.subject
    if params.body is not None:
        if params.preserve_history:
            raw_draft = await g.get(
                f"/me/messages/{params.draft_id}",
                params={"$select": "id,isDraft,body"},
                headers={"Prefer": 'outlook.body-content-type="html"'},
            )
            current = GraphMessage.model_validate(raw_draft or {})
            if not current.id or not current.is_draft:
                return UpdateDraftResponse(
                    success=False,
                    action="update_draft",
                    draft_id=params.draft_id,
                    updated_fields=[],
                    error="Microsoft Graph did not return a valid draft to update.",
                )
            patch["body"] = {
                "contentType": "HTML",
                "content": _prepend_draft_body(
                    new_body=params.body,
                    new_body_type=params.body_type,
                    existing_body=current.body.content or "",
                    existing_body_type=current.body.content_type or "html",
                ),
            }
        else:
            patch["body"] = {
                "contentType": "HTML" if params.body_type.lower() == "html" else "Text",
                "content": params.body,
            }
    if params.to is not None:
        patch["toRecipients"] = parse_recipients(params.to)
    if params.cc is not None:
        patch["ccRecipients"] = parse_recipients(params.cc)

    if not patch:
        return UpdateDraftResponse(
            success=False,
            action="update_draft",
            draft_id=params.draft_id,
            updated_fields=[],
            error="No fields to update.",
        )

    result = await g.patch(f"/me/messages/{params.draft_id}", json=patch)
    updated = GraphMessage.model_validate(result or {"id": params.draft_id})
    updated_fields = ", ".join(patch.keys())
    return UpdateDraftResponse(
        success=True,
        action="update_draft",
        draft_id=updated.id or params.draft_id,
        updated_fields=list(patch.keys()),
        updated_fields_display=updated_fields,
    )


# ---------------------------------------------------------------------------
# send_draft
# ---------------------------------------------------------------------------


async def send_draft(params: SendDraftInput) -> SendDraftResponse:
    """
    Send an existing email draft from the Drafts folder.

    Args:
        draft_id: The Graph message ID of the draft to send.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured send confirmation.
    """
    g = get_graph(params.profile)
    await g.post(f"/me/messages/{params.draft_id}/send", json={})
    return SendDraftResponse(success=True, action="send_draft", draft_id=params.draft_id)


def register(server) -> None:
    """Register all draft tools with the given FastMCP server instance."""
    register_tool(server, create_draft, annotations=WRITE_TOOL)
    register_tool(server, create_reply_draft, annotations=WRITE_TOOL)
    register_tool(server, list_drafts, annotations=READ_ONLY_TOOL)
    register_tool(server, get_draft, annotations=READ_ONLY_TOOL)
    register_tool(server, update_draft, annotations=WRITE_TOOL)
    register_tool(server, send_draft, annotations=WRITE_TOOL)
