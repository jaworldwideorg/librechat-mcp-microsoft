"""
Contacts tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.
Endpoints live under /me/contacts and /me/contactFolders in Graph v1.0.

Implemented:
  - list_contacts
  - get_contact
  - create_contact
  - update_contact
  - delete_contact
  - list_contact_folders
  - search_contacts
  - get_contact_photo
"""

from __future__ import annotations

import base64
import binascii
import hashlib
import hmac
from pathlib import Path
from typing import Any, Literal, Optional
from urllib.parse import unquote, urlsplit

from pydantic import BaseModel, ConfigDict, Field, ValidationError

from mcp_microsoft.common.tooling import DESTRUCTIVE_TOOL, READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.config import get_app_config
from mcp_microsoft.feature_flags import is_deletion_disabled
from mcp_microsoft.graph_types import (
    GraphContact,
    GraphContactFolder,
    parse_graph_collection,
)
from mcp_microsoft.models import (
    ContactEmailAddress,
    ContactFolderInfo,
    ContactInfo,
    CreateContactResponse,
    DeleteContactResponse,
    GetContactPhotoResponse,
    GetContactResponse,
    ListContactFoldersResponse,
    ListContactsResponse,
    SearchContactsResponse,
    UpdateContactResponse,
)
from mcp_microsoft.graph import get_graph
from mcp_microsoft.common.request_model import ToolRequestModel

# Fields to $select for contact list/search operations (lighter payload)
_CONTACT_SELECT = (
    "id,displayName,givenName,surname,emailAddresses,"
    "mobilePhone,businessPhones,jobTitle,companyName,department"
)

# Fields to $select for full contact detail
_CONTACT_DETAIL_SELECT = (
    "id,displayName,givenName,surname,emailAddresses,"
    "mobilePhone,businessPhones,jobTitle,companyName,department,personalNotes"
)

_CONTACT_SEARCH_MAX_PAGES = 5
_CONTACT_SEARCH_SCAN_PAGE_SIZE = 100
_CONTACT_SEARCH_CURSOR_MAX_LENGTH = 16_384
_GRAPH_CONTINUATION_MAX_LENGTH = 12_000
_CONTACT_SEARCH_CURSOR_PREFIX = "c1."


class _ContactCursorState(BaseModel):
    """Versioned state for resuming inside a locally filtered Graph page."""

    model_config = ConfigDict(extra="forbid")

    version: Literal[1] = 1
    fingerprint: str = Field(pattern=r"^[0-9a-f]{64}$")
    page_link: str | None = Field(default=None, max_length=_GRAPH_CONTINUATION_MAX_LENGTH)
    offset: int = Field(ge=0, le=_CONTACT_SEARCH_SCAN_PAGE_SIZE)


class ListContactsInput(ToolRequestModel):
    folder_id: str | None = None
    search: str | None = Field(default=None, min_length=1, max_length=256)
    top: int = 50
    skip_token: str | None = Field(default=None, max_length=_CONTACT_SEARCH_CURSOR_MAX_LENGTH)
    profile: str | None = None


class GetContactInput(ToolRequestModel):
    contact_id: str
    profile: str | None = None


class CreateContactInput(ToolRequestModel):
    display_name: str
    given_name: Optional[str] = None
    surname: Optional[str] = None
    email_addresses: Optional[list[dict]] = None
    mobile_phone: Optional[str] = None
    business_phones: Optional[list[str]] = None
    job_title: Optional[str] = None
    company_name: Optional[str] = None
    department: Optional[str] = None
    notes: Optional[str] = None
    folder_id: Optional[str] = None
    profile: str | None = None


class UpdateContactInput(ToolRequestModel):
    contact_id: str
    display_name: Optional[str] = None
    given_name: Optional[str] = None
    surname: Optional[str] = None
    email_addresses: Optional[list[dict]] = None
    mobile_phone: Optional[str] = None
    business_phones: Optional[list[str]] = None
    job_title: Optional[str] = None
    company_name: Optional[str] = None
    department: Optional[str] = None
    notes: Optional[str] = None
    profile: str | None = None


class DeleteContactInput(ToolRequestModel):
    contact_id: str
    profile: str | None = None


class ListContactFoldersInput(ToolRequestModel):
    profile: str | None = None


class SearchContactsInput(ToolRequestModel):
    query: str = Field(min_length=1, max_length=256)
    top: int = Field(default=25, ge=1, le=100)
    skip_token: str | None = Field(default=None, max_length=_CONTACT_SEARCH_CURSOR_MAX_LENGTH)
    profile: str | None = None


class GetContactPhotoInput(ToolRequestModel):
    contact_id: str
    save_path: Optional[str] = None
    profile: str | None = None


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


def _normalize_contact(c: GraphContact | dict[str, Any]) -> ContactInfo:
    """Normalize a typed Graph contact object into a ContactInfo summary."""
    if not isinstance(c, GraphContact):
        c = GraphContact.model_validate(c)
    return ContactInfo(
        id=c.id,
        display_name=c.display_name or "",
        given_name=c.given_name or "",
        surname=c.surname or "",
        email_addresses=[
            ContactEmailAddress(
                address=ea.address,
                name=ea.name,
            )
            for ea in (c.email_addresses or [])
        ],
        mobile_phone=c.mobile_phone or "",
        business_phones=c.business_phones or [],
        job_title=c.job_title or "",
        company_name=c.company_name or "",
        department=c.department or "",
    )


def _normalize_search_query(query: str) -> str:
    """Validate a local contact-search term before any Graph requests."""
    normalized = query.strip()
    if not normalized:
        raise ValueError("search query must not be empty")
    if len(normalized) > 256 or any(ord(char) < 0x20 for char in normalized):
        raise ValueError("search query must be ≤256 printable characters")
    return normalized


def _contact_cursor_fingerprint(path: str, query: str | None) -> str:
    """Bind a cursor to one contact collection and normalized search term."""
    material = f"{path}\0{(query or '').casefold()}".encode()
    return hashlib.sha256(material).hexdigest()


def _graph_continuation_request(next_link: str, expected_path: str) -> str:
    """Validate a Graph nextLink and preserve its complete path and query."""
    if not next_link or len(next_link) > _GRAPH_CONTINUATION_MAX_LENGTH:
        raise ValueError("Microsoft Graph returned an invalid contact continuation link")

    parsed = urlsplit(next_link)
    expected_graph_path = f"/v1.0{expected_path}"
    if (
        parsed.scheme.casefold() != "https"
        or parsed.netloc.casefold() != "graph.microsoft.com"
        or unquote(parsed.path) != expected_graph_path
        or not parsed.query
        or parsed.fragment
    ):
        raise ValueError("Microsoft Graph returned an invalid contact continuation link")

    relative_path = parsed.path.removeprefix("/v1.0")
    return f"{relative_path}?{parsed.query}"


def _encode_contact_cursor(
    *,
    fingerprint: str,
    page_link: str | None,
    offset: int,
) -> str:
    state = _ContactCursorState(
        fingerprint=fingerprint,
        page_link=page_link,
        offset=offset,
    )
    payload = base64.urlsafe_b64encode(state.model_dump_json().encode()).rstrip(b"=")
    cursor = f"{_CONTACT_SEARCH_CURSOR_PREFIX}{payload.decode('ascii')}"
    if len(cursor) > _CONTACT_SEARCH_CURSOR_MAX_LENGTH:
        raise ValueError("Microsoft Graph returned an oversized contact continuation link")
    return cursor


def _decode_contact_cursor(
    cursor: str,
    *,
    fingerprint: str,
    expected_path: str,
) -> tuple[str | None, int]:
    try:
        if not cursor.startswith(_CONTACT_SEARCH_CURSOR_PREFIX):
            raise ValueError
        encoded = cursor.removeprefix(_CONTACT_SEARCH_CURSOR_PREFIX)
        padding = "=" * (-len(encoded) % 4)
        payload = base64.b64decode(
            f"{encoded}{padding}",
            altchars=b"-_",
            validate=True,
        )
        state = _ContactCursorState.model_validate_json(payload)
    except (binascii.Error, UnicodeError, ValidationError, ValueError) as exc:
        raise ValueError("contact search cursor is invalid or expired") from None

    if not hmac.compare_digest(state.fingerprint, fingerprint):
        raise ValueError("contact search cursor does not match this query or folder")
    if state.page_link is not None:
        _graph_continuation_request(state.page_link, expected_path)
    return state.page_link, state.offset


def _contact_matches(contact: GraphContact, query: str) -> bool:
    needle = query.casefold()
    values = [contact.display_name, contact.given_name, contact.surname]
    values.extend(address.address for address in (contact.email_addresses or []))
    values.extend(address.name for address in (contact.email_addresses or []))
    return any(needle in (value or "").casefold() for value in values)


async def _scan_contact_pages(
    g: Any,
    *,
    path: str,
    query: str | None,
    top: int,
    cursor: str | None,
) -> tuple[list[GraphContact], str | None, int]:
    """Page contacts with a scan size independent from the returned result size."""
    normalized_query = _normalize_search_query(query) if query is not None else None
    fingerprint = _contact_cursor_fingerprint(path, normalized_query)
    page_link, page_offset = (
        _decode_contact_cursor(
            cursor,
            fingerprint=fingerprint,
            expected_path=path,
        )
        if cursor is not None
        else (None, 0)
    )
    matches: list[GraphContact] = []
    pages_scanned = 0
    max_pages = _CONTACT_SEARCH_MAX_PAGES if normalized_query is not None else 1

    while pages_scanned < max_pages:
        current_page_link = page_link
        if current_page_link is None:
            result = await g.get(
                path,
                params={
                    "$select": _CONTACT_SELECT,
                    "$top": _CONTACT_SEARCH_SCAN_PAGE_SIZE,
                    "$orderby": "displayName",
                },
            )
        else:
            request_path = _graph_continuation_request(current_page_link, path)
            result = await g.get(request_path)

        page = parse_graph_collection(result or {}, GraphContact)
        pages_scanned += 1
        if len(page) > _CONTACT_SEARCH_SCAN_PAGE_SIZE:
            raise ValueError("Microsoft Graph returned an oversized contact page")
        if page_offset > len(page):
            raise ValueError("contact search cursor is invalid or expired")

        next_link = (result or {}).get("@odata.nextLink") or None
        if next_link is not None:
            _graph_continuation_request(next_link, path)

        index = page_offset
        while index < len(page):
            contact = page[index]
            index += 1
            if normalized_query is None or _contact_matches(contact, normalized_query):
                matches.append(contact)
                if len(matches) == top:
                    if index < len(page):
                        next_cursor = _encode_contact_cursor(
                            fingerprint=fingerprint,
                            page_link=current_page_link,
                            offset=index,
                        )
                    elif next_link is not None:
                        next_cursor = _encode_contact_cursor(
                            fingerprint=fingerprint,
                            page_link=next_link,
                            offset=0,
                        )
                    else:
                        next_cursor = None
                    return matches, next_cursor, pages_scanned

        if next_link is None:
            return matches, None, pages_scanned
        page_link = next_link
        page_offset = 0

    return (
        matches,
        _encode_contact_cursor(
            fingerprint=fingerprint,
            page_link=page_link,
            offset=page_offset,
        ),
        pages_scanned,
    )


# ---------------------------------------------------------------------------
# list_contacts
# ---------------------------------------------------------------------------


async def list_contacts(params: ListContactsInput) -> ListContactsResponse:
    """
    List contacts from the user's default contacts folder or a specific folder.

    Args:
        top: Maximum number of contacts to return (1-100). Defaults to 25.
        folder_id: Optional contact folder ID. Omit to use the default contacts folder.
        search: Optional case-insensitive text matched against contact names and
            email addresses.
        skip_token: Opaque pagination cursor returned by a previous call.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured list of contacts with id, name, email, phone, and job info.
        When has_more is True, pass next_page_token as skip_token to continue
        scanning without repeating contacts.
    """
    g = get_graph(params.profile)
    top = max(1, min(params.top, 100))

    if params.folder_id:
        path = f"/me/contactFolders/{params.folder_id}/contacts"
    else:
        path = "/me/contacts"

    contacts, next_page_token, pages_scanned = await _scan_contact_pages(
        g,
        path=path,
        query=params.search,
        top=top,
        cursor=params.skip_token,
    )

    return ListContactsResponse(
        count=len(contacts),
        folder_id=params.folder_id,
        contacts=[_normalize_contact(contact) for contact in contacts],
        next_page_token=next_page_token,
        has_more=(next_page_token is not None),
        pages_scanned=pages_scanned,
    )


# ---------------------------------------------------------------------------
# get_contact
# ---------------------------------------------------------------------------


async def get_contact(params: GetContactInput) -> GetContactResponse:
    """
    Fetch a single contact by ID with full details including notes.

    Args:
        contact_id: The Graph contact ID to fetch.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Full contact details including personal notes.
    """
    g = get_graph(params.profile)
    query = {"$select": _CONTACT_DETAIL_SELECT}

    c = GraphContact.model_validate(await g.get(f"/me/contacts/{params.contact_id}", params=query))

    return GetContactResponse(
        id=c.id or params.contact_id,
        display_name=c.display_name or "",
        given_name=c.given_name or "",
        surname=c.surname or "",
        email_addresses=[
            ContactEmailAddress(
                address=ea.address,
                name=ea.name,
            )
            for ea in (c.email_addresses or [])
        ],
        mobile_phone=c.mobile_phone or "",
        business_phones=c.business_phones or [],
        job_title=c.job_title or "",
        company_name=c.company_name or "",
        department=c.department or "",
        notes=c.personal_notes or "",
    )


# ---------------------------------------------------------------------------
# create_contact
# ---------------------------------------------------------------------------


async def create_contact(params: CreateContactInput) -> CreateContactResponse:
    """
    Create a new contact in the user's contacts folder.

    Args:
        display_name: Contact display name shown in address books.
        given_name: Optional first name.
        surname: Optional last name.
        email_addresses: Optional list of `{address, name}` email objects.
        mobile_phone: Optional mobile phone number.
        business_phones: Optional list of business phone numbers.
        job_title: Optional job title.
        company_name: Optional company name.
        department: Optional department name.
        notes: Optional personal notes.
        folder_id: Optional contact folder ID. Omit to create in the default folder.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Created contact ID and display name.
    """
    g = get_graph(params.profile)

    body: dict[str, Any] = {"displayName": params.display_name}

    if params.given_name is not None:
        body["givenName"] = params.given_name
    if params.surname is not None:
        body["surname"] = params.surname
    if params.email_addresses is not None:
        body["emailAddresses"] = [
            {"address": ea.get("address", ""), "name": ea.get("name", "")}
            for ea in params.email_addresses
        ]
    if params.mobile_phone is not None:
        body["mobilePhone"] = params.mobile_phone
    if params.business_phones is not None:
        body["businessPhones"] = params.business_phones
    if params.job_title is not None:
        body["jobTitle"] = params.job_title
    if params.company_name is not None:
        body["companyName"] = params.company_name
    if params.department is not None:
        body["department"] = params.department
    if params.notes is not None:
        body["personalNotes"] = params.notes

    if params.folder_id:
        path = f"/me/contactFolders/{params.folder_id}/contacts"
    else:
        path = "/me/contacts"

    created = GraphContact.model_validate(await g.post(path, json=body) or {})

    return CreateContactResponse(
        success=True,
        action="create_contact",
        contact_id=created.id or "unknown",
        display_name=created.display_name or params.display_name,
        folder_id=params.folder_id,
    )


# ---------------------------------------------------------------------------
# update_contact
# ---------------------------------------------------------------------------


async def update_contact(params: UpdateContactInput) -> UpdateContactResponse:
    """
    Update an existing contact. Only provided fields are changed (PATCH semantics).

    Args:
        contact_id: The Graph contact ID to update.
        display_name: Replace the display name.
        given_name: Replace the first name.
        surname: Replace the last name.
        email_addresses: Replace the full email-address list.
        mobile_phone: Replace the mobile phone number.
        business_phones: Replace the business phone list.
        job_title: Replace the job title.
        company_name: Replace the company name.
        department: Replace the department.
        notes: Replace the personal notes.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Updated contact ID and list of changed fields.
    """
    g = get_graph(params.profile)
    patch: dict[str, Any] = {}

    if params.display_name is not None:
        patch["displayName"] = params.display_name
    if params.given_name is not None:
        patch["givenName"] = params.given_name
    if params.surname is not None:
        patch["surname"] = params.surname
    if params.email_addresses is not None:
        patch["emailAddresses"] = [
            {"address": ea.get("address", ""), "name": ea.get("name", "")}
            for ea in params.email_addresses
        ]
    if params.mobile_phone is not None:
        patch["mobilePhone"] = params.mobile_phone
    if params.business_phones is not None:
        patch["businessPhones"] = params.business_phones
    if params.job_title is not None:
        patch["jobTitle"] = params.job_title
    if params.company_name is not None:
        patch["companyName"] = params.company_name
    if params.department is not None:
        patch["department"] = params.department
    if params.notes is not None:
        patch["personalNotes"] = params.notes

    if not patch:
        return UpdateContactResponse(
            success=False,
            action="update_contact",
            contact_id=params.contact_id,
            updated_fields=[],
            error="No fields to update.",
        )

    result = await g.patch(f"/me/contacts/{params.contact_id}", json=patch)
    updated = GraphContact.model_validate(result or {"id": params.contact_id})
    updated_fields = list(patch.keys())

    return UpdateContactResponse(
        success=True,
        action="update_contact",
        contact_id=updated.id or params.contact_id,
        updated_fields=updated_fields,
        updated_fields_display=", ".join(updated_fields),
    )


# ---------------------------------------------------------------------------
# delete_contact
# ---------------------------------------------------------------------------


async def delete_contact(params: DeleteContactInput) -> DeleteContactResponse:
    """
    Permanently delete a contact by ID.

    Args:
        contact_id: The Graph contact ID to delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Deletion confirmation.
    """
    g = get_graph(params.profile)
    await g.delete(f"/me/contacts/{params.contact_id}")
    return DeleteContactResponse(
        success=True,
        action="delete_contact",
        contact_id=params.contact_id,
    )


# ---------------------------------------------------------------------------
# list_contact_folders
# ---------------------------------------------------------------------------


async def list_contact_folders(
    params: ListContactFoldersInput,
) -> ListContactFoldersResponse:
    """
    List all contact folders in the user's mailbox.

    Returns folder IDs that can be passed to list_contacts or create_contact
    to scope operations to a specific folder.

    Args:
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured list of contact folders with id and displayName.
        Microsoft Graph contactFolder does not expose an item count field, so
        total_item_count is returned as 0 for compatibility.
    """
    g = get_graph(params.profile)
    query = {
        "$select": "id,displayName,parentFolderId",
        "$top": 100,
    }

    result = await g.get("/me/contactFolders", params=query)
    folders = parse_graph_collection(result or {}, GraphContactFolder)

    return ListContactFoldersResponse(
        count=len(folders),
        folders=[
            ContactFolderInfo(
                id=f.id,
                display_name=f.display_name,
                parent_folder_id=f.parent_folder_id or "",
                total_item_count=0,
            )
            for f in folders
        ],
    )


# ---------------------------------------------------------------------------
# search_contacts
# ---------------------------------------------------------------------------


async def search_contacts(params: SearchContactsInput) -> SearchContactsResponse:
    """
    Search contacts locally across names and email addresses.

    Microsoft Graph does not support general contact search. This tool scans at
    most five ordered Graph pages per call and returns an opaque cursor when
    more contacts remain.

    Args:
        query: Case-insensitive text to match against names and email addresses.
        top: Maximum number of contacts to return (1-100). Defaults to 50.
        skip_token: Opaque pagination cursor returned by a previous call.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Matching contacts and a continuation cursor when more contacts remain
        to scan. A continuation can yield no matches before the scan completes.
    """
    g = get_graph(params.profile)
    contacts, next_page_token, pages_scanned = await _scan_contact_pages(
        g,
        path="/me/contacts",
        query=params.query,
        top=params.top,
        cursor=params.skip_token,
    )

    return SearchContactsResponse(
        query=params.query,
        count=len(contacts),
        contacts=[_normalize_contact(c) for c in contacts],
        next_page_token=next_page_token,
        has_more=(next_page_token is not None),
        pages_scanned=pages_scanned,
    )


# ---------------------------------------------------------------------------
# get_contact_photo
# ---------------------------------------------------------------------------


async def get_contact_photo(params: GetContactPhotoInput) -> GetContactPhotoResponse:
    """
    Retrieve the profile photo for a contact.

    Args:
        contact_id: The Graph contact ID.
        save_path: Optional filesystem path to write the photo to. Omit to return base64 only.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Base64-encoded photo bytes and/or saved file path confirmation.
    """
    if params.save_path and get_app_config().transport == "http":
        raise ValueError(
            "save_path is not available in multi-user http mode (the server's "
            "disk is not the caller's disk); omit it to receive base64 only."
        )

    g = get_graph(params.profile)
    raw: bytes = await g.get_raw(f"/me/contacts/{params.contact_id}/photo/$value")

    photo_b64 = base64.b64encode(raw).decode("ascii")
    saved_path = ""

    if params.save_path:
        dest = Path(params.save_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(raw)
        saved_path = str(dest)

    return GetContactPhotoResponse(
        success=True,
        action="get_contact_photo",
        contact_id=params.contact_id,
        photo_base64=photo_b64,
        saved_path=saved_path,
        size_bytes=len(raw),
    )


def register(server) -> None:
    """Register all contact tools with the given FastMCP server instance."""
    register_tool(server, list_contacts, annotations=READ_ONLY_TOOL)
    register_tool(server, get_contact, annotations=READ_ONLY_TOOL)
    register_tool(server, create_contact, annotations=WRITE_TOOL)
    register_tool(server, update_contact, annotations=WRITE_TOOL)
    if not is_deletion_disabled():
        register_tool(server, delete_contact, annotations=DESTRUCTIVE_TOOL)
    register_tool(server, list_contact_folders, annotations=READ_ONLY_TOOL)
    register_tool(server, search_contacts, annotations=READ_ONLY_TOOL)
    register_tool(server, get_contact_photo, annotations=READ_ONLY_TOOL)
