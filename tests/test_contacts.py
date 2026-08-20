"""
tests/test_contacts.py — Test coverage for the Contacts module.

Uses monkeypatch to mock the Graph API client, following the
pattern established in test_review_fixes.py.
"""

from __future__ import annotations

import base64
from pathlib import Path

import pytest

from mcp_microsoft.models import (
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
import mcp_microsoft.server as server
from mcp_microsoft.tools import contacts


# ---------------------------------------------------------------------------
# Tool Registration Test
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_contact_tools_are_registered() -> None:
    """Verify all 8 contact tools are registered with the MCP server."""
    tool_names = {tool.name for tool in await server.mcp.list_tools(run_middleware=False)}
    assert "list_contacts" in tool_names
    assert "get_contact" in tool_names
    assert "create_contact" in tool_names
    assert "update_contact" in tool_names
    assert "delete_contact" in tool_names
    assert "list_contact_folders" in tool_names
    assert "search_contacts" in tool_names
    assert "get_contact_photo" in tool_names


# ---------------------------------------------------------------------------
# list_contacts
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_list_contacts_returns_structured_response(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_contacts returns ListContactsResponse with normalized contacts."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            return {
                "value": [
                    {
                        "id": "contact-1",
                        "displayName": "Alice Smith",
                        "givenName": "Alice",
                        "surname": "Smith",
                        "emailAddresses": [{"address": "alice@example.com", "name": "Alice"}],
                        "mobilePhone": "555-1234",
                        "businessPhones": ["555-5678"],
                        "jobTitle": "Engineer",
                        "companyName": "Acme Corp",
                        "department": "Engineering",
                    }
                ]
            }

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.list_contacts(contacts.ListContactsInput(top=10))

    assert isinstance(result, ListContactsResponse)
    assert result.count == 1
    assert len(result.contacts) == 1
    assert result.contacts[0].display_name == "Alice Smith"


@pytest.mark.asyncio
async def test_list_contacts_with_folder_id(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_contacts respects folder_id parameter."""
    captured_path: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            captured_path["path"] = path
            return {"value": []}

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    await contacts.list_contacts(contacts.ListContactsInput(folder_id="folder-123"))

    assert captured_path["path"] == "/me/contactFolders/folder-123/contacts"


# ---------------------------------------------------------------------------
# get_contact
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_get_contact_returns_full_details(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify get_contact returns GetContactResponse including personal notes."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "contact-1",
                "displayName": "Alice Smith",
                "givenName": "Alice",
                "surname": "Smith",
                "emailAddresses": [{"address": "alice@example.com", "name": "Alice"}],
                "mobilePhone": "555-1234",
                "businessPhones": [],
                "jobTitle": "Engineer",
                "companyName": "Acme Corp",
                "department": "Engineering",
                "personalNotes": "Known since 2020",
            }

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.get_contact(contacts.GetContactInput(contact_id="contact-1"))

    assert isinstance(result, GetContactResponse)
    assert result.id == "contact-1"
    assert result.notes == "Known since 2020"


@pytest.mark.asyncio
async def test_get_contact_normalizes_nullable_graph_string_fields(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Verify nullable Graph contact fields are accepted and normalized."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "contact-null",
                "displayName": "Null Test",
                "givenName": None,
                "surname": None,
                "emailAddresses": None,
                "mobilePhone": None,
                "businessPhones": None,
                "jobTitle": None,
                "companyName": None,
                "department": None,
                "personalNotes": None,
            }

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.get_contact(contacts.GetContactInput(contact_id="contact-null"))

    assert result.id == "contact-null"
    assert result.given_name == ""
    assert result.surname == ""
    assert result.email_addresses == []
    assert result.mobile_phone == ""
    assert result.business_phones == []
    assert result.job_title == ""
    assert result.company_name == ""
    assert result.department == ""
    assert result.notes == ""


# ---------------------------------------------------------------------------
# create_contact
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_create_contact_posts_to_graph(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify create_contact sends correct POST body to Graph API."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {"id": "new-contact-id", "displayName": "Bob Johnson"}

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.create_contact(
        contacts.CreateContactInput(
            display_name="Bob Johnson",
            given_name="Bob",
            surname="Johnson",
            email_addresses=[{"address": "bob@example.com", "name": "Bob"}],
        )
    )

    assert isinstance(result, CreateContactResponse)
    assert result.success is True
    assert result.contact_id == "new-contact-id"
    assert captured["path"] == "/me/contacts"
    assert captured["json"]["displayName"] == "Bob Johnson"
    assert captured["json"]["givenName"] == "Bob"


# ---------------------------------------------------------------------------
# update_contact
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_update_contact_patches_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify update_contact sends PATCH with only changed fields."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def patch(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {"id": "contact-1"}

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.update_contact(
        contacts.UpdateContactInput(
            contact_id="contact-1",
            job_title="Senior Engineer",
            department="R&D",
        )
    )

    assert isinstance(result, UpdateContactResponse)
    assert result.success is True
    assert set(result.updated_fields) == {"jobTitle", "department"}
    assert captured["json"] == {"jobTitle": "Senior Engineer", "department": "R&D"}


@pytest.mark.asyncio
async def test_update_contact_with_no_fields_returns_error(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify update_contact returns error when no fields provided."""
    monkeypatch.setattr(contacts, "get_graph", lambda _profile: None)
    result = await contacts.update_contact(
        contacts.UpdateContactInput(contact_id="contact-1")
    )

    assert result.success is False
    assert result.error == "No fields to update."


# ---------------------------------------------------------------------------
# delete_contact
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_delete_contact_calls_delete(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify delete_contact sends DELETE request to correct path."""
    captured_path: dict[str, str] = {}

    class DummyGraph:
        async def delete(self, path: str):
            captured_path["path"] = path

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.delete_contact(
        contacts.DeleteContactInput(contact_id="contact-1")
    )

    assert isinstance(result, DeleteContactResponse)
    assert result.success is True
    assert captured_path["path"] == "/me/contacts/contact-1"


# ---------------------------------------------------------------------------
# list_contact_folders
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_list_contact_folders_returns_structured_response(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_contact_folders returns folder list with item counts."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "folder-1",
                        "displayName": "Contacts",
                        "parentFolderId": "parent-1",
                    }
                ]
            }

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.list_contact_folders(contacts.ListContactFoldersInput())

    assert isinstance(result, ListContactFoldersResponse)
    assert result.count == 1
    assert result.folders[0].display_name == "Contacts"
    assert result.folders[0].total_item_count == 0


# ---------------------------------------------------------------------------
# search_contacts
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_search_contacts_matches_names_and_emails_locally(monkeypatch: pytest.MonkeyPatch) -> None:
    """The contacts endpoint does not support search; matching is done locally."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            captured["params"] = params
            return {
                "value": [
                    {"id": "1", "displayName": "Alice Jones", "emailAddresses": []},
                    {"id": "2", "displayName": "Bob Jones", "emailAddresses": [{"address": "alice@example.com"}]},
                    {"id": "3", "displayName": "Carol Smith", "emailAddresses": []},
                ]
            }

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.search_contacts(contacts.SearchContactsInput(query="Alice"))

    assert "$filter" not in captured["params"]
    assert [contact.id for contact in result.contacts] == ["1", "2"]


@pytest.mark.asyncio
async def test_search_contacts_treats_quotes_as_plain_text(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            captured["params"] = params
            return {"value": [{"id": "1", "displayName": "O'Reilly", "emailAddresses": []}]}

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.search_contacts(contacts.SearchContactsInput(query="O'Reilly"))

    assert "$filter" not in captured["params"]
    assert [contact.display_name for contact in result.contacts] == ["O'Reilly"]


@pytest.mark.asyncio
async def test_search_contacts_does_not_send_query_as_odata(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            captured["params"] = params
            return {"value": []}

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    # Classic OData break-out attempt: close the string and append a tautology.
    await contacts.search_contacts(contacts.SearchContactsInput(query="x') or (true"))

    assert "$filter" not in captured["params"]


# ---------------------------------------------------------------------------
# get_contact_photo
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_get_contact_photo_returns_base64(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify get_contact_photo returns base64-encoded photo bytes."""
    photo_bytes = b"fake-image-data"

    class DummyGraph:
        async def get_raw(self, path: str):
            return photo_bytes

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.get_contact_photo(
        contacts.GetContactPhotoInput(contact_id="contact-1")
    )

    assert isinstance(result, GetContactPhotoResponse)
    assert result.success is True
    assert result.contact_id == "contact-1"
    assert result.photo_base64 == base64.b64encode(photo_bytes).decode("ascii")
    assert result.size_bytes == len(photo_bytes)


@pytest.mark.asyncio
async def test_get_contact_photo_saves_to_file(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Verify get_contact_photo saves photo bytes to disk when save_path is provided."""
    photo_bytes = b"fake-image-data"
    save_path = str(tmp_path / "photo.jpg")

    class DummyGraph:
        async def get_raw(self, path: str):
            return photo_bytes

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.get_contact_photo(
        contacts.GetContactPhotoInput(contact_id="contact-1", save_path=save_path)
    )

    assert result.success is True
    assert result.saved_path == save_path
    assert Path(save_path).read_bytes() == photo_bytes
