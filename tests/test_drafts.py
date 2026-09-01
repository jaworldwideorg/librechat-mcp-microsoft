"""
tests/test_drafts.py — Test coverage for the Drafts module.

Uses monkeypatch to mock the Graph API client, following the
pattern established in test_calendar.py.
"""

from __future__ import annotations

import json

import httpx
import pytest

from mcp_microsoft.graph_types import GraphEmailAddress, GraphRecipient
from mcp_microsoft.models import (
    CreateDraftResponse,
    CreateReplyDraftResponse,
    DraftDetailResponse,
    ListDraftsResponse,
    SendDraftResponse,
    UpdateDraftResponse,
)
import mcp_microsoft.server as server
from mcp_microsoft.tools import drafts


def _make_recipient(address: str, name: str = "") -> GraphRecipient:
    """Build a GraphRecipient as the real Graph client would produce."""
    return GraphRecipient(emailAddress=GraphEmailAddress(address=address, name=name))


# ---------------------------------------------------------------------------
# Tool Registration Test
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_draft_tools_are_registered() -> None:
    """Verify all draft tools are registered with the MCP server."""
    tool_names = {tool.name for tool in await server.mcp.list_tools(run_middleware=False)}
    assert "create_draft" in tool_names
    assert "create_reply_draft" in tool_names
    assert "list_drafts" in tool_names
    assert "get_draft" in tool_names
    assert "update_draft" in tool_names
    assert "send_draft" in tool_names


# ---------------------------------------------------------------------------
# create_draft
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_create_draft_single_recipient(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify create_draft posts a message and returns CreateDraftResponse."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {"id": "draft-abc"}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.create_draft(
        drafts.CreateDraftInput(
            to="alice@example.com",
            subject="Hello Alice",
            body="Just checking in.",
        )
    )

    assert isinstance(result, CreateDraftResponse)
    assert result.success is True
    assert result.draft_id == "draft-abc"
    assert result.subject == "Hello Alice"
    assert "alice@example.com" in result.to
    assert result.cc == []
    assert captured["path"] == "/me/messages"
    body = captured["json"]
    assert body["subject"] == "Hello Alice"
    assert body["body"]["contentType"] == "Text"
    assert body["body"]["content"] == "Just checking in."


@pytest.mark.asyncio
async def test_create_draft_multiple_recipients_and_cc(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify create_draft includes CC when provided and handles multiple To addresses."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {"id": "draft-multi"}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.create_draft(
        drafts.CreateDraftInput(
            to=["alice@example.com", "bob@example.com"],
            subject="Team Update",
            body="<p>Updates</p>",
            cc="manager@example.com",
            body_type="html",
        )
    )

    assert result.success is True
    assert len(result.to) == 2
    assert result.body_type == "html"
    body = captured["json"]
    assert body["body"]["contentType"] == "HTML"
    assert "ccRecipients" in body
    assert len(body["ccRecipients"]) == 1
    assert body["ccRecipients"][0]["emailAddress"]["address"] == "manager@example.com"


@pytest.mark.asyncio
async def test_create_draft_comma_separated_to(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify create_draft splits a comma-separated to string into multiple recipients."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {"id": "draft-csv"}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.create_draft(
        drafts.CreateDraftInput(
            to="alice@example.com, bob@example.com",
            subject="CSV test",
            body="body",
        )
    )

    assert result.success is True
    body = captured["json"]
    assert len(body["toRecipients"]) == 2


# ---------------------------------------------------------------------------
# create_reply_draft
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_create_empty_reply_draft(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {"id": "reply-draft", "isDraft": True}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.create_reply_draft(
        drafts.CreateReplyDraftInput(message_id="original-message")
    )

    assert isinstance(result, CreateReplyDraftResponse)
    assert result.success is True
    assert result.action == "create_reply_draft"
    assert result.draft_id == "reply-draft"
    assert result.original_message_id == "original-message"
    assert result.reply_all is False
    assert result.body_type == "text"
    assert captured == {
        "path": "/me/messages/original-message/createReply",
        "json": {},
    }


@pytest.mark.asyncio
async def test_create_text_reply_draft_uses_comment(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {"id": "reply-draft", "isDraft": True}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    await drafts.create_reply_draft(
        drafts.CreateReplyDraftInput(
            message_id="original-message",
            body="Thanks for the update.",
        )
    )

    assert captured["json"] == {"comment": "Thanks for the update."}


@pytest.mark.asyncio
async def test_create_html_reply_all_draft_uses_comment_to_preserve_history(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {"id": "reply-all-draft", "isDraft": True}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.create_reply_draft(
        drafts.CreateReplyDraftInput(
            message_id="original-message",
            body="<p>Thanks everyone.</p>",
            body_type="html",
            reply_all=True,
        )
    )

    assert result.reply_all is True
    assert result.body_type == "html"
    assert captured["path"] == "/me/messages/original-message/createReplyAll"
    assert captured["json"] == {"comment": "<p>Thanks everyone.</p>"}
    assert "message" not in captured["json"]


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "graph_response",
    [
        {"isDraft": True},
        {"id": "not-a-draft", "isDraft": False},
    ],
)
async def test_create_reply_draft_rejects_invalid_graph_response(
    monkeypatch: pytest.MonkeyPatch,
    graph_response: dict,
) -> None:
    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            return graph_response

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.create_reply_draft(
        drafts.CreateReplyDraftInput(message_id="original-message")
    )

    assert result.success is False
    assert result.action == "create_reply_draft"
    assert result.error == "Microsoft Graph did not return a valid reply draft."
    assert result.original_message_id == "original-message"


# ---------------------------------------------------------------------------
# list_drafts
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_list_drafts_returns_structured_response(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_drafts returns ListDraftsResponse with DraftSummary items."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "draft-1",
                        "subject": "Draft One",
                        "toRecipients": [_make_recipient("alice@example.com", "Alice")],
                        "lastModifiedDateTime": "2026-04-01T10:00:00Z",
                        "bodyPreview": "Hello there",
                    },
                    {
                        "id": "draft-2",
                        "subject": None,
                        "toRecipients": [],
                        "lastModifiedDateTime": "2026-04-02T09:00:00Z",
                        "bodyPreview": "",
                    },
                ]
            }

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.list_drafts(drafts.ListDraftsInput(max_results=10))

    assert isinstance(result, ListDraftsResponse)
    assert result.count == 2
    assert result.drafts[0].id == "draft-1"
    assert result.drafts[0].subject == "Draft One"
    assert result.drafts[0].preview == "Hello there"
    assert result.drafts[1].subject == "(no subject)"


@pytest.mark.asyncio
async def test_list_drafts_empty(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_drafts handles an empty drafts folder gracefully."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {"value": []}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.list_drafts(drafts.ListDraftsInput())

    assert isinstance(result, ListDraftsResponse)
    assert result.count == 0
    assert result.drafts == []


@pytest.mark.asyncio
async def test_list_drafts_uses_correct_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_drafts requests the correct Graph endpoint."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    await drafts.list_drafts(drafts.ListDraftsInput())

    assert captured["path"] == "/me/mailFolders/drafts/messages"


# ---------------------------------------------------------------------------
# get_draft
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_get_draft_returns_full_details(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify get_draft returns DraftDetailResponse with all major fields."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "draft-xyz",
                "subject": "My Draft",
                "toRecipients": [_make_recipient("bob@example.com", "Bob")],
                "ccRecipients": [],
                "bccRecipients": [],
                "lastModifiedDateTime": "2026-04-01T12:00:00Z",
                "body": {"contentType": "text", "content": "Draft body text."},
                "bodyPreview": "Draft body text.",
                "isDraft": True,
            }

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.get_draft(drafts.GetDraftInput(draft_id="draft-xyz"))

    assert isinstance(result, DraftDetailResponse)
    assert result.success is True
    assert result.action == "get_draft"
    assert result.id == "draft-xyz"
    assert result.subject == "My Draft"
    assert result.body == "Draft body text."
    assert result.body_content_type == "text"
    assert result.is_draft is True
    assert len(result.to) == 1
    assert result.to[0].address == "bob@example.com"
    assert result.cc == []


@pytest.mark.asyncio
async def test_get_draft_strips_html_body(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify get_draft strips HTML tags when content type is html."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "draft-html",
                "subject": "HTML Draft",
                "toRecipients": [],
                "ccRecipients": [],
                "bccRecipients": [],
                "lastModifiedDateTime": None,
                "body": {"contentType": "html", "content": "<p>Hello <b>world</b></p>"},
                "bodyPreview": "Hello world",
                "isDraft": True,
            }

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.get_draft(drafts.GetDraftInput(draft_id="draft-html"))

    assert result.body_content_type == "html"
    assert "<p>" not in result.body
    assert "<b>" not in result.body
    assert "Hello" in result.body
    assert "world" in result.body


@pytest.mark.asyncio
async def test_get_draft_uses_correct_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify get_draft fetches /me/messages/<draft_id>."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {
                "id": "draft-path",
                "subject": "Path test",
                "toRecipients": [],
                "ccRecipients": [],
                "bccRecipients": [],
                "lastModifiedDateTime": None,
                "body": {"contentType": "text", "content": ""},
                "isDraft": True,
            }

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    await drafts.get_draft(drafts.GetDraftInput(draft_id="draft-path"))

    assert captured["path"] == "/me/messages/draft-path"


@pytest.mark.asyncio
async def test_get_draft_returns_structured_not_found(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    request = httpx.Request("GET", "https://graph.microsoft.com/v1.0/me/messages/missing")
    response = httpx.Response(404, request=request)

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            raise httpx.HTTPStatusError(
                "404 Not Found",
                request=request,
                response=response,
            )

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.get_draft(drafts.GetDraftInput(draft_id="missing"))

    assert isinstance(result, DraftDetailResponse)
    assert result.success is False
    assert result.action == "get_draft"
    assert result.error == (
        "Draft not found. It may have been deleted or sent, or the draft ID "
        "may no longer be valid."
    )
    assert result.id == ""
    assert result.is_draft is False


@pytest.mark.asyncio
async def test_get_draft_does_not_mask_other_graph_errors(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    request = httpx.Request("GET", "https://graph.microsoft.com/v1.0/me/messages/forbidden")
    response = httpx.Response(403, request=request)
    error = httpx.HTTPStatusError(
        "403 Forbidden",
        request=request,
        response=response,
    )

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            raise error

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())

    with pytest.raises(httpx.HTTPStatusError) as exc_info:
        await drafts.get_draft(drafts.GetDraftInput(draft_id="forbidden"))

    assert exc_info.value is error


# ---------------------------------------------------------------------------
# update_draft
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_update_draft_patches_subject_and_body(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify update_draft sends PATCH with only the provided fields."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def patch(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {"id": "draft-upd"}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.update_draft(
        drafts.UpdateDraftInput(
            draft_id="draft-upd",
            subject="New Subject",
            body="Updated content",
        )
    )

    assert isinstance(result, UpdateDraftResponse)
    assert result.success is True
    assert result.draft_id == "draft-upd"
    assert "subject" in result.updated_fields
    assert "body" in result.updated_fields
    assert "toRecipients" not in result.updated_fields
    assert captured["json"]["subject"] == "New Subject"
    assert captured["json"]["body"]["content"] == "Updated content"
    assert captured["path"] == "/me/messages/draft-upd"


@pytest.mark.asyncio
async def test_update_draft_can_prepend_body_and_preserve_history(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """History-preserving updates fetch and retain the complete current body."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            captured["get"] = (path, params, headers)
            return {
                "id": "reply-draft",
                "isDraft": True,
                "body": {
                    "contentType": "html",
                    "content": "<html><body><div id='quoted'>Original history</div></body></html>",
                },
            }

        async def patch(self, path: str, json: dict = None):
            captured["patch"] = (path, json)
            return {"id": "reply-draft"}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.update_draft(
        drafts.UpdateDraftInput(
            draft_id="reply-draft",
            body="Thanks <team>\nI will review this.",
            preserve_history=True,
        )
    )

    assert result.success is True
    assert captured["get"] == (
        "/me/messages/reply-draft",
        {"$select": "id,isDraft,body"},
        {"Prefer": 'outlook.body-content-type="html"'},
    )
    patch_path, patch_json = captured["patch"]
    assert patch_path == "/me/messages/reply-draft"
    assert patch_json["body"]["contentType"] == "HTML"
    content = patch_json["body"]["content"]
    assert content.startswith("<html><body><div>Thanks &lt;team&gt;<br>I will review this.</div>")
    assert "<div id='quoted'>Original history</div>" in content


@pytest.mark.asyncio
async def test_update_draft_preserve_history_fails_closed_for_invalid_draft(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Do not PATCH when Graph does not return an existing draft."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            return {"id": "sent-message", "isDraft": False}

        async def patch(self, path: str, json: dict = None):
            pytest.fail("PATCH must not be called for a non-draft message")

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.update_draft(
        drafts.UpdateDraftInput(
            draft_id="sent-message",
            body="Do not apply",
            preserve_history=True,
        )
    )

    assert result.success is False
    assert result.updated_fields == []
    assert result.error == "Microsoft Graph did not return a valid draft to update."


@pytest.mark.asyncio
async def test_update_draft_updates_recipients(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify update_draft updates toRecipients and ccRecipients when provided."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def patch(self, path: str, json: dict = None):
            captured["json"] = json
            return {"id": "draft-recip"}

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.update_draft(
        drafts.UpdateDraftInput(
            draft_id="draft-recip",
            to=["new@example.com"],
            cc="cc@example.com",
        )
    )

    assert result.success is True
    assert "toRecipients" in result.updated_fields
    assert "ccRecipients" in result.updated_fields
    body = captured["json"]
    assert len(body["toRecipients"]) == 1
    assert body["toRecipients"][0]["emailAddress"]["address"] == "new@example.com"


@pytest.mark.asyncio
async def test_update_draft_no_fields_returns_error(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify update_draft returns an error response when no fields are provided."""
    monkeypatch.setattr(drafts, "get_graph", lambda _profile: None)
    result = await drafts.update_draft(
        drafts.UpdateDraftInput(draft_id="draft-none")
    )

    assert result.success is False
    assert result.error == "No fields to update."
    assert result.updated_fields == []


# ---------------------------------------------------------------------------
# send_draft
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_send_draft_posts_to_send_endpoint(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify send_draft posts to the /send endpoint and returns success."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            return None

    monkeypatch.setattr(drafts, "get_graph", lambda _profile: DummyGraph())
    result = await drafts.send_draft(drafts.SendDraftInput(draft_id="draft-send"))

    assert isinstance(result, SendDraftResponse)
    assert result.success is True
    assert result.draft_id == "draft-send"
    assert captured["path"] == "/me/messages/draft-send/send"


# ---------------------------------------------------------------------------
# ToolRequestModel input validation
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_input_camelcase_normalization() -> None:
    """Verify ToolRequestModel accepts camelCase keys and normalises to snake_case."""
    params = drafts.CreateDraftInput.model_validate(
        {"to": "alice@example.com", "subject": "Hi", "body": "Hello", "bodyType": "html"}
    )
    assert params.body_type == "html"


@pytest.mark.asyncio
async def test_input_stringified_json_payload() -> None:
    """Verify ToolRequestModel parses a JSON-encoded string as the full payload."""
    raw = json.dumps({"to": "alice@example.com", "subject": "Test", "body": "Body text"})
    params = drafts.CreateDraftInput.model_validate(raw)
    assert params.subject == "Test"
    assert params.body == "Body text"
    # A single address string is accepted by to: str | list[str]
    assert "alice@example.com" in params.to


@pytest.mark.asyncio
async def test_list_drafts_input_defaults() -> None:
    """Verify ListDraftsInput default max_results is 10."""
    params = drafts.ListDraftsInput.model_validate({})
    assert params.max_results == 10


@pytest.mark.asyncio
async def test_update_draft_input_camelcase() -> None:
    """Verify UpdateDraftInput accepts camelCase draftId."""
    params = drafts.UpdateDraftInput.model_validate(
        {"draftId": "draft-camel", "subject": "Camel Subject"}
    )
    assert params.draft_id == "draft-camel"
    assert params.subject == "Camel Subject"
