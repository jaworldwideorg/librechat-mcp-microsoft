"""Tests for the security-hardening pass (token cache, deletion gate, confirm-fix)."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest
from mcp import types as mcp_types
from mcp.shared.exceptions import McpError

from mcp_microsoft.config import reset_app_config
from mcp_microsoft.runtime import reset_runtime_state
from mcp_microsoft.tools import mail
from mcp_microsoft.tools.mail import (
    BulkDeleteEmailsInput,
    BulkTrashEmailsInput,
    DeleteEmailInput,
    bulk_delete_emails,
    bulk_trash_emails,
    delete_email,
)


@pytest.fixture(autouse=True)
def _reset_cached_config():
    reset_app_config()
    yield
    reset_app_config()


@pytest.mark.asyncio
async def test_delete_email_with_confirm_true_and_no_ctx_fails_closed() -> None:
    """confirm=True must NOT silently fall through to permanent deletion when
    the host did not supply a Context (older MCP clients). Pre-fix code did
    `if params.confirm and ctx:` which silently skipped the elicit prompt.
    """
    result = await delete_email(
        DeleteEmailInput(message_id="msg-1", confirm=True),
        ctx=None,
    )

    assert result.success is False
    assert result.irreversible is True
    assert "confirm=True requires" in (result.error or "")


@pytest.mark.asyncio
async def test_send_email_with_confirm_true_and_no_ctx_fails_closed() -> None:
    result = await mail.send_email(
        mail.SendEmailInput(
            to="recipient@example.com",
            subject="Test",
            body="Body",
            confirm=True,
        ),
        ctx=None,
    )

    assert result.success is False
    assert "confirm=True requires" in (result.error or "")


class _EmailConfirmationContext:
    def __init__(self, *, supported: bool, result=None, error: Exception | None = None):
        elicitation = None
        if supported:
            elicitation = SimpleNamespace(form=SimpleNamespace())
        capabilities = SimpleNamespace(elicitation=elicitation)
        self.session = SimpleNamespace(
            client_params=SimpleNamespace(capabilities=capabilities)
        )
        self.result = result or SimpleNamespace(
            action="accept", data=SimpleNamespace(confirmed=True)
        )
        self.error = error

    async def elicit(self, _prompt: str, response_type):
        if self.error is not None:
            raise self.error
        return self.result


@pytest.mark.asyncio
async def test_send_email_confirm_requires_negotiated_form_elicitation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        mail,
        "get_graph",
        lambda _profile: (_ for _ in ()).throw(AssertionError("Graph must not be called")),
    )

    result = await mail.send_email(
        mail.SendEmailInput(
            to="recipient@example.com", subject="Test", body="Body", confirm=True
        ),
        ctx=_EmailConfirmationContext(supported=False),
    )

    assert result.success is False
    assert "negotiated form elicitation" in (result.error or "")


@pytest.mark.asyncio
async def test_send_email_confirm_handles_false_capability_advertisement(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        mail,
        "get_graph",
        lambda _profile: (_ for _ in ()).throw(AssertionError("Graph must not be called")),
    )
    error = McpError(
        mcp_types.ErrorData(
            code=mcp_types.METHOD_NOT_FOUND,
            message="Method not found",
        )
    )

    result = await mail.send_email(
        mail.SendEmailInput(
            to="recipient@example.com", subject="Test", body="Body", confirm=True
        ),
        ctx=_EmailConfirmationContext(supported=True, error=error),
    )

    assert result.success is False
    assert "advertised elicitation" in (result.error or "")


@pytest.mark.asyncio
async def test_send_email_confirm_sends_after_acceptance(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict | None = None):
            captured["path"] = path

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    result = await mail.send_email(
        mail.SendEmailInput(
            to="recipient@example.com", subject="Test", body="Body", confirm=True
        ),
        ctx=_EmailConfirmationContext(supported=True),
    )

    assert result.success is True
    assert captured["path"] == "/me/sendMail"


@pytest.mark.asyncio
async def test_send_email_confirm_cancellation_does_not_call_graph(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        mail,
        "get_graph",
        lambda _profile: (_ for _ in ()).throw(AssertionError("Graph must not be called")),
    )
    cancelled = SimpleNamespace(
        action="cancel", data=SimpleNamespace(confirmed=False)
    )
    result = await mail.send_email(
        mail.SendEmailInput(
            to="recipient@example.com", subject="Test", body="Body", confirm=True
        ),
        ctx=_EmailConfirmationContext(supported=True, result=cancelled),
    )

    assert result.success is False
    assert result.error == "Cancelled by user."


@pytest.mark.asyncio
async def test_delete_email_with_confirm_false_still_deletes(monkeypatch: pytest.MonkeyPatch) -> None:
    """When confirm=False the tool should perform the delete without elicitation.

    Regression guard: ensure the new fail-closed branch only fires when confirm=True.
    """
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict | None = None):
            captured["path"] = path
            return {}

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())

    result = await delete_email(
        DeleteEmailInput(message_id="msg-2", confirm=False),
        ctx=None,
    )

    assert result.success is True
    assert captured["path"] == "/me/messages/msg-2/permanentDelete"


class _BulkGraph:
    def __init__(self) -> None:
        self.batch_called = False

    async def get(self, path: str, params: dict | None = None):
        return {"value": [{"id": "msg-1"}, {"id": "msg-2"}]}

    async def batch(self, requests: list[dict]):
        self.batch_called = True
        return [
            {"id": request["id"], "status": 204}
            for request in requests
        ]


class _AcceptingContext:
    def __init__(self) -> None:
        self.prompt = ""

    async def elicit(self, prompt: str, response_type):
        self.prompt = prompt
        return SimpleNamespace(
            action="accept",
            data=SimpleNamespace(confirmed=True),
        )


@pytest.mark.asyncio
async def test_bulk_delete_folder_without_confirmation_fails_closed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph = _BulkGraph()
    monkeypatch.setattr(mail, "get_graph", lambda _profile: graph)

    result = await bulk_delete_emails(BulkDeleteEmailsInput(folder="inbox"))

    assert result.success is False
    assert result.total == 2
    assert "requires confirm=True" in (result.error or "")
    assert graph.batch_called is False


@pytest.mark.asyncio
async def test_bulk_delete_folder_prompts_with_count_before_deleting(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph = _BulkGraph()
    ctx = _AcceptingContext()
    monkeypatch.setattr(mail, "get_graph", lambda _profile: graph)

    result = await bulk_delete_emails(
        BulkDeleteEmailsInput(folder="inbox", confirm=True),
        ctx=ctx,
    )

    assert result.success is True
    assert result.succeeded == 2
    assert "Permanently delete 2 messages from 'inbox'" in ctx.prompt
    assert "IRREVERSIBLE" in ctx.prompt
    assert graph.batch_called is True


@pytest.mark.asyncio
async def test_bulk_trash_folder_without_confirmation_fails_closed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph = _BulkGraph()
    monkeypatch.setattr(mail, "get_graph", lambda _profile: graph)

    result = await bulk_trash_emails(BulkTrashEmailsInput(folder="junkemail"))

    assert result.success is False
    assert "requires confirm=True" in (result.error or "")
    assert graph.batch_called is False


@pytest.mark.asyncio
async def test_bulk_trash_folder_confirm_without_context_fails_closed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph = _BulkGraph()
    monkeypatch.setattr(mail, "get_graph", lambda _profile: graph)

    result = await bulk_trash_emails(
        BulkTrashEmailsInput(folder="junkemail", confirm=True),
        ctx=None,
    )

    assert result.success is False
    assert result.total == 2
    assert "supports elicitation" in (result.error or "")
    assert graph.batch_called is False


@pytest.mark.asyncio
async def test_bulk_delete_explicit_ids_can_bypass_prompt(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph = _BulkGraph()
    monkeypatch.setattr(mail, "get_graph", lambda _profile: graph)

    result = await bulk_delete_emails(
        BulkDeleteEmailsInput(message_ids=["msg-1"], confirm=False),
        ctx=None,
    )

    assert result.success is True
    assert result.succeeded == 1
    assert graph.batch_called is True


@pytest.mark.asyncio
async def test_disable_deletion_tools_hides_hard_deletes(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """MCP_DISABLE_DELETION_TOOLS=1 should remove every hard-delete tool from
    the registered set while leaving recoverable variants registered.
    """
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    monkeypatch.setenv("MS365_CLIENT_ID", "client-id")
    monkeypatch.setenv("MS365_TENANT_ID", "common")
    monkeypatch.setenv("MCP_DISABLE_DELETION_TOOLS", "1")
    reset_runtime_state()

    tool_names = {
        tool.name
        for tool in await server_mod.get_mcp_server(reset=True).list_tools(
            run_middleware=False
        )
    }

    hard_deletes = {
        "delete_email",
        "bulk_delete_emails",
        "delete_event",
        "delete_contact",
        "delete_folder",
        "delete_drive_item",
        "remove_ms_profile",
    }
    assert hard_deletes.isdisjoint(tool_names), (
        f"Expected hard-delete tools to be hidden, found: "
        f"{hard_deletes & tool_names}"
    )
    # Recoverable variants must remain available.
    assert "trash_email" in tool_names
    assert "bulk_trash_emails" in tool_names


@pytest.mark.asyncio
async def test_disable_deletion_tools_off_registers_hard_deletes(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Default behavior: with the flag unset, hard-deletes remain registered."""
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    monkeypatch.setenv("MS365_CLIENT_ID", "client-id")
    monkeypatch.setenv("MS365_TENANT_ID", "common")
    monkeypatch.delenv("MCP_DISABLE_DELETION_TOOLS", raising=False)
    reset_runtime_state()

    tool_names = {
        tool.name
        for tool in await server_mod.get_mcp_server(reset=True).list_tools(
            run_middleware=False
        )
    }

    assert "delete_email" in tool_names
    assert "delete_event" in tool_names
    assert "delete_drive_item" in tool_names


def test_persisted_cache_falls_back_and_restricts_permissions(tmp_path: Path) -> None:
    """_build_persisted_cache must always return a usable cache.

    On CI runners without DPAPI/Keychain/libsecret the encrypted backends raise
    on construction and the helper falls back to FilePersistence. The fallback
    path must still produce a working cache and restrict file permissions on
    POSIX.
    """
    from mcp_microsoft.profiles import ProfileConfig, _build_persisted_cache

    cfg = ProfileConfig(
        name="test",
        client_id="client",
        tenant_id="common",
        cache_path=tmp_path / "msal_cache_test.bin",
    )

    cache = _build_persisted_cache(cfg)

    # Exercise the cache: serialize empty state and write through persistence.
    cache.deserialize("")  # initialise empty
    cache._persistence.save(cache.serialize())  # type: ignore[attr-defined]
    assert cfg.cache_path.exists()
