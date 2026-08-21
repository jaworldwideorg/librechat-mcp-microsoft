"""Tests for dual-transport config, scope building, and http-mode wiring."""

from __future__ import annotations

from pathlib import Path

import pytest
from fastmcp.exceptions import ToolError

from mcp_microsoft.config import AppConfig, reset_app_config, validate_http_config
from mcp_microsoft.runtime import reset_runtime_state
from mcp_microsoft.server import build_graph_authorize_scopes

PROFILE_TOOLS = frozenset(
    {
        "list_ms_profiles",
        "add_ms_profile",
        "remove_ms_profile",
        "authenticate_ms_profile",
        "set_default_ms_profile",
    }
)

_HTTP_ENV_VARS = (
    "MCP_TRANSPORT",
    "MCP_HTTP_HOST",
    "MCP_HTTP_PORT",
    "MCP_HTTP_STATELESS",
    "MCP_BASE_URL",
    "MCP_AUTH_CLIENT_ID",
    "MCP_AUTH_CLIENT_SECRET",
    "MCP_AUTH_TENANT_ID",
    "MCP_AUTH_REQUIRED_SCOPE",
    "MCP_RATE_LIMIT_RPS",
    "MCP_STATS_TOKEN",
)

# A dummy but structurally valid Directory (tenant) ID GUID. http mode now
# requires a concrete GUID (fastmcp pins the issuer to a literal URL built from
# it, so pseudo-tenants like "organizations" can never validate a real token).
_TENANT_GUID = "11111111-1111-1111-1111-111111111111"

_FULL_HTTP_ENV = {
    "MCP_TRANSPORT": "http",
    "MCP_BASE_URL": "https://mcp.example.com",
    "MCP_AUTH_CLIENT_ID": "cid",
    "MCP_AUTH_CLIENT_SECRET": "secret",
    "MCP_AUTH_TENANT_ID": _TENANT_GUID,
}


def _set_full_http_env(monkeypatch: pytest.MonkeyPatch, **overrides: str) -> None:
    """Set the minimal env needed to construct a valid http-mode server."""
    for name, value in {**_FULL_HTTP_ENV, **overrides}.items():
        monkeypatch.setenv(name, value)


@pytest.fixture(autouse=True)
def _reset_cached_config() -> None:
    reset_app_config()
    yield
    reset_app_config()


def _clear_http_env(monkeypatch: pytest.MonkeyPatch) -> None:
    for name in _HTTP_ENV_VARS:
        monkeypatch.delenv(name, raising=False)


def _http_config(**overrides: object) -> AppConfig:
    base: dict[str, object] = dict(
        transport="http",
        base_url="https://mcp.example.com",
        auth_client_id="cid",
        auth_client_secret="secret",
        auth_tenant_id=_TENANT_GUID,
    )
    base.update(overrides)
    return AppConfig(**base)  # type: ignore[arg-type]


# --------------------------------------------------------------------------
# config.py — env parsing & normalization
# --------------------------------------------------------------------------


def test_from_env_defaults_to_stdio(monkeypatch: pytest.MonkeyPatch) -> None:
    _clear_http_env(monkeypatch)
    config = AppConfig.from_env()

    assert config.transport == "stdio"
    assert config.http_host == "127.0.0.1"
    assert config.http_port == 8000
    assert config.http_stateless is False
    assert config.base_url == ""
    assert config.auth_client_id == ""
    assert config.auth_required_scope == "mcp-access"


def test_from_env_parses_http_fields_and_normalizes_transport(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.setenv("MCP_TRANSPORT", "  HTTP  ")  # mixed case + whitespace
    monkeypatch.setenv("MCP_HTTP_HOST", "0.0.0.0")
    monkeypatch.setenv("MCP_HTTP_PORT", "9001")
    monkeypatch.setenv("MCP_HTTP_STATELESS", "yes")
    monkeypatch.setenv("MCP_BASE_URL", "https://mcp.example.com")
    monkeypatch.setenv("MCP_AUTH_CLIENT_ID", "cid")
    monkeypatch.setenv("MCP_AUTH_CLIENT_SECRET", "secret")
    monkeypatch.setenv("MCP_AUTH_TENANT_ID", "tenant-guid")
    monkeypatch.setenv("MCP_AUTH_REQUIRED_SCOPE", "custom-scope")

    config = AppConfig.from_env()

    assert config.transport == "http"
    assert config.http_host == "0.0.0.0"
    assert config.http_port == 9001
    assert config.http_stateless is True
    assert config.base_url == "https://mcp.example.com"
    assert config.auth_client_id == "cid"
    assert config.auth_client_secret == "secret"
    assert config.auth_tenant_id == "tenant-guid"
    assert config.auth_required_scope == "custom-scope"


def test_from_env_blank_required_scope_falls_back_to_default(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.setenv("MCP_AUTH_REQUIRED_SCOPE", "   ")
    assert AppConfig.from_env().auth_required_scope == "mcp-access"


def test_from_env_invalid_port_raises_clear_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.setenv("MCP_HTTP_PORT", "not-a-number")
    with pytest.raises(ValueError, match="MCP_HTTP_PORT must be an integer"):
        AppConfig.from_env()


def test_from_env_rate_limit_rps_defaults_to_ten(monkeypatch: pytest.MonkeyPatch) -> None:
    _clear_http_env(monkeypatch)
    assert AppConfig.from_env().rate_limit_rps == 10.0


def test_from_env_parses_rate_limit_rps(monkeypatch: pytest.MonkeyPatch) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.setenv("MCP_RATE_LIMIT_RPS", "25.5")
    assert AppConfig.from_env().rate_limit_rps == 25.5


def test_from_env_invalid_rate_limit_raises_clear_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.setenv("MCP_RATE_LIMIT_RPS", "not-a-number")
    with pytest.raises(ValueError, match="MCP_RATE_LIMIT_RPS must be a number"):
        AppConfig.from_env()


def test_from_env_stats_token_defaults_empty_and_strips(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.delenv("MCP_STATS_TOKEN", raising=False)
    assert AppConfig.from_env().stats_token == ""

    monkeypatch.setenv("MCP_STATS_TOKEN", "  s3cret  ")
    assert AppConfig.from_env().stats_token == "s3cret"


def test_validate_http_config_does_not_require_stats_token() -> None:
    # stats_token is optional observability config; its absence must not block
    # http startup.
    assert validate_http_config(_http_config(stats_token="")) == []


# --------------------------------------------------------------------------
# config.py — validate_http_config
# --------------------------------------------------------------------------


def test_validate_http_config_accepts_complete_config() -> None:
    assert validate_http_config(_http_config()) == []


def test_validate_http_config_reports_every_missing_field() -> None:
    problems = validate_http_config(AppConfig(transport="http"))
    joined = " ".join(problems)

    assert "MCP_BASE_URL" in joined
    assert "MCP_AUTH_CLIENT_ID" in joined
    assert "MCP_AUTH_CLIENT_SECRET" in joined
    assert "MCP_AUTH_TENANT_ID" in joined


def test_validate_http_config_rejects_non_http_base_url() -> None:
    problems = validate_http_config(_http_config(base_url="mcp.example.com"))
    assert any("http:// or https://" in p for p in problems)


def test_validate_http_config_rejects_invalid_transport() -> None:
    problems = validate_http_config(_http_config(transport="grpc"))
    assert any("MCP_TRANSPORT" in p for p in problems)


@pytest.mark.parametrize("bad_port", [0, -1, 65536, 99999])
def test_validate_http_config_rejects_out_of_range_port(bad_port: int) -> None:
    problems = validate_http_config(_http_config(http_port=bad_port))
    assert any("MCP_HTTP_PORT" in p and "65535" in p for p in problems)


@pytest.mark.parametrize("ok_port", [1, 8000, 65535])
def test_validate_http_config_accepts_in_range_port(ok_port: int) -> None:
    assert validate_http_config(_http_config(http_port=ok_port)) == []


@pytest.mark.parametrize(
    "guid",
    [
        "11111111-1111-1111-1111-111111111111",
        "9a8b7c6d-1234-4abc-8def-0123456789ab",
        "9A8B7C6D-1234-4ABC-8DEF-0123456789AB",  # upper-case accepted
        "AaBbCcDd-1122-3344-5566-778899AaBbCc",  # mixed case accepted
    ],
)
def test_validate_http_config_accepts_tenant_guid_any_casing(guid: str) -> None:
    assert validate_http_config(_http_config(auth_tenant_id=guid)) == []


@pytest.mark.parametrize(
    "bad",
    [
        "organizations",
        "common",
        "consumers",
        "contoso.onmicrosoft.com",
        "not-a-guid",
        "11111111-1111-1111-1111-11111111111",  # 11 hex in last group
        "11111111-1111-1111-1111-1111111111111",  # 13 hex in last group
        "gggggggg-1111-1111-1111-111111111111",  # non-hex chars
    ],
)
def test_validate_http_config_rejects_non_guid_tenant(bad: str) -> None:
    problems = validate_http_config(_http_config(auth_tenant_id=bad))
    # Exactly the tenant problem, with the actionable guidance.
    tenant_problems = [p for p in problems if "MCP_AUTH_TENANT_ID" in p]
    assert len(tenant_problems) == 1
    msg = tenant_problems[0]
    assert "tenant GUID" in msg
    assert "Directory (tenant) ID" in msg
    assert "Microsoft Entra ID" in msg
    # Explains WHY (fastmcp's literal issuer pinning).
    assert "iss" in msg and "issuer" in msg


# --------------------------------------------------------------------------
# server.py — build_graph_authorize_scopes helper
# --------------------------------------------------------------------------


def test_graph_scopes_default_only(monkeypatch: pytest.MonkeyPatch) -> None:
    _clear_http_env(monkeypatch)
    # http mode + unset flags -> optional services resolve False (no fallback).
    scopes = build_graph_authorize_scopes(_http_config())

    assert "https://graph.microsoft.com/Mail.ReadWrite" in scopes
    assert "https://graph.microsoft.com/Files.ReadWrite" in scopes
    # No Teams/SharePoint scopes when their flags are unset.
    assert "https://graph.microsoft.com/Team.ReadBasic.All" not in scopes
    assert "https://graph.microsoft.com/Sites.ReadWrite.All" not in scopes
    # offline_access is present, unprefixed, and last.
    assert scopes[-1] == "offline_access"
    assert not any(s.endswith("/offline_access") for s in scopes)


def test_graph_scopes_include_teams_and_sharepoint_when_enabled(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    config = _http_config(
        enable_teams=True,
        enable_sharepoint=True,
        enable_teams_meeting_artifacts=True,
        enable_teams_ai_insights=True,
    )
    scopes = build_graph_authorize_scopes(config)

    assert "https://graph.microsoft.com/Team.ReadBasic.All" in scopes
    assert "https://graph.microsoft.com/OnlineMeetings.ReadWrite" in scopes
    assert "https://graph.microsoft.com/OnlineMeetingTranscript.Read.All" in scopes
    assert "https://graph.microsoft.com/OnlineMeetingAiInsight.Read.All" in scopes
    assert "https://graph.microsoft.com/Sites.ReadWrite.All" in scopes
    assert "offline_access" in scopes


def test_graph_scopes_meeting_artifacts_require_teams(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    # Artifact/insight flags on, but Teams off -> no Teams-derived scopes at all.
    config = _http_config(
        enable_teams=False,
        enable_teams_meeting_artifacts=True,
        enable_teams_ai_insights=True,
    )
    scopes = build_graph_authorize_scopes(config)

    assert "https://graph.microsoft.com/Team.ReadBasic.All" not in scopes
    assert "https://graph.microsoft.com/OnlineMeetingTranscript.Read.All" not in scopes


# --------------------------------------------------------------------------
# server.py — tool-registration matrix per transport
# --------------------------------------------------------------------------


async def _tool_names(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, **env: str) -> set[str]:
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    monkeypatch.setenv("MS365_CLIENT_ID", "boot-client")
    monkeypatch.setenv("MS365_TENANT_ID", "common")
    for name, value in env.items():
        monkeypatch.setenv(name, value)
    reset_runtime_state()

    return {
        tool.name
        for tool in await server_mod.get_mcp_server(reset=True).list_tools(
            run_middleware=False
        )
    }


@pytest.mark.asyncio
async def test_http_mode_omits_profile_tools_but_keeps_core(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _clear_http_env(monkeypatch)
    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID=_TENANT_GUID,
    )

    assert PROFILE_TOOLS.isdisjoint(names), (
        f"profile tools must not be registered in http mode: {PROFILE_TOOLS & names}"
    )
    # Core mail/calendar tools still register.
    assert "send_email" in names
    assert "create_event" in names


@pytest.mark.asyncio
async def test_stdio_mode_registers_profile_tools(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _clear_http_env(monkeypatch)
    names = await _tool_names(monkeypatch, tmp_path)

    assert PROFILE_TOOLS.issubset(names), (
        f"profile tools must be registered in stdio mode; missing: "
        f"{PROFILE_TOOLS - names}"
    )
    assert "send_email" in names


@pytest.mark.asyncio
async def test_http_mode_disables_corporate_profile_teams_fallback(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """A corporate ('common') default profile auto-enables Teams in stdio mode,
    but http mode must resolve optional-service flags from env only.
    """
    _clear_http_env(monkeypatch)
    # MCP_ENABLE_TEAMS/SHAREPOINT deliberately unset -> would fall back to
    # corporate-account detection in stdio mode.
    monkeypatch.delenv("MCP_ENABLE_TEAMS", raising=False)
    monkeypatch.delenv("MCP_ENABLE_SHAREPOINT", raising=False)

    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID=_TENANT_GUID,
    )

    assert "teams_list_joined" not in names
    assert "search_sharepoint_sites" not in names


# --------------------------------------------------------------------------
# sharepoint._get_sharepoint_graph — transport-aware consumer-tenant guard
# --------------------------------------------------------------------------


def _http_only_config() -> AppConfig:
    return _http_config()


def test_sharepoint_http_mode_rejects_consumer_tid(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """http mode: a consumer ``tid`` claim raises the work/school ValueError."""
    import fastmcp.server.dependencies as deps

    from mcp_microsoft.tools import sharepoint

    monkeypatch.setattr(sharepoint, "get_app_config", lambda: _http_only_config())

    class _Tok:
        # The well-known consumer-tenant GUID (== profiles._PERSONAL_TENANT_IDS).
        claims = {"tid": "9188040d-6c67-4c5b-b112-36a304b66dad"}
        token = "assertion"

    monkeypatch.setattr(deps, "get_access_token", lambda: _Tok())

    with pytest.raises(ValueError, match="work or school"):
        sharepoint._get_sharepoint_graph("ignored-in-http")


def test_sharepoint_http_mode_proceeds_when_tid_missing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """http mode: absent ``tid`` claim is not fatal; Graph enforces support later."""
    import fastmcp.server.dependencies as deps

    from mcp_microsoft.tools import sharepoint

    monkeypatch.setattr(sharepoint, "get_app_config", lambda: _http_only_config())

    class _Tok:
        claims: dict = {}  # no tid
        token = "assertion"

    monkeypatch.setattr(deps, "get_access_token", lambda: _Tok())
    # Never consults ProfileManager; returns whatever get_graph yields.
    sentinel = object()
    monkeypatch.setattr(sharepoint, "get_graph", lambda profile: sentinel)

    assert sharepoint._get_sharepoint_graph("ignored-in-http") is sentinel


def test_sharepoint_http_mode_proceeds_when_unauthenticated(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """http mode: no ambient token -> guard is a no-op (get_graph resolves)."""
    import fastmcp.server.dependencies as deps

    from mcp_microsoft.tools import sharepoint

    monkeypatch.setattr(sharepoint, "get_app_config", lambda: _http_only_config())
    monkeypatch.setattr(deps, "get_access_token", lambda: None)
    sentinel = object()
    monkeypatch.setattr(sharepoint, "get_graph", lambda profile: sentinel)

    assert sharepoint._get_sharepoint_graph("ignored-in-http") is sentinel


def test_sharepoint_stdio_mode_still_rejects_consumer_profile(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """stdio path unchanged: consumer profile tenant raises via resolve_profile."""
    from mcp_microsoft.tools import sharepoint

    monkeypatch.setattr(
        sharepoint, "get_app_config", lambda: AppConfig(transport="stdio")
    )

    class _Cfg:
        tenant_id = "consumers"

    class _PM:
        def resolve_profile(self, profile: str | None) -> object:
            return _Cfg()

    monkeypatch.setattr(sharepoint, "get_profile_manager", lambda *a, **k: _PM())

    with pytest.raises(ValueError, match="work or school"):
        sharepoint._get_sharepoint_graph("personal")


# --------------------------------------------------------------------------
# server.main() — transport dispatch
# --------------------------------------------------------------------------


def test_main_aborts_on_unknown_transport(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """An unrecognized MCP_TRANSPORT must abort startup, never silently run stdio."""
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    monkeypatch.setenv("MCP_TRANSPORT", "garbage")
    reset_runtime_state()

    def _fail_if_called(*_args: object, **_kwargs: object) -> None:
        raise AssertionError(
            "get_mcp_server() must not be reached for an unknown transport"
        )

    monkeypatch.setattr(server_mod, "get_mcp_server", _fail_if_called)

    with pytest.raises(SystemExit, match="MCP_TRANSPORT must be 'stdio' or 'http'"):
        server_mod.main()


def test_main_rejects_incomplete_http_config(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Regression guard: http transport with missing auth config still aborts."""
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    monkeypatch.setenv("MCP_TRANSPORT", "http")
    reset_runtime_state()

    monkeypatch.setattr(
        server_mod,
        "get_mcp_server",
        lambda *a, **k: (_ for _ in ()).throw(
            AssertionError("get_mcp_server() must not be reached")
        ),
    )

    with pytest.raises(SystemExit, match="Cannot start http transport"):
        server_mod.main()


def test_main_http_mode_runs_with_http_transport_kwargs(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """http mode dispatches .run(transport='http', host, port, stateless_http)
    with the values drawn from config."""
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    _set_full_http_env(
        monkeypatch,
        MCP_HTTP_HOST="0.0.0.0",
        MCP_HTTP_PORT="9443",
        MCP_HTTP_STATELESS="true",
    )
    reset_runtime_state()

    recorded: dict[str, object] = {}

    class _StubServer:
        def run(self, **kwargs: object) -> None:
            recorded.update(kwargs)

    monkeypatch.setattr(server_mod, "get_mcp_server", lambda *a, **k: _StubServer())

    server_mod.main()

    assert recorded == {
        "transport": "http",
        "host": "0.0.0.0",
        "port": 9443,
        "stateless_http": True,
    }


def test_main_stdio_mode_runs_with_no_kwargs(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """stdio mode calls .run() with no arguments (fastmcp's stdio default)."""
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    reset_runtime_state()

    recorded: dict[str, object] = {}

    class _StubServer:
        def run(self, *args: object, **kwargs: object) -> None:
            recorded["args"] = args
            recorded["kwargs"] = kwargs

    monkeypatch.setattr(server_mod, "get_mcp_server", lambda *a, **k: _StubServer())

    server_mod.main()

    assert recorded == {"args": (), "kwargs": {}}


# --------------------------------------------------------------------------
# create_mcp_server — ProfileManager warm-up is stdio-only
# --------------------------------------------------------------------------


def test_http_mode_skips_profile_manager_warmup(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """http mode never uses ProfileManager, so it must not create its dir."""
    import mcp_microsoft.server as server_mod

    creds_dir = tmp_path / "creds"
    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(creds_dir))
    _clear_http_env(monkeypatch)
    _set_full_http_env(monkeypatch)
    reset_runtime_state()

    server_mod.get_mcp_server(reset=True)

    assert not creds_dir.exists(), (
        "http mode must not warm up ProfileManager (it would needlessly "
        "create the credentials directory)"
    )


def test_stdio_mode_still_warms_up_profile_manager(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Regression guard: stdio mode's eager ProfileManager warm-up is unchanged."""
    import mcp_microsoft.server as server_mod

    creds_dir = tmp_path / "creds"
    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(creds_dir))
    _clear_http_env(monkeypatch)
    reset_runtime_state()

    server_mod.get_mcp_server(reset=True)

    assert creds_dir.exists()


# --------------------------------------------------------------------------
# create_mcp_server — http-only middleware stack (rate limit + audit log)
# and error masking
# --------------------------------------------------------------------------


def test_http_mode_wires_rate_limit_and_audit_middleware(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    import mcp_microsoft.server as server_mod

    from mcp_microsoft.middleware import AuditLoggingMiddleware, UserRateLimitMiddleware

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    _set_full_http_env(monkeypatch)
    reset_runtime_state()

    mcp = server_mod.get_mcp_server(reset=True)

    rate_limiters = [mw for mw in mcp.middleware if isinstance(mw, UserRateLimitMiddleware)]
    audit_loggers = [mw for mw in mcp.middleware if isinstance(mw, AuditLoggingMiddleware)]
    assert len(rate_limiters) == 1
    assert rate_limiters[0].max_requests_per_second == 10.0
    assert len(audit_loggers) == 1
    # Rate limiting runs before audit logging in the stack.
    assert mcp.middleware.index(rate_limiters[0]) < mcp.middleware.index(audit_loggers[0])


def test_http_mode_threads_non_default_rate_limit_rps(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """A non-default MCP_RATE_LIMIT_RPS is threaded into the middleware.

    The sibling test above asserts 10.0, which coincides with the default and
    so cannot prove the value is actually plumbed through; this uses 42.5.
    """
    import mcp_microsoft.server as server_mod

    from mcp_microsoft.middleware import UserRateLimitMiddleware

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    _set_full_http_env(monkeypatch, MCP_RATE_LIMIT_RPS="42.5")
    reset_runtime_state()

    mcp = server_mod.get_mcp_server(reset=True)

    limiters = [mw for mw in mcp.middleware if isinstance(mw, UserRateLimitMiddleware)]
    assert len(limiters) == 1
    assert limiters[0].max_requests_per_second == 42.5
    # Burst mirrors fastmcp's default (2x rps).
    assert limiters[0].burst_capacity == int(42.5 * 2)


def test_http_mode_zero_rate_limit_disables_limiter_but_keeps_audit_log(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    import mcp_microsoft.server as server_mod

    from mcp_microsoft.middleware import AuditLoggingMiddleware, UserRateLimitMiddleware

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    _set_full_http_env(monkeypatch, MCP_RATE_LIMIT_RPS="0")
    reset_runtime_state()

    mcp = server_mod.get_mcp_server(reset=True)

    assert not any(isinstance(mw, UserRateLimitMiddleware) for mw in mcp.middleware)
    assert any(isinstance(mw, AuditLoggingMiddleware) for mw in mcp.middleware)


def test_stdio_mode_has_no_http_only_middleware(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    import mcp_microsoft.server as server_mod

    from mcp_microsoft.middleware import AuditLoggingMiddleware, UserRateLimitMiddleware

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    reset_runtime_state()

    mcp = server_mod.get_mcp_server(reset=True)

    assert not any(isinstance(mw, UserRateLimitMiddleware) for mw in mcp.middleware)
    assert not any(isinstance(mw, AuditLoggingMiddleware) for mw in mcp.middleware)


def test_http_mode_enables_error_masking(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    _set_full_http_env(monkeypatch)
    reset_runtime_state()

    mcp = server_mod.get_mcp_server(reset=True)

    assert mcp._mask_error_details is True


def test_stdio_mode_does_not_force_error_masking(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Regression guard: stdio mode keeps fastmcp's own default (False)."""
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    reset_runtime_state()

    mcp = server_mod.get_mcp_server(reset=True)

    assert mcp._mask_error_details is False


# --------------------------------------------------------------------------
# create_mcp_server — GET /health (http mode only, unauthenticated)
# --------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_http_mode_health_endpoint_is_unauthenticated(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    import httpx

    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    _set_full_http_env(monkeypatch)
    reset_runtime_state()

    mcp = server_mod.get_mcp_server(reset=True)
    app = mcp.http_app()
    transport = httpx.ASGITransport(app=app)

    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.get("/health")

    assert response.status_code == 200
    assert response.json() == {"status": "ok", "transport": "http"}


def test_stdio_mode_registers_no_custom_http_routes(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    _clear_http_env(monkeypatch)
    reset_runtime_state()

    mcp = server_mod.get_mcp_server(reset=True)

    assert mcp._additional_http_routes == []


# --------------------------------------------------------------------------
# Local-disk tool audit — mixed-interface tools reject disk paths in http
# mode; disk-only tools are not registered in http mode.
# --------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_onedrive_upload_file_rejects_local_path_in_http_mode(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    from mcp_microsoft.tools import onedrive

    monkeypatch.setattr(onedrive, "get_app_config", lambda: AppConfig(transport="http"))
    local_file = tmp_path / "f.txt"
    local_file.write_text("hi")

    with pytest.raises(ToolError, match="not available in multi-user http mode"):
        await onedrive.upload_file(onedrive.UploadFileInput(local_path=local_file))


@pytest.mark.asyncio
async def test_onedrive_upload_file_allows_content_base64_in_http_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import base64

    from mcp_microsoft.tools import onedrive

    monkeypatch.setattr(onedrive, "get_app_config", lambda: AppConfig(transport="http"))

    class DummyGraph:
        async def put(self, path: str, content: bytes) -> dict:
            return {"id": "drive-item", "webUrl": "https://example.invalid/file"}

    monkeypatch.setattr(onedrive, "get_graph", lambda _profile: DummyGraph())

    result = await onedrive.upload_file(
        onedrive.UploadFileInput(
            local_path=None,
            filename="report.txt",
            content_base64=base64.b64encode(b"hello").decode("ascii"),
        )
    )

    assert result.success is True


@pytest.mark.asyncio
async def test_onedrive_download_file_not_registered_in_http_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from fastmcp import FastMCP

    from mcp_microsoft.tools import onedrive

    monkeypatch.setattr(onedrive, "get_app_config", lambda: AppConfig(transport="http"))
    mcp = FastMCP("test-server")
    onedrive.register(mcp)

    names = {t.name for t in await mcp.list_tools(run_middleware=False)}
    assert "download_file" not in names
    assert "upload_file" in names  # mixed-interface tools stay registered


@pytest.mark.asyncio
async def test_onedrive_download_file_registered_in_stdio_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from fastmcp import FastMCP

    from mcp_microsoft.tools import onedrive

    monkeypatch.setattr(onedrive, "get_app_config", lambda: AppConfig(transport="stdio"))
    mcp = FastMCP("test-server")
    onedrive.register(mcp)

    names = {t.name for t in await mcp.list_tools(run_middleware=False)}
    assert "download_file" in names


@pytest.mark.asyncio
async def test_download_attachment_http_schema_omits_save_path(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from fastmcp import FastMCP

    from mcp_microsoft.tools import attachments

    monkeypatch.setattr(attachments, "get_app_config", lambda: AppConfig(transport="http"))
    mcp = FastMCP("test-server")
    attachments.register(mcp)

    tool = next(
        tool
        for tool in await mcp.list_tools(run_middleware=False)
        if tool.name == "download_attachment"
    )
    names = {tool.name for tool in await mcp.list_tools(run_middleware=False)}

    assert "read_attachment" in names
    input_properties = tool.parameters["$defs"]["DownloadAttachmentHttpInput"][
        "properties"
    ]
    assert "save_path" not in input_properties
    assert set(input_properties) == {
        "message_id",
        "attachment_id",
        "profile",
    }


@pytest.mark.asyncio
async def test_send_email_http_schema_omits_elicitation_confirmation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from fastmcp import FastMCP

    from mcp_microsoft.tools import mail

    monkeypatch.setattr(mail, "get_app_config", lambda: AppConfig(transport="http"))
    mcp = FastMCP("test-server")
    mail.register(mcp)

    tool = next(
        tool
        for tool in await mcp.list_tools(run_middleware=False)
        if tool.name == "send_email"
    )
    input_properties = tool.parameters["$defs"]["SendEmailHttpInput"]["properties"]

    assert "confirm" not in input_properties


@pytest.mark.asyncio
async def test_send_email_http_posts_without_elicitation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from mcp_microsoft.tools import mail

    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict | None = None):
            captured["path"] = path
            captured["json"] = json

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    result = await mail.send_email_http(
        mail.SendEmailHttpInput(
            to="recipient@example.com",
            subject="Test",
            body="Body",
        )
    )

    assert result.success is True
    assert captured["path"] == "/me/sendMail"


@pytest.mark.asyncio
async def test_download_attachment_http_returns_file_inline(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import base64

    from fastmcp.utilities.types import File

    from mcp_microsoft.tools import attachments

    monkeypatch.setattr(attachments, "get_app_config", lambda: AppConfig(transport="http"))

    class DummyGraph:
        async def get(self, _path: str):
            return {
                "name": "report.txt",
                "contentType": "text/plain",
                "contentBytes": base64.b64encode(b"report").decode("ascii"),
            }

    monkeypatch.setattr(attachments, "get_graph", lambda _profile: DummyGraph())

    result = await attachments.download_attachment_http(
        attachments.DownloadAttachmentHttpInput(
            message_id="message-id",
            attachment_id="attachment-id",
        )
    )

    assert isinstance(result, File)
    assert result.data == b"report"


@pytest.mark.asyncio
async def test_get_contact_photo_rejects_save_path_in_http_mode(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    from mcp_microsoft.tools import contacts

    monkeypatch.setattr(contacts, "get_app_config", lambda: AppConfig(transport="http"))

    with pytest.raises(ValueError, match="not available in multi-user http mode"):
        await contacts.get_contact_photo(
            contacts.GetContactPhotoInput(
                contact_id="c1",
                save_path=str(tmp_path / "photo.jpg"),
            )
        )


@pytest.mark.asyncio
async def test_upload_to_site_rejects_local_path_in_http_mode(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    from mcp_microsoft.tools import sharepoint

    monkeypatch.setattr(sharepoint, "get_app_config", lambda: AppConfig(transport="http"))
    local_file = tmp_path / "f.txt"
    local_file.write_text("hi")

    with pytest.raises(ToolError, match="not available in multi-user http mode"):
        await sharepoint.upload_to_site(
            sharepoint.UploadToSiteInput(
                site_id="site",
                drive_id="drive",
                local_path=local_file,
            )
        )


@pytest.mark.asyncio
async def test_download_from_site_not_registered_in_http_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from fastmcp import FastMCP

    from mcp_microsoft.tools import sharepoint

    monkeypatch.setattr(sharepoint, "get_app_config", lambda: AppConfig(transport="http"))
    mcp = FastMCP("test-server")
    sharepoint.register(mcp)

    names = {t.name for t in await mcp.list_tools(run_middleware=False)}
    assert "download_from_site" not in names
    assert "upload_to_site" in names  # mixed-interface tools stay registered


@pytest.mark.asyncio
async def test_download_from_site_registered_in_stdio_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from fastmcp import FastMCP

    from mcp_microsoft.tools import sharepoint

    monkeypatch.setattr(sharepoint, "get_app_config", lambda: AppConfig(transport="stdio"))
    mcp = FastMCP("test-server")
    sharepoint.register(mcp)

    names = {t.name for t in await mcp.list_tools(run_middleware=False)}
    assert "download_from_site" in names


@pytest.mark.asyncio
async def test_teams_download_recording_not_registered_in_http_mode(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Disk-only Teams recording download stays gated in http mode even when
    the meeting-artifacts flag (explicit-only in all modes) is turned on.
    """
    _clear_http_env(monkeypatch)
    monkeypatch.delenv("MCP_ENABLE_TEAMS_MEETING_ARTIFACTS", raising=False)
    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID=_TENANT_GUID,
        MCP_ENABLE_TEAMS="true",
        MCP_ENABLE_TEAMS_MEETING_ARTIFACTS="true",
    )

    assert "teams_list_meeting_recordings" in names  # sanity: artifacts flag worked
    assert "teams_download_meeting_recording" not in names


@pytest.mark.asyncio
async def test_teams_download_recording_registered_in_stdio_mode(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.delenv("MCP_ENABLE_TEAMS_MEETING_ARTIFACTS", raising=False)
    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_ENABLE_TEAMS="true",
        MCP_ENABLE_TEAMS_MEETING_ARTIFACTS="true",
    )

    assert "teams_download_meeting_recording" in names


# --------------------------------------------------------------------------
# Kill-switch & optional-service gating verified specifically in http mode
# --------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_http_mode_disable_deletion_tools_removes_hard_deletes(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """MCP_DISABLE_DELETION_TOOLS must still hide hard-delete tools in http mode."""
    _clear_http_env(monkeypatch)
    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID=_TENANT_GUID,
        MCP_DISABLE_DELETION_TOOLS="1",
    )

    # remove_ms_profile is already absent in http mode regardless (profile
    # tools aren't registered at all); the other six are the kill switch's
    # http-relevant surface.
    hard_deletes = {
        "delete_email",
        "bulk_delete_emails",
        "delete_event",
        "delete_contact",
        "delete_folder",
        "delete_drive_item",
    }
    assert hard_deletes.isdisjoint(names), (
        f"Expected hard-delete tools hidden in http mode, found: {hard_deletes & names}"
    )
    assert "remove_ms_profile" not in names
    # Recoverable variants remain available.
    assert "trash_email" in names


@pytest.mark.asyncio
async def test_http_mode_disable_deletion_tools_off_registers_hard_deletes(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Default (flag unset): hard-deletes stay registered in http mode too."""
    _clear_http_env(monkeypatch)
    monkeypatch.delenv("MCP_DISABLE_DELETION_TOOLS", raising=False)
    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID=_TENANT_GUID,
    )

    assert "delete_email" in names
    assert "delete_event" in names
    assert "delete_drive_item" in names


@pytest.mark.asyncio
async def test_http_mode_teams_and_sharepoint_stay_unregistered_without_flags(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Teams/SharePoint require explicit env flags in http mode (no fallback)."""
    _clear_http_env(monkeypatch)
    monkeypatch.delenv("MCP_ENABLE_TEAMS", raising=False)
    monkeypatch.delenv("MCP_ENABLE_SHAREPOINT", raising=False)
    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID=_TENANT_GUID,
    )

    assert "teams_list_joined" not in names
    assert "search_sharepoint_sites" not in names


@pytest.mark.asyncio
async def test_http_mode_teams_and_sharepoint_register_with_explicit_flags(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _clear_http_env(monkeypatch)
    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID=_TENANT_GUID,
        MCP_ENABLE_TEAMS="true",
        MCP_ENABLE_SHAREPOINT="true",
    )

    assert "teams_list_joined" in names
    assert "search_sharepoint_sites" in names


# --------------------------------------------------------------------------
# File-upload app registration matrix per transport / explicit flag
# --------------------------------------------------------------------------

# Model-visible tools contributed by the FileUpload provider. ``store_files`` is
# app-visibility only (called by the UI via CallTool, not the model), so it does
# not appear in a model tool listing even when the feature is on.
_UPLOAD_MODEL_TOOLS = frozenset({"file_manager", "list_files", "read_file"})


@pytest.mark.asyncio
async def test_http_mode_registers_file_upload_by_default(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.delenv("MCP_ENABLE_FILE_UPLOAD", raising=False)
    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID=_TENANT_GUID,
    )

    assert _UPLOAD_MODEL_TOOLS.issubset(names), (
        f"file-upload tools missing in http default: {_UPLOAD_MODEL_TOOLS - names}"
    )


@pytest.mark.asyncio
async def test_stdio_mode_omits_file_upload_by_default(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.delenv("MCP_ENABLE_FILE_UPLOAD", raising=False)
    names = await _tool_names(monkeypatch, tmp_path)

    assert _UPLOAD_MODEL_TOOLS.isdisjoint(names), (
        f"file-upload tools must be absent in stdio default: "
        f"{_UPLOAD_MODEL_TOOLS & names}"
    )


@pytest.mark.asyncio
async def test_http_mode_explicit_flag_off_omits_file_upload(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _clear_http_env(monkeypatch)
    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID=_TENANT_GUID,
        MCP_ENABLE_FILE_UPLOAD="false",
    )

    assert _UPLOAD_MODEL_TOOLS.isdisjoint(names)


@pytest.mark.asyncio
async def test_stdio_mode_explicit_flag_on_registers_file_upload(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _clear_http_env(monkeypatch)
    names = await _tool_names(monkeypatch, tmp_path, MCP_ENABLE_FILE_UPLOAD="true")

    assert _UPLOAD_MODEL_TOOLS.issubset(names)
