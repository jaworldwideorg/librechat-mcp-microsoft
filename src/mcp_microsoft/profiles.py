"""
Multi-account profile manager for mcp-microsoft.

Manages named profiles, each with its own Azure App Registration (client_id),
tenant, MSAL token cache, and optional scope overrides.  Provides profile-aware
token acquisition and GraphClient instances.

Profiles are stored in profiles.json.  When no profiles exist, the server starts
with zero profiles and the user must call add_ms_profile() to create the first one.
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import msal
from msal_extensions import FilePersistence, PersistedTokenCache

from mcp_microsoft.config import AppConfig, get_app_config
from mcp_microsoft.feature_flags import resolve_optional_service_enabled

logger = logging.getLogger(__name__)


def _cache_path_for(base_dir: Path, name: str) -> Path:
    """Encrypted-cache file path for a profile (kept as .bin to signal opaque)."""
    return base_dir / f"msal_cache_{name}.bin"


def _legacy_cache_path_for(base_dir: Path, name: str) -> Path:
    """Pre-encryption plaintext path; used only for one-time migration."""
    return base_dir / f"msal_cache_{name}.json"


def _restrict_permissions(path: Path, mode: int) -> None:
    """Apply restrictive POSIX permissions; silently skip on Windows / on error.

    Windows relies on the user-profile ACL which already restricts access to
    the current user. ``chmod`` is essentially a no-op there.
    """
    if sys.platform == "win32":
        return
    try:
        os.chmod(path, mode)
    except OSError as exc:
        logger.debug("chmod %o on %s failed: %s", mode, path, exc)


def _build_persisted_cache(cfg: "ProfileConfig") -> msal.SerializableTokenCache:
    """Return an OS-native encrypted MSAL token cache for *cfg*.

    Uses DPAPI on Windows, Keychain on macOS, libsecret on Linux. When no
    encrypted backend is available (typically headless Linux without
    libsecret/dbus) falls back to plaintext ``FilePersistence`` with the
    file restricted to mode 0600 and logs a warning.

    Migrates any legacy plaintext cache (``msal_cache_<name>.json``) into the
    encrypted store on first run, then deletes the legacy file.
    """
    encrypted_path = cfg.cache_path
    persistence: Any = None

    try:
        if sys.platform == "win32":
            from msal_extensions import FilePersistenceWithDataProtection

            persistence = FilePersistenceWithDataProtection(str(encrypted_path))
        elif sys.platform == "darwin":
            from msal_extensions import KeychainPersistence

            persistence = KeychainPersistence(
                str(encrypted_path),
                service_name="mcp-microsoft",
                account_name=cfg.name,
            )
        else:
            from msal_extensions import LibsecretPersistence

            persistence = LibsecretPersistence(
                str(encrypted_path),
                schema_name="mcp-microsoft",
                attributes={"profile": cfg.name},
            )
    except Exception as exc:
        logger.warning(
            "Encrypted token storage unavailable for profile %s (%s). "
            "Falling back to plaintext cache with 0600 permissions.",
            cfg.name,
            exc,
        )
        persistence = FilePersistence(str(encrypted_path))

    cache = PersistedTokenCache(persistence)

    legacy_path = _legacy_cache_path_for(encrypted_path.parent, cfg.name)
    if legacy_path.exists() and not encrypted_path.exists():
        try:
            legacy_content = legacy_path.read_text(encoding="utf-8")
            persistence.save(legacy_content)
            cache.deserialize(legacy_content)
            legacy_path.unlink()
            logger.info(
                "Migrated token cache for profile %s to encrypted store.",
                cfg.name,
            )
        except Exception as exc:
            logger.warning(
                "Failed to migrate legacy token cache for profile %s: %s",
                cfg.name,
                exc,
            )

    if isinstance(persistence, FilePersistence) and encrypted_path.exists():
        _restrict_permissions(encrypted_path, 0o600)

    return cache

# ---------------------------------------------------------------------------
# Default scopes (importable by auth.py facade)
# ---------------------------------------------------------------------------

DEFAULT_SCOPES: list[str] = [
    # Mail
    "Mail.ReadWrite",
    "Mail.Send",
    # Calendar
    "Calendars.ReadWrite",
    "Calendars.Read.Shared",
    # Contacts
    "Contacts.ReadWrite",
    # OneDrive
    "Files.ReadWrite",
]

TEAMS_SCOPES: list[str] = [
    # Teams — channels and channel messages
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "Channel.Create",
    "ChannelMessage.Read.All",
    "ChannelMessage.Send",
    # Teams — chats (1:1 and group)
    "Chat.ReadWrite",
    "Chat.Create",
    # Teams — online meetings
    "OnlineMeetings.ReadWrite",
]

TEAMS_MEETING_ARTIFACT_SCOPES: list[str] = [
    "OnlineMeetingTranscript.Read.All",
    "OnlineMeetingRecording.Read.All",
]

TEAMS_AI_INSIGHT_SCOPES: list[str] = [
    "OnlineMeetingAiInsight.Read.All",
]

SHAREPOINT_SCOPES: list[str] = [
    "Sites.ReadWrite.All",
]


def build_default_scopes(profile_name: str | None = None) -> list[str]:
    """Build the consent scope list for profiles without explicit overrides."""
    scopes = list(DEFAULT_SCOPES)
    teams_enabled = resolve_optional_service_enabled("MCP_ENABLE_TEAMS", profile_name)
    if teams_enabled:
        scopes.extend(TEAMS_SCOPES)
    if teams_enabled and resolve_optional_service_enabled("MCP_ENABLE_TEAMS_MEETING_ARTIFACTS", profile_name):
        scopes.extend(TEAMS_MEETING_ARTIFACT_SCOPES)
    if teams_enabled and resolve_optional_service_enabled("MCP_ENABLE_TEAMS_AI_INSIGHTS", profile_name):
        scopes.extend(TEAMS_AI_INSIGHT_SCOPES)
    if resolve_optional_service_enabled("MCP_ENABLE_SHAREPOINT", profile_name):
        scopes.extend(SHAREPOINT_SCOPES)
    return list(dict.fromkeys(scopes))

# ---------------------------------------------------------------------------
# Profile name validation
# ---------------------------------------------------------------------------

_VALID_NAME_RE = re.compile(r"^[a-zA-Z0-9_-]+$")


def _validate_name(name: str) -> None:
    if not name or not _VALID_NAME_RE.match(name):
        raise ValueError(
            f"Invalid profile name {name!r}. "
            "Use only letters, digits, hyphens, and underscores."
        )


# ---------------------------------------------------------------------------
# ProfileConfig
# ---------------------------------------------------------------------------


@dataclass
class ProfileConfig:
    """Configuration for a single Microsoft 365 profile."""

    name: str
    client_id: str
    tenant_id: str = "common"
    scopes: list[str] | None = None  # None = use DEFAULT_SCOPES
    cache_path: Path = field(default_factory=lambda: Path())

    @property
    def effective_scopes(self) -> list[str]:
        if self.scopes:
            return self.scopes
        return build_default_scopes(self.name)

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"

    def to_dict(self) -> dict[str, Any]:
        """Serialize for profiles.json (excludes computed fields)."""
        d: dict[str, Any] = {
            "client_id": self.client_id,
            "tenant_id": self.tenant_id,
        }
        if self.scopes is not None:
            d["scopes"] = self.scopes
        return d


# ---------------------------------------------------------------------------
# Device-code login session (non-blocking two-phase auth)
# ---------------------------------------------------------------------------


class DeviceLoginSession:
    """A device-code sign-in in progress for one profile.

    Holds the MSAL flow dict (user_code, verification_uri, message,
    expires_at) plus the background thread that polls Entra until the user
    completes sign-in or the flow expires.  The thread writes ``result`` /
    ``error`` when it finishes; the acquired token lands in the profile's
    persisted MSAL cache, so the regular silent path picks it up.
    """

    def __init__(self, profile: str, flow: dict[str, Any]) -> None:
        self.profile = profile
        self.flow = flow
        self.thread: threading.Thread | None = None
        self.result: dict[str, Any] | None = None
        self.error: str | None = None

    @property
    def pending(self) -> bool:
        return self.thread is not None and self.thread.is_alive()

    def cancel(self) -> None:
        """Ask MSAL to abort the polling loop on its next iteration."""
        self.flow["expires_at"] = 0

    def expires_in_seconds(self) -> int:
        return max(0, int(self.flow.get("expires_at", 0) - time.time()))


# ---------------------------------------------------------------------------
# ProfileManager (singleton)
# ---------------------------------------------------------------------------


class ProfileManager:
    """
    Manages multi-account profiles, MSAL authentication, and GraphClient
    instances keyed by profile name.
    """

    def __init__(self, config: AppConfig | None = None) -> None:
        self._config = config or get_app_config()
        self._profiles: dict[str, ProfileConfig] = {}
        self._default_profile: str = ""
        self._base_dir: Path = self._resolve_base_dir(self._config)
        self._msal_apps: dict[str, msal.PublicClientApplication] = {}
        self._graph_clients: dict[str, Any] = {}
        self._last_device_code_message: str | None = None
        self._device_sessions: dict[str, DeviceLoginSession] = {}
        self._device_sessions_lock = threading.Lock()
        self._load()

    # --- Base directory ---------------------------------------------------

    @staticmethod
    def _resolve_base_dir(config: AppConfig) -> Path:
        base = config.credentials_dir
        base.mkdir(parents=True, exist_ok=True)
        _restrict_permissions(base, 0o700)
        return base

    # --- Loading / saving ------------------------------------------------

    @property
    def _config_path(self) -> Path:
        return self._base_dir / "profiles.json"

    def _load(self) -> None:
        """Load profiles from profiles.json, or bootstrap from env vars."""
        if self._config_path.exists():
            self._load_from_file()
        else:
            self._bootstrap_from_env()

    def _load_from_file(self) -> None:
        """Parse profiles.json."""
        try:
            data = json.loads(self._config_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as exc:
            raise RuntimeError(
                f"Failed to read {self._config_path}: {exc}"
            ) from exc

        raw_profiles = data.get("profiles", {})
        if not raw_profiles:
            self._profiles = {}
            self._default_profile = ""
            return

        for name, cfg in raw_profiles.items():
            _validate_name(name)
            self._profiles[name] = ProfileConfig(
                name=name,
                client_id=cfg["client_id"],
                tenant_id=cfg.get("tenant_id", "common"),
                scopes=cfg.get("scopes"),
                cache_path=_cache_path_for(self._base_dir, name),
            )

        self._default_profile = data.get("default_profile", "")
        if self._default_profile not in self._profiles:
            # Fall back to first profile
            self._default_profile = next(iter(self._profiles))

    def _bootstrap_from_env(self) -> None:
        """Auto-create a 'default' profile from env vars (set by MCPB user_config).

        If MS365_CLIENT_ID is set, creates and persists a default profile so the
        user is ready to authenticate immediately. If not set, the server starts
        with zero profiles — the user must call add_ms_profile().
        """
        client_id = self._config.bootstrap_client_id
        if not client_id:
            return
        tenant_id = self._config.bootstrap_tenant_id
        self._profiles["default"] = ProfileConfig(
            name="default",
            client_id=client_id,
            tenant_id=tenant_id,
            cache_path=_cache_path_for(self._base_dir, "default"),
        )
        self._default_profile = "default"
        self._save()

    def _save(self) -> None:
        """Persist current profiles to profiles.json."""
        data = {
            "default_profile": self._default_profile,
            "profiles": {
                name: cfg.to_dict() for name, cfg in self._profiles.items()
            },
        }
        self._config_path.write_text(
            json.dumps(data, indent=2) + "\n", encoding="utf-8"
        )
        _restrict_permissions(self._config_path, 0o600)

    # --- Profile resolution -----------------------------------------------

    def resolve_profile(self, profile: str | None = None) -> ProfileConfig:
        """
        Resolve a profile name to its config.

        None / "" -> default profile.  Raises ValueError if not found.
        """
        if not self._profiles:
            raise ValueError(
                "No profiles configured. Use add_ms_profile to create one."
            )
        name = profile if profile else self._default_profile
        if name not in self._profiles:
            available = ", ".join(sorted(self._profiles.keys()))
            raise ValueError(
                f"Unknown profile {name!r}. Available profiles: {available}"
            )
        return self._profiles[name]

    # --- MSAL auth --------------------------------------------------------

    def _get_msal_app(self, cfg: ProfileConfig) -> msal.PublicClientApplication:
        """Build or return a cached MSAL PublicClientApplication for a profile."""
        if cfg.name in self._msal_apps:
            return self._msal_apps[cfg.name]

        cache = _build_persisted_cache(cfg)

        app = msal.PublicClientApplication(
            client_id=cfg.client_id,
            authority=cfg.authority,
            token_cache=cache,
        )
        self._msal_apps[cfg.name] = app
        return app

    def _save_cache(self, cfg: ProfileConfig, app: msal.PublicClientApplication) -> None:
        """PersistedTokenCache auto-saves through ``modify()``; this is a no-op.

        Retained for call-site stability; remove once no internal callers
        invoke it directly.
        """
        return

    def get_token(self, profile: str | None = None, *, allow_interactive: bool = True) -> str:
        """Acquire a valid access token for the given profile.

        Args:
            profile: Profile name, or None for the default profile.
            allow_interactive: When True (default), falls back to interactive
                browser auth and then device code flow if the silent cache miss.
                Set to False for tool-call paths: a cache miss raises immediately
                with a user-facing message telling Claude to call
                ``authenticate_ms_profile`` instead of blocking the tool call
                on a device code prompt that only appears in the server log.
        """
        cfg = self.resolve_profile(profile)
        app = self._get_msal_app(cfg)
        scopes = cfg.effective_scopes

        # Try silent flow first (fast path — no network round-trip if cached)
        accounts = app.get_accounts()
        result = None
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])

        self._last_device_code_message = None
        if not result:
            if not allow_interactive:
                # Called from a regular tool — refuse to block on device code.
                # Surface a clear error to Claude so the user knows exactly
                # what to do, instead of the tool hanging silently while the
                # sign-in instructions sit unseen in the server log.
                pending = self._pending_device_message(cfg.name)
                if pending:
                    raise RuntimeError(
                        f"Profile {cfg.name!r} has a device-code sign-in waiting "
                        f"for the user to complete it. Show these instructions to "
                        f"the user: {pending}"
                    )
                raise RuntimeError(
                    f"Profile {cfg.name!r} is not authenticated or the token has "
                    "expired. Call authenticate_ms_profile to sign in — it will "
                    "return the sign-in URL and code directly in the chat."
                )

            # Fall back to interactive browser, then device code (headless fallback)
            try:
                result = app.acquire_token_interactive(scopes=scopes)
            except Exception:
                result = None

            if not result or "access_token" not in result:
                # Device code flow — works headless (MCPB, SSH, containers).
                # The message is stored so authenticate_ms_profile can return
                # it to Claude as a structured response field.
                flow = app.initiate_device_flow(scopes=scopes)
                if "user_code" not in flow:
                    raise RuntimeError(
                        f"Could not initiate device code flow for profile {cfg.name!r}: "
                        f"{flow.get('error_description', 'unknown error')}"
                    )
                self._last_device_code_message = flow["message"]
                logger.warning("Interactive auth unavailable. %s", flow["message"])
                result = app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            error = result.get("error", "unknown_error")
            description = result.get("error_description", "No description available.")

            if "AADSTS65001" in description:
                raise RuntimeError(
                    f"Admin consent required for profile {cfg.name!r}: your tenant "
                    f"administrator must pre-approve this application. Share this URL "
                    f"with your IT admin:\n"
                    f"https://login.microsoftonline.com/common/adminconsent"
                    f"?client_id={cfg.client_id}\n"
                    f"Original error: {description}"
                )

            raise RuntimeError(
                f"Token acquisition failed for profile {cfg.name!r} "
                f"[{error}]: {description}"
            )

        self._save_cache(cfg, app)
        return result["access_token"]

    # --- Non-blocking device-code login -----------------------------------

    def _pending_device_message(self, name: str) -> str | None:
        """Return the sign-in instructions of an in-flight device flow, if any."""
        with self._device_sessions_lock:
            session = self._device_sessions.get(name)
        if session is not None and session.pending:
            return session.flow.get("message")
        return None

    def begin_device_login(self, profile: str | None = None) -> dict[str, Any]:
        """Start — or report progress on — a non-blocking device-code sign-in.

        Unlike :meth:`get_token`, this never blocks waiting for the user.  The
        Entra polling loop runs in a daemon thread; the acquired token is
        written to the profile's persisted MSAL cache, after which the normal
        silent path (and every Graph tool) works again.

        Returns a status dict:

        * ``{"status": "authenticated"}`` — a valid token is already available
          (silent cache hit, or a previously started flow just completed).
        * ``{"status": "awaiting_user", "user_code": ..., "verification_uri":
          ..., "message": ..., "expires_in_seconds": ...}`` — the user must
          open the URL and enter the code.  Repeat calls return the same code
          until the flow expires.
        * ``{"status": "error", "error": ...}`` — the previous flow failed or
          expired.  The session is cleared; the next call starts fresh.
        """
        cfg = self.resolve_profile(profile)
        app = self._get_msal_app(cfg)
        scopes = cfg.effective_scopes

        # Report on an existing session first (pending → same code; finished →
        # success or error exactly once, then cleared).
        with self._device_sessions_lock:
            session = self._device_sessions.get(cfg.name)
            if session is not None:
                if session.pending:
                    return {
                        "status": "awaiting_user",
                        "user_code": session.flow.get("user_code", ""),
                        "verification_uri": session.flow.get("verification_uri", ""),
                        "message": session.flow.get("message", ""),
                        "expires_in_seconds": session.expires_in_seconds(),
                    }
                self._device_sessions.pop(cfg.name, None)
                result = session.result or {}
                if "access_token" in result:
                    return {"status": "authenticated"}
                error = (
                    session.error
                    or result.get("error_description")
                    or result.get("error")
                    or "Device-code sign-in expired before it was completed."
                )
                return {"status": "error", "error": error}

        # No session — a silent hit means nothing to do.
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
            if result and "access_token" in result:
                return {"status": "authenticated"}

        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(
                f"Could not initiate device code flow for profile {cfg.name!r}: "
                f"{flow.get('error_description', 'unknown error')}"
            )

        session = DeviceLoginSession(profile=cfg.name, flow=flow)

        def _poll() -> None:
            try:
                session.result = app.acquire_token_by_device_flow(flow)
            except Exception as exc:
                session.error = str(exc)

        session.thread = threading.Thread(
            target=_poll, name=f"msal-device-login-{cfg.name}", daemon=True
        )

        with self._device_sessions_lock:
            existing = self._device_sessions.get(cfg.name)
            if existing is not None and existing.pending:
                # A concurrent call won the race while we were initiating —
                # abandon our flow and reuse the existing one.
                session.cancel()
                return {
                    "status": "awaiting_user",
                    "user_code": existing.flow.get("user_code", ""),
                    "verification_uri": existing.flow.get("verification_uri", ""),
                    "message": existing.flow.get("message", ""),
                    "expires_in_seconds": existing.expires_in_seconds(),
                }
            self._device_sessions[cfg.name] = session

        session.thread.start()
        self._last_device_code_message = flow.get("message")
        logger.info(
            "Device code sign-in started for profile %s. %s",
            cfg.name,
            flow.get("message", ""),
        )
        return {
            "status": "awaiting_user",
            "user_code": flow.get("user_code", ""),
            "verification_uri": flow.get("verification_uri", ""),
            "message": flow.get("message", ""),
            "expires_in_seconds": session.expires_in_seconds(),
        }

    def get_headers(self, profile: str | None = None) -> dict[str, str]:
        """Return authenticated HTTP headers for the given profile."""
        token = self.get_token(profile)
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    # --- GraphClient access -----------------------------------------------

    def get_graph(self, profile: str | None = None) -> Any:
        """
        Return a GraphClient bound to the resolved profile.

        Instances are cached per profile name for reuse.
        Lazy import to avoid circular dependency with graph.py.
        """
        cfg = self.resolve_profile(profile)
        if cfg.name not in self._graph_clients:
            from mcp_microsoft.graph import GraphClient

            self._graph_clients[cfg.name] = GraphClient(profile=cfg.name)
        return self._graph_clients[cfg.name]

    # --- Profile CRUD (used by management tools) --------------------------

    @property
    def default_profile_name(self) -> str:
        return self._default_profile

    @property
    def profiles(self) -> dict[str, ProfileConfig]:
        return dict(self._profiles)

    def add_profile(
        self,
        name: str,
        client_id: str,
        tenant_id: str = "common",
        scopes: list[str] | None = None,
        set_as_default: bool = False,
    ) -> ProfileConfig:
        """Add a new profile and persist to profiles.json."""
        _validate_name(name)
        if name in self._profiles:
            raise ValueError(f"Profile {name!r} already exists.")
        if not client_id or not client_id.strip():
            raise ValueError("client_id is required and cannot be empty.")
        if not tenant_id or not tenant_id.strip():
            raise ValueError("tenant_id is required and cannot be empty.")

        cfg = ProfileConfig(
            name=name,
            client_id=client_id.strip(),
            tenant_id=tenant_id.strip(),
            scopes=scopes,
            cache_path=_cache_path_for(self._base_dir, name),
        )
        self._profiles[name] = cfg

        # If this is the first profile, make it the default
        if set_as_default or not self._default_profile:
            self._default_profile = name

        self._save()
        return cfg

    def remove_profile(self, name: str) -> None:
        """Remove a profile and its cached tokens."""
        if name not in self._profiles:
            raise ValueError(f"Profile {name!r} not found.")
        if len(self._profiles) <= 1:
            raise ValueError("Cannot remove the last remaining profile.")

        cfg = self._profiles.pop(name)
        self._msal_apps.pop(name, None)
        self._graph_clients.pop(name, None)
        with self._device_sessions_lock:
            stale_session = self._device_sessions.pop(name, None)
        if stale_session is not None:
            stale_session.cancel()

        # Remove token cache file(s) — both encrypted and any leftover legacy file
        for stale in (cfg.cache_path, _legacy_cache_path_for(self._base_dir, name)):
            if stale.exists():
                try:
                    stale.unlink()
                except OSError:
                    pass

        # If we removed the default, pick the first remaining
        if self._default_profile == name:
            self._default_profile = next(iter(self._profiles))

        self._save()

    def set_default(self, name: str) -> None:
        """Change the default profile."""
        if name not in self._profiles:
            raise ValueError(f"Profile {name!r} not found.")
        self._default_profile = name
        self._save()

    def is_authenticated(self, profile: str | None = None) -> bool:
        """Check if a profile has cached tokens (without triggering interactive auth)."""
        cfg = self.resolve_profile(profile)
        app = self._get_msal_app(cfg)
        accounts = app.get_accounts()
        if not accounts:
            return False
        result = app.acquire_token_silent(cfg.effective_scopes, account=accounts[0])
        return result is not None and "access_token" in result


# ---------------------------------------------------------------------------
# Corporate-account guard (used by server.py to gate Teams tool registration)
# ---------------------------------------------------------------------------

#: Tenant IDs that are definitively personal / consumer accounts.
_PERSONAL_TENANT_IDS: frozenset[str] = frozenset(
    {
        "consumers",
        # The well-known GUID Microsoft uses as the "consumers" alias
        "9188040d-6c67-4c5b-b112-36a304b66dad",
    }
)


def is_corporate_account(profile: ProfileConfig) -> bool:
    """Return True if *profile* is a work/school (corporate) Microsoft account.

    Teams is an M365 corporate product and its Graph API endpoints always
    return 401 for personal (Outlook.com / Hotmail / Live) accounts.  Use
    this helper to decide whether to expose Teams tools.

    Decision table
    --------------
    tenant_id value          | result | reason
    -------------------------|--------|----------------------------------
    None / ""                | False  | no tenant configured — fail-safe
    "consumers"              | False  | explicit personal-account alias
    "9188040d-..."           | False  | GUID form of "consumers"
    "common"                 | True   | ambiguous sign-in audience — err toward corporate
    "organizations"          | True   | explicit corporate alias
    any GUID / named tenant  | True   | real AAD tenant → corporate
    """
    tid = (profile.tenant_id or "").strip().lower()
    if not tid:
        return False
    return tid not in _PERSONAL_TENANT_IDS


_profile_manager: ProfileManager | None = None


def get_profile_manager(config: AppConfig | None = None) -> ProfileManager:
    global _profile_manager
    if (
        _profile_manager is None
        or (config is not None and _profile_manager._config != config)
    ):
        _profile_manager = ProfileManager(config=config)
    return _profile_manager


def reset_profile_manager() -> None:
    global _profile_manager
    _profile_manager = None
