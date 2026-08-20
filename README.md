# mcp-microsoft

![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)
![License MIT](https://img.shields.io/badge/license-MIT-green)
![MCP](https://img.shields.io/badge/MCP-compatible-purple)
[![Tests](https://github.com/guinacio/mcp-microsoft/actions/workflows/ci.yml/badge.svg)](https://github.com/guinacio/mcp-microsoft/actions/workflows/ci.yml)
[![SafeSkill 91/100](https://img.shields.io/badge/SafeSkill-91%2F100_Verified%20Safe-brightgreen)](https://safeskill.dev/scan/guinacio-mcp-microsoft)

Microsoft 365 MCP server — Mail, Calendar, OneDrive, SharePoint, Contacts, and Teams via the Microsoft Graph API, with multi-account support.

> **Documentation:** full setup & operations reference in [`docs/DEVOPS_GUIDE.md`](docs/DEVOPS_GUIDE.md) (English) · [`docs/DEVOPS_GUIDE.pt-BR.md`](docs/DEVOPS_GUIDE.pt-BR.md) (Guia em Português).

## Overview

`mcp-microsoft` is a [Model Context Protocol](https://modelcontextprotocol.io) server that gives Claude (and any other MCP client) full access to your Microsoft 365 account. It covers six surface areas of the Microsoft Graph API: email, calendar, OneDrive file storage, SharePoint, contacts, and Teams — **96 tools** in stdio mode with every optional service enabled. The exact count depends on your feature flags and, since 0.8.0, on transport mode: the multi-user remote (`http`) mode omits profile-management and local-disk tools, landing at **88 tools** with everything else enabled — see [Remote server — multi-user (Streamable HTTP)](#remote-server--multi-user-streamable-http) below.

The server works with both personal Microsoft accounts (Outlook.com, Live) and enterprise accounts (Azure AD / Entra ID) using a single App Registration. Teams and SharePoint require a work or school account and are gated behind feature flags (`MCP_ENABLE_TEAMS` / `MCP_ENABLE_SHAREPOINT`). On manual installs, they auto-enable for corporate-oriented tenant values (`common`, `organizations`, or a specific tenant ID). In Claude Desktop / MCPB, the installer toggles remain authoritative. You can always override the default with the environment flags to force either service on or off. Teams meeting transcripts/recordings and Copilot AI insights are separate explicit opt-ins so the server does not request those additional scopes unless you enable them.

Multi-account support is a first-class feature. Named profiles let you configure separate client IDs for each account and switch between them on any tool call by passing `profile="work"`. Profiles and MSAL token caches are stored in `~/.microsoft-mcp/` and survive server restarts without re-authentication.

The server ships as an MCPB bundle (`mcp-microsoft.mcpb`) for zero-friction installation through the Claude Desktop Extension installer. It can also be run from source or wired directly into `claude_desktop_config.json`. Built with [FastMCP](https://github.com/jlowin/fastmcp), MSAL, and async httpx.

## Features

### Tools (95 total in stdio mode, all optional services enabled)

#### Mail (25 tools)

- `list_emails` — list messages from any folder with pagination and unread filter
- `read_email` — fetch the full body of a message by ID (supports summary mode)
- `search_emails` — search using Microsoft Graph KQL `$search` syntax (max 25 results)
- `filter_emails` — find emails by sender, recipient, subject, date range, or attachments with full pagination
- `send_email` — compose and send a new message (to/cc/bcc, HTML or plain text)
- `reply_email` — reply or reply-all to an existing message
- `forward_email` — forward a message to one or more recipients
- `mark_as_read` / `mark_as_unread` — toggle read state
- `move_email` — move to any folder by well-known name or folder ID
- `trash_email` — soft-delete to Deleted Items (recoverable)
- `delete_email` — permanently delete a message (irreversible)
- `bulk_move_emails` — move multiple messages to a folder in one operation
- `bulk_trash_emails` — move multiple messages to Deleted Items
- `bulk_delete_emails` — permanently delete multiple messages (irreversible)
- `create_draft` / `get_draft` / `list_drafts` / `update_draft` / `send_draft` — full draft lifecycle
- `list_folders` / `create_folder` / `delete_folder` — manage mailbox folders
- `list_attachments` / `download_attachment` — inspect and save attachments

#### Calendar (10 tools)

- `list_calendars` — enumerate all calendars in the mailbox
- `list_events` — list events from a calendar with optional date filtering
- `list_upcoming_events` — list events using calendarView with recurring-instance expansion
- `get_event` — fetch full event details including attendees, body, and recurrence
- `create_event` — create an event (subject, datetime, timezone, attendees, location, online meeting flag)
- `update_event` / `delete_event` — modify or remove an event
- `rsvp_event` — accept, tentatively accept, or decline an invitation
- `get_free_busy` — check availability for one or more people in a time window
- `find_meeting_times` — get meeting time suggestions for a set of attendees

#### OneDrive (8 tools)

- `list_drive_items` — browse files and folders by path or item ID
- `get_drive_item` — get metadata for a specific file or folder
- `search_drive` — full-text search across OneDrive
- `upload_file` — upload a local file (auto-switches to resumable upload for files over 4 MB)
- `download_file` — download a file to a local path
- `create_drive_folder` — create a new folder at any path
- `move_or_copy_item` — move or copy items within OneDrive
- `delete_drive_item` — delete a file or folder (moves to recycle bin)

#### SharePoint (13 tools)

> SharePoint tools require a work or school account (Azure AD / Entra ID). They are not available for personal Outlook.com / Live accounts, which do not support the `Sites.ReadWrite.All` Graph permission. `Sites.ReadWrite.All` requires one-time admin consent in enterprise tenants.

- `search_content` — tenant-wide full-text search across content via the Microsoft Search API (KQL queries over files, list items, sites, messages, and events)
- `search_sharepoint_sites` — search or list SharePoint sites the user can access
- `get_sharepoint_site` — get details of a specific site
- `list_site_libraries` — list document libraries in a site
- `list_site_files` / `get_site_file` — browse files in a document library
- `upload_to_site` / `download_from_site` — transfer files to/from SharePoint
- `list_site_lists` — list all SharePoint lists in a site
- `get_list_items` / `create_list_item` / `update_list_item` / `delete_list_item` — manage list records

#### Contacts (8 tools)

- `list_contacts` — list contacts with optional folder scope and field selection
- `get_contact` — retrieve a single contact by ID
- `create_contact` — create a new contact with name, email, phone, org, and notes
- `update_contact` — update any subset of contact fields
- `delete_contact` — delete a contact by ID (irreversible)
- `list_contact_folders` — enumerate contact folders in the mailbox
- `search_contacts` — search contacts by display name or email
- `get_contact_photo` — fetch a contact's profile photo as base64 or save to disk

#### Teams (25 tools)

> Teams tools require a work or school account (Azure AD / Entra ID) with a specific tenant ID. They are not available for personal Outlook.com / Live accounts.

- `teams_list_joined` / `teams_get` — list and inspect teams
- `teams_list_channels` / `teams_get_channel` / `teams_create_channel` — manage channels
- `teams_list_channel_messages` / `teams_get_channel_message` / `teams_send_channel_message` — read and post channel messages
- `teams_reply_to_channel_message` / `teams_list_message_replies` — manage channel threads
- `teams_list_chats` / `teams_get_chat` / `teams_list_chat_messages` / `teams_send_chat_message` / `teams_create_chat` — 1:1 and group chats
- `teams_create_meeting` / `teams_get_meeting` / `teams_find_meeting_by_url` / `teams_list_meetings` — online meetings with join URLs
- `teams_list_meeting_transcripts` / `teams_get_meeting_transcript` — Teams meeting transcript metadata and VTT content
- `teams_list_meeting_recordings` / `teams_download_meeting_recording` — Teams meeting recording metadata and downloads
- `teams_list_meeting_ai_insights` / `teams_get_meeting_ai_insight` — Copilot meeting recap/insight metadata and full detail

> Teams meeting transcripts and recordings require explicit opt-in via `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` plus the delegated permissions `OnlineMeetingTranscript.Read.All` and `OnlineMeetingRecording.Read.All`. Teams AI insights require explicit opt-in via `MCP_ENABLE_TEAMS_AI_INSIGHTS`, the delegated permission `OnlineMeetingAiInsight.Read.All`, and Microsoft 365 Copilot licensing.

#### Profile Management (5 tools)

- `list_ms_profiles` — list all configured profiles and which is the default
- `add_ms_profile` — add a new account (name, client_id, tenant_id)
- `remove_ms_profile` — remove a profile and delete its cached tokens
- `authenticate_ms_profile` — start or check a sign-in for a profile; returns a device code + URL to show the user in the chat (never blocks waiting for them)
- `set_default_ms_profile` — change which profile is used when none is specified

#### Service utilities (1 tool)

- `list_enabled_services` — report which optional service groups (SharePoint, Teams, Teams meeting artifacts, Teams AI insights) are currently enabled

## Installation

### Option A: Claude Desktop Extension (MCPB) — Recommended

```bash
npx @anthropic-ai/mcpb install mcp-microsoft-0.8.0.mcpb
```

The installer prompts for your Azure App Registration details (see [Azure Setup](#azure-setup)):

| Prompt | Description |
|---|---|
| **Azure Client ID** | Application (client) ID from your App Registration |
| **Tenant ID** | `common` for personal + work, `consumers` for personal only, or your org's tenant ID/domain |
| **Credentials Directory** | Optional. Defaults to `~/.microsoft-mcp/` |
| **Enable Teams Tools** | Toggle Teams tools on/off (requires work/school account) |
| **Enable Teams Meeting Artifacts** | Toggle Teams transcript/recording tools on/off (requires work/school account and extra Graph permissions) |
| **Enable Teams AI Insights** | Toggle Teams Copilot AI insight tools on/off (requires work/school account, extra Graph permissions, and Copilot licensing) |
| **Enable SharePoint Tools** | Toggle SharePoint tools on/off (requires work/school account) |

A `default` profile is created automatically from these values.

### Option B: From Source

```bash
git clone https://github.com/guinacio/mcp-microsoft.git
cd mcp-microsoft
uv sync
export MS365_CLIENT_ID=your-client-id
export MS365_TENANT_ID=common
# Optional overrides:
# export MCP_ENABLE_SHAREPOINT=false
# export MCP_ENABLE_TEAMS=false
# export MCP_ENABLE_TEAMS_MEETING_ARTIFACTS=true
# export MCP_ENABLE_TEAMS_AI_INSIGHTS=true
uv run mcp-microsoft
```

### Option C: Add to claude_desktop_config.json

```json
{
  "mcpServers": {
    "mcp-microsoft": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/mcp-microsoft", "mcp-microsoft"],
      "env": {
        "MS365_CLIENT_ID": "your-client-id",
        "MS365_TENANT_ID": "common"
      }
    }
  }
}
```

## Remote server — multi-user (Streamable HTTP)

Everything above runs `mcp-microsoft` as a single-user **stdio** server: one process, one local profile, launched directly by your MCP client. As of 0.8.0 the server also supports a second, mutually exclusive mode — a shared **remote server** that any number of people can connect to over the network, where each person signs in with their own Microsoft account and every Graph call runs under their own delegated identity. No profile configuration, no shared credentials: identity comes from a per-user **On-Behalf-Of (OBO)** token exchange derived from the OAuth token each client presents.

Under the hood: Streamable HTTP per the MCP **2025-11-25** spec, served at `/mcp`; authentication via FastMCP's `AzureProvider`, which implements the OAuth-proxy pattern Microsoft Entra ID needs since Entra doesn't support Dynamic Client Registration.

### Quickstart

**From source:**

```bash
export MCP_TRANSPORT=http
export MCP_BASE_URL=https://mcp.example.com   # your public HTTPS URL (behind a reverse proxy)
export MCP_AUTH_CLIENT_ID=your-confidential-client-id
export MCP_AUTH_CLIENT_SECRET=your-client-secret
export MCP_AUTH_TENANT_ID=your-directory-tenant-id-guid   # Entra > Overview > Directory (tenant) ID
uv run mcp-microsoft
```

**With Docker:**

```bash
cp .env.template .env
# fill in the "Remote server (http) mode" section of .env, then:
docker compose up -d
curl http://localhost:8000/health
```

See [`docs/azure-setup.md`](docs/azure-setup.md#app-registration-for-the-remote-http-server) for the Azure App Registration this mode requires — it is a **separate, confidential-client registration**, distinct from the public-client one used for stdio installs.

### How MCP clients connect

Point an MCP client with OAuth support at `https://your-host/mcp`. The client discovers and drives the OAuth flow itself (the server advertises `/.well-known/oauth-protected-resource` per RFC 9728, as MCP 2025-11-25 authorization requires); when the user completes Microsoft sign-in, the client starts sending an Entra-derived bearer token with every request, and the server exchanges it On-Behalf-Of for a Graph token scoped to that user on each call. Clients without OAuth support (or without Streamable HTTP support) cannot use this mode — use stdio instead.

### http mode vs. stdio — what's different

- **Profile-management tools are not registered.** `add_ms_profile`, `list_ms_profiles`, `remove_ms_profile`, `authenticate_ms_profile`, and `set_default_ms_profile` don't exist in http mode — identity management is a server-operator/stdio concern, not something a remote caller should be able to do.
- **The `profile` argument on every other tool is inert.** It's accepted for API compatibility but silently ignored; identity always comes from the caller's bearer token, never from a name in the request.
- **Local-disk tools are not available.** The server's disk is not the caller's disk. `download_file`, `download_from_site`, and `teams_download_meeting_recording` are not registered at all; `upload_file` / `upload_to_site` reject `local_path` (use `content_base64` or the file-upload UI below), `download_attachment` omits `save_path` and returns the file inline, and `read_attachment` extracts PDF/text content directly for the model.
- **Drag-and-drop file uploads are on by default.** A file-upload app (see [Uploading files](#uploading-files)) lets users drop files straight into the server — bypassing the model's context window — and `upload_file` / `upload_to_site` consume them by name via `uploaded_file`. Off in stdio (local users have `local_path`); toggle with `MCP_ENABLE_FILE_UPLOAD`.
- **Feature flags must be explicit.** `MCP_ENABLE_TEAMS`, `MCP_ENABLE_SHAREPOINT`, `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS`, and `MCP_ENABLE_TEAMS_AI_INSIGHTS` need to be set directly — the corporate-account auto-detect fallback that manual stdio installs get doesn't apply, since there's no single configured profile to inspect.
- **The deletion kill-switch still works.** `MCP_DISABLE_DELETION_TOOLS=true` suppresses the same permanent-delete tools as in stdio mode.
- **Work/school accounts only, single concrete tenant.** `MCP_AUTH_TENANT_ID` must be your directory's **tenant GUID** (Entra → Overview → Directory (tenant) ID). Pseudo-tenants (`organizations`, `common`, `consumers`) and verified domains (`contoso.onmicrosoft.com`) are rejected at startup — fastmcp's `AzureProvider` pins the accepted token issuer to a literal URL built from this value, and real Entra tokens carry the concrete GUID, so a pseudo-tenant/domain never validates. Personal Microsoft accounts (Outlook.com/Live) remain stdio-only — OBO and custom API scopes aren't reliably supported for consumer accounts. Multi-tenant deployments are future work (they need issuer-validation skipping plus per-tenant OBO authority).
- **Unauthenticated `GET /health`** is available for load balancers and container healthchecks; every other route requires a valid bearer token.
- **Rate limiting is on by default.** `MCP_RATE_LIMIT_RPS` (default `10`) caps requests per user (per second) via a bounded per-user (tenant + object id) token bucket; set it to `0` or a negative number to disable.
- **Every tool call is audit-logged**: tool name, caller `oid` and `preferred_username` (from the token's claims), duration, and outcome — never the arguments, results, or the token itself.
- **Error details are masked** in responses sent to remote clients (internal exception messages are logged server-side but not echoed back).

### Security notes

- **TLS is not terminated by this server.** It speaks plain HTTP; put a reverse proxy (Traefik, Caddy, nginx, your cloud load balancer, etc.) in front of it for TLS, and set `MCP_BASE_URL` to the proxy's public HTTPS URL — not this process's bind address. See the commented example in `docker-compose.yml`.
- **`MCP_BASE_URL` must match exactly** (scheme, host, port, path) what MCP clients connect to and what's registered as the Azure redirect URI base. A mismatch breaks the OAuth redirect and the JWT audience/issuer checks.
- **Single concrete work/school tenant only** — see above. `MCP_AUTH_TENANT_ID` must be a tenant GUID; `consumers`, `common`, `organizations`, and domain values are rejected at startup because fastmcp validates the token `iss` against a literal issuer URL built from this value.
- **Single worker only.** The OAuth-proxy client store and the per-user OBO credential cache both live in this process's memory. Running more than one worker/replica splits that state and breaks sessions unpredictably. Horizontal scaling requires wiring fastmcp's external `client_storage` backend (a pluggable key-value store) in place of the in-memory default — not implemented here; treat it as a prerequisite before scaling beyond one process.
- **Rate limiting and audit logging are on by default** in http mode (see above) — there is no equivalent in stdio mode, since stdio has exactly one caller.
- **The built-in rate limit covers MCP tool traffic only.** `MCP_RATE_LIMIT_RPS` throttles authenticated calls to the `/mcp` endpoint via a bounded per-user (`tid`+`oid`) token bucket — bounded memory (LRU-capped, idle buckets pruned) so per-user state can't grow without limit. It does not protect the unauthenticated OAuth endpoints (`/authorize`, `/token`, `/register`, `/auth/callback`). Throttle those at your reverse proxy.
- **Secrets belong in the environment, not in files you commit.** `MCP_AUTH_CLIENT_SECRET` in particular; consider a secrets manager (Azure Key Vault, etc.) that injects it as an env var at deploy time rather than storing it in `.env` on disk long-term.

### Observability

http mode can expose lightweight, in-process traffic/usage metrics for DevOps. It is **off by default** and turns on only when you set `MCP_STATS_TOKEN` to a non-empty secret. When enabled, three routes are served by the same server (all requiring the token; `/health` is unaffected and stays open):

| Route | Returns | Use |
|---|---|---|
| `GET /metrics` | Prometheus text exposition (`text/plain; version=0.0.4`) | Scrape target for Prometheus/Grafana |
| `GET /stats` | JSON snapshot (server uptime/totals, per-minute traffic, per-tool latency p50/p95/avg, per-user activity) | Programmatic dashboards, ad-hoc `curl` |
| `GET /dashboard` | A single self-contained HTML page (inline CSS/JS, no external requests) that polls `/stats` every 10s | Eyeball it in a browser |

- **Auth**: send either `Authorization: Bearer <MCP_STATS_TOKEN>` or HTTP Basic with any username and the token as the password (so a browser can open `/dashboard` with its native login prompt). The comparison is timing-safe; the token is never logged.
- **Prometheus scrape**: point your scraper at `https://your-host/metrics` with a `bearer_token` (or `basic_auth` password) equal to `MCP_STATS_TOKEN`. Emitted metrics: `mcp_uptime_seconds`, `mcp_calls_total`, `mcp_errors_total`, `mcp_users_tracked`, `mcp_users_evicted_total`, and per-tool `mcp_tool_calls_total`/`mcp_tool_errors_total`/`mcp_tool_duration_ms{stat="p50|p95|avg"}`. There are deliberately **no per-user label series** (that would be unbounded-cardinality — an anti-pattern); per-user detail lives in `/stats` and `/dashboard` instead.
- **Metrics are in-memory and reset on restart** — there is no persistence. They cover only the single worker the process runs (consistent with the single-worker constraint above).
- **Protect it at the proxy too.** The token is the only gate; treat these routes as sensitive operational data and additionally restrict them (IP allowlist / separate auth) at your reverse proxy if the server is internet-facing. Leaving `MCP_STATS_TOKEN` unset disables the routes entirely.

### Uploading files

In http mode there's no shared disk between you and the server, so passing a file as base64 would push its whole content through the model's context window. The **file-upload app** (FastMCP's `FileUpload` provider, `fastmcp[apps]`) avoids that: it adds a drag-and-drop UI whose files go **straight to the server**, into a per-user upload area, without touching the model context. `upload_file` (OneDrive) and `upload_to_site` (SharePoint) then take an `uploaded_file` name (mutually exclusive with `local_path` / `content_base64`) and stream those bytes into the same Graph upload path.

- **On by default in http mode, off in stdio** (local users already have `local_path`). Override with `MCP_ENABLE_FILE_UPLOAD` (explicit value wins).
- **Requires an MCP-Apps-capable client** (e.g. Claude Desktop) to show the drop-zone UI. The `list_files` / `read_file` tools are model-visible regardless.
- **Bounded, in-memory, per-user**: max `MCP_UPLOAD_MAX_MB` (default 10 MB) per file, 20 files and 100 MB (decoded) per user, a global `MCP_UPLOAD_GLOBAL_BUDGET_MB` (default 1024 MB) encoded budget across all users, 1000 users tracked (LRU-evicted), 2 h idle TTL. Uploads are lost on restart and file content is never logged.

## Azure Setup

You need an Azure App Registration to get a `client_id`. This is a one-time step.

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**.
2. Name it anything (e.g., `mcp-microsoft`).
3. Under **Supported account types**, choose based on your use case:
   - *Personal Microsoft accounts only* — Outlook.com / Live users
   - *Accounts in any organizational directory and personal Microsoft accounts* — personal and work
4. Under **Redirect URI**, select **Mobile and desktop applications** and enter `http://localhost`.
5. Under **Authentication**, enable **Allow public client flows** (required for the interactive loopback OAuth flow — no client secret needed).
6. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** and add:
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `Calendars.Read.Shared` *(required by `find_meeting_times`)*
   - `Contacts.ReadWrite`
   - `Files.ReadWrite`
   - `offline_access` *(usually pre-added)*
7. If you want to enable **SharePoint** tools, also add `Sites.ReadWrite.All`.
8. If you want to enable **Teams** tools, also add:
   - `Team.ReadBasic.All`
   - `Channel.ReadBasic.All`
   - `Channel.Create`
   - `ChannelMessage.Read.All`
   - `ChannelMessage.Send`
   - `Chat.ReadWrite`
   - `Chat.Create`
   - `OnlineMeetings.ReadWrite`
9. If you want to enable **Teams meeting transcripts and recordings**, also add:
   - `OnlineMeetingTranscript.Read.All`
   - `OnlineMeetingRecording.Read.All`
10. If you want to enable **Teams Copilot AI insights**, also add:
   - `OnlineMeetingAiInsight.Read.All`
   All users of applications that call this API need a Microsoft 365 Copilot license.
11. For admin-restricted permissions such as `Sites.ReadWrite.All`, click **Grant admin consent**. Your IT administrator must approve this once per tenant.
12. From the **Overview** page, copy the **Application (client) ID** and, if targeting a specific tenant, the **Directory (tenant) ID**.

For a detailed walkthrough with screenshots, see [`docs/azure-setup.md`](docs/azure-setup.md).

## Account Type Compatibility

| Module | Personal (Outlook.com/Live) | Work/School (Azure AD/Entra) |
|---|---|---|
| Mail | Yes | Yes |
| Calendar | Yes, except availability and meeting-time suggestions | Yes |
| OneDrive | Yes | Yes |
| Contacts | Yes | Yes |
| SharePoint | No | Yes (admin consent required) |
| Teams | No | Yes |

Personal accounts use `tenant_id=consumers`. For access to both personal and work accounts via a single profile, use `tenant_id=common`. Teams transcript/recording and AI-insight APIs remain work-account-only and are additional explicit opt-ins on top of the base Teams tools.

## Profile Management

The server supports multiple Microsoft 365 accounts as named profiles. Each profile has its own `client_id`, `tenant_id`, and MSAL token cache.

**Bootstrap:** On first start, if `MS365_CLIENT_ID` is set (via the MCPB installer or environment variable), a `default` profile is created and persisted to `profiles.json` automatically. If the variable is not set, the server starts with zero profiles and you must call `add_ms_profile`.

**Add accounts:**

```
add_ms_profile(name="personal", client_id="...", tenant_id="consumers")
add_ms_profile(name="work", client_id="...", tenant_id="mycompany.onmicrosoft.com")
```

**Use a specific profile** on any tool call:

```
list_emails(folder="Inbox", profile="work")
search_drive(query="Q1 report", profile="personal")
```

**Authenticate** (returns a sign-in URL + device code in the chat; open the URL, enter the code, then call the tool again to confirm):

```
authenticate_ms_profile(profile="work")
```

**List profiles:**

```
list_ms_profiles()
```

**Change the default:**

```
set_default_ms_profile(profile="work")
```

Profiles are stored in `~/.microsoft-mcp/profiles.json`. Token caches are stored as `~/.microsoft-mcp/msal_cache_{name}.bin` and encrypted at rest via OS-native APIs (DPAPI on Windows, Keychain on macOS, libsecret on Linux) using [msal-extensions](https://github.com/AzureAD/microsoft-authentication-extensions-for-python). Legacy plaintext caches from older versions are migrated automatically on first run. After the first interactive login, MSAL handles token refresh silently.

> **Security note:** `profiles.json` contains the client/tenant IDs only — it has no secrets but is restricted to mode `0600` on POSIX. The `msal_cache_*.bin` files hold encrypted refresh tokens; do not commit either to version control. `MS365_CLIENT_ID` is not a secret and can be committed.
>
> **Disabling destructive tools:** set `MCP_DISABLE_DELETION_TOOLS=1` to suppress registration of all permanent-delete tools (`delete_email`, `bulk_delete_emails`, `delete_event`, `delete_contact`, `delete_folder`, `delete_drive_item`, `delete_list_item`, `remove_ms_profile`). Recoverable variants such as `trash_email` remain available.

## Configuration

| Variable | Required | Default | Description |
|---|---|---|---|
| `MS365_CLIENT_ID` | Yes (for bootstrap) | — | Azure App Registration client ID for the default profile |
| `MS365_TENANT_ID` | No | `common` | Tenant ID for the default profile |
| `MS365_CREDENTIALS_DIR` | No | `~/.microsoft-mcp/` | Directory for `profiles.json` and token caches |
| `MCP_ENABLE_SHAREPOINT` | No | auto-detect | Set to `true` to force-enable SharePoint tools or `false` to force-disable them |
| `MCP_ENABLE_TEAMS` | No | auto-detect | Set to `true` to force-enable Teams tools or `false` to force-disable them |
| `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` | No | `false` | Set to `true` to register Teams transcript and recording tools and request the related delegated scopes |
| `MCP_ENABLE_TEAMS_AI_INSIGHTS` | No | `false` | Set to `true` to register Teams Copilot AI insight tools and request `OnlineMeetingAiInsight.Read.All` |
| `MCP_DISABLE_DELETION_TOOLS` | No | `false` | Set to `true` to suppress registration of all permanent-delete tools (recoverable trash variants remain) |

`MS365_CLIENT_ID`, `MS365_TENANT_ID`, and `MS365_CREDENTIALS_DIR` are only used to bootstrap the `default` profile on first run. `MCP_ENABLE_TEAMS` and `MCP_ENABLE_SHAREPOINT` participate in the existing auto-detection behavior; `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` and `MCP_ENABLE_TEAMS_AI_INSIGHTS` are explicit opt-ins only.

## Development

```bash
# Install dependencies
uv sync

# Start the MCP server (stdio mode)
uv run mcp-microsoft

# Run the test suite
uv run pytest -q

# Rebuild the MCPB bundle
npx @anthropic-ai/mcpb pack
```

Tool implementations are organized by surface area under `src/mcp_microsoft/tools/`: `mail.py`, `drafts.py`, `folders.py`, `attachments.py`, `calendar.py`, `onedrive.py`, `sharepoint.py`, `contacts.py`, `teams.py`, `profiles.py`. Authentication is handled by [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python) with a per-profile serializable token cache. HTTP calls go through a shared async `httpx` client initialized at server startup.

## License

MIT
