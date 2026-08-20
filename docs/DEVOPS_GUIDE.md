# mcp-microsoft — DevOps & Operations Guide

A complete setup and operations reference for `mcp-microsoft` **v0.8.0**, covering both
deployment modes: the single-user **stdio** server (Claude Desktop / MCPB / from source) and the
multi-user **remote HTTP** server (`MCP_TRANSPORT=http`). It is written to be self-contained — you
should not need to read the source to stand up either mode.

> **Looking for the quick version?** The project [`README`](../README.md) has condensed quickstarts.
> This guide is the long-form operator reference: every environment variable, every Azure step,
> the security model, and troubleshooting for the errors you will actually hit.

**Português:** a Portuguese-language edition of this guide lives at
[`docs/DEVOPS_GUIDE.pt-BR.md`](DEVOPS_GUIDE.pt-BR.md) ("Guia em Português").

---

## Table of contents

1. [Overview](#1-overview)
2. [Requirements](#2-requirements)
3. [Understanding the credentials](#3-understanding-the-credentials)
4. [Azure App Registration](#4-azure-app-registration)
5. [Setup: stdio (single user)](#5-setup-stdio-single-user)
6. [Setup: http (multi-user server)](#6-setup-http-multi-user-server)
7. [Observability](#7-observability)
8. [Security checklist](#8-security-checklist)
9. [Troubleshooting](#9-troubleshooting)
10. [Limits & FAQ](#10-limits--faq)
11. [Appendix: env vars & tool inventory](#11-appendix-env-vars--tool-inventory)

---

## 1. Overview

`mcp-microsoft` is a [Model Context Protocol](https://modelcontextprotocol.io) server that exposes
Microsoft 365 — Mail, Calendar, OneDrive, SharePoint, Contacts, and Teams — to any MCP client
(Claude Desktop, Claude Code, VS Code, etc.) through the **Microsoft Graph API**. It works with both
personal Microsoft accounts (Outlook.com / Live) and enterprise accounts (Microsoft Entra ID /
Azure AD).

As of 0.8.0 the server runs in one of **two mutually exclusive modes**, chosen at startup by the
`MCP_TRANSPORT` environment variable:

| | **stdio** (default) | **http** (remote, multi-user) |
|---|---|---|
| `MCP_TRANSPORT` | `stdio` | `http` |
| Transport | stdio (local subprocess) | Streamable HTTP (MCP spec 2025-11-25), served at `/mcp` |
| Who runs it | Launched by your MCP client on your machine | A shared server you host behind a reverse proxy |
| Identity | Named local **profiles**; interactive/device-code OAuth; encrypted disk token cache | Each user signs in with **their own** Entra account; per-request bearer token |
| How Graph calls are authorized | MSAL public-client token for the active profile | Per-user **On-Behalf-Of (OBO)** token exchange from the caller's bearer token |
| Azure app type | **Public** client (Mobile & desktop platform) | **Confidential** client (Web platform + secret) |
| Accounts | Personal **and** work/school | Work/school **only**, single concrete tenant |
| Tool count (all optional services on) | **96** | **88** (profile-management + local-disk tools omitted) |

**Architecture in five lines:**

1. Built on **FastMCP** (`fastmcp[azure]>=3.4.4`), MSAL, and async `httpx`.
2. In **stdio** mode a `ProfileManager` holds named profiles; each acquires a Graph token via MSAL
   (interactive browser, falling back to device code) and caches it, OS-encrypted, on local disk.
3. In **http** mode FastMCP's **`AzureProvider`** implements the OAuth-proxy pattern Entra requires
   (Entra has no Dynamic Client Registration), advertising `/.well-known/oauth-protected-resource`
   (RFC 9728) so MCP clients can discover and drive the OAuth flow.
4. Every Graph call in http mode runs under the caller's identity via a per-user **On-Behalf-Of**
   exchange (`azure.identity` `OnBehalfOfCredential`) — no shared service credentials, no
   server-side profiles.
5. All tool bodies funnel through one seam (`GraphClient`), so the two identity paths share the same
   Graph tool implementations.

**Which mode should you choose?**

- **stdio** — a single person on their own laptop, personal or work account, no server to operate.
  This is the default and the MCPB / Claude Desktop experience.
- **http** — you want to offer Microsoft 365 tools to many users from one hosted service, each with
  their own delegated identity and no local install. Requires a work/school tenant, a public HTTPS
  endpoint, and a confidential Azure App Registration.

---

## 2. Requirements

### 2.1 Common to both modes

- **Python `>=3.11`** (the Docker image ships 3.12).
- **[uv](https://docs.astral.sh/uv/)** for dependency management and running:
  ```bash
  # macOS / Linux
  curl -LsSf https://astral.sh/uv/install.sh | sh
  # Windows (PowerShell)
  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
  ```
- **Outbound network** to `https://graph.microsoft.com` (Graph API) and
  `https://login.microsoftonline.com` (Entra OAuth / token endpoints).
- A **Microsoft account** — personal (Outlook.com / Live) for stdio, or an **Entra tenant** where
  you can create App Registrations for either mode.

### 2.2 stdio / MCPB mode

- **Claude Desktop** (or any MCP client that launches stdio servers). Two install paths:
  - **MCPB bundle** — double-click the `.mcpb` file to install through Claude Desktop's Extension
    installer, **or** install from a shell:
    ```bash
    npx @anthropic-ai/mcpb install mcp-microsoft-0.8.0.mcpb
    ```
    **Node.js is only needed for the `npx` route** — the double-click path needs nothing extra.
  - **From source** — clone the repo and run via `uv` (see [§5.2](#52-from-source)). Works with any
    MCP client that accepts a stdio command.
- **OS token-cache encryption backends** (used automatically by `msal-extensions` to encrypt the
  refresh-token cache at rest):

  | OS | Backend | Notes |
  |---|---|---|
  | Windows | **DPAPI** | Always available; relies on the user-profile ACL. |
  | macOS | **Keychain** | Service name `mcp-microsoft`. |
  | Linux | **libsecret** | Needs a keyring/D-Bus session. Headless Linux without libsecret falls back to a plaintext cache restricted to mode `0600` (a warning is logged). |

### 2.3 http (multi-user) mode

- A **host** to run the process — bare metal, a VM, or a container. For Docker, use a reasonably
  current engine (Docker 20.10+ / Compose v2); the provided image builds on `python:3.12-slim`.
- A **public HTTPS URL** and a **reverse proxy** (Caddy, nginx, Traefik, or a cloud load balancer)
  to terminate TLS. **The server itself speaks plain HTTP** and never terminates TLS.
- An **Entra tenant** where you can create App Registrations and (typically) grant tenant-wide
  **admin consent**. http mode targets **one concrete tenant** (its GUID).
- An **MCP client with OAuth + Streamable HTTP support**. Clients lacking either cannot use http
  mode — those users should use stdio instead.

---

## 3. Understanding the credentials

This section is safe to hand to anyone who needs to gather the values, regardless of their Azure
familiarity.

| Term | What it is | Secret? |
|---|---|---|
| **Application (client) ID** | The public identifier of your Azure App Registration. Every request references it. | **No** — safe to commit / share. |
| **Client secret** | A password for a *confidential* app (http mode only). Proves the server is the registered app so it can perform the OBO exchange. | **Yes** — treat like a password. |
| **Directory (tenant) ID** | The GUID of your Entra directory. Identifies *which* organization. | **No** — not secret, but identifying. |
| **Scope / delegated permission** | A specific capability the app may use *on a signed-in user's behalf* (e.g. `Mail.Send`, `Calendars.ReadWrite`). "Delegated" = acts as the user, never more than the user can do. | No. |
| **Admin consent** | A tenant administrator approving a permission once for the whole organization, so individual users are not each prompted (required for admin-restricted scopes like `Sites.ReadWrite.All`). | N/A. |
| **`MCP_STATS_TOKEN`** | A shared secret that gates the observability routes (http mode). | **Yes** — treat like a password. |

**Where to find the tenant ID:** Azure Portal → **Microsoft Entra ID** → **Overview** →
**Directory (tenant) ID**. Copy the GUID (format `8-4-4-4-12` hex, e.g.
`9a8b7c6d-1234-5678-90ab-cdef01234567`).

> **http mode requires the concrete tenant GUID.** Pseudo-tenants (`organizations`, `common`,
> `consumers`) and verified domains (`contoso.onmicrosoft.com`) are **rejected at startup**. The
> reason: FastMCP's `AzureProvider` pins the accepted token issuer (`iss` claim) to a single literal
> URL built from `MCP_AUTH_TENANT_ID`, and real Entra tokens always carry the *concrete* tenant GUID
> in `iss`. A pseudo-tenant or domain value would never match, so every request would fail
> authentication. The server refuses to start rather than boot into a broken state — see the exact
> error in [§9.1](#91-boot-time-configuration-errors). (stdio profiles, by contrast, happily accept
> `common` / `consumers` / a domain.)

**Which values are secret:**

- **Secret:** `MCP_AUTH_CLIENT_SECRET`, `MCP_STATS_TOKEN`, and the contents of the MSAL token caches
  (`msal_cache_*.bin`).
- **Not secret:** `MS365_CLIENT_ID`, `MCP_AUTH_CLIENT_ID`, tenant IDs, `MCP_BASE_URL`, and
  `profiles.json` (it holds only client/tenant IDs, no secrets).

---

## 4. Azure App Registration

You need **one App Registration per mode**, and they are **not interchangeable**: stdio uses a
**public** client, http uses a **confidential** client with a different redirect platform and a
secret. Do not bolt the http configuration onto the public-client registration your desktop users
depend on — create a separate one.

### 4.1 stdio — public client

1. [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** →
   **+ New registration**.
2. **Name:** anything, e.g. `mcp-microsoft`.
3. **Supported account types:**
   - Personal Outlook.com / Live only → *Personal Microsoft accounts only*
   - Work/school only → *Accounts in this organizational directory only*
   - Both → *Accounts in any organizational directory and personal Microsoft accounts*
4. **Redirect URI:** platform **Mobile and desktop applications**, URI `http://localhost`.
5. **Register.**
6. **Authentication** → **Advanced settings** → set **Allow public client flows** to **Yes**, then
   **Save**. (Required for the interactive loopback / device-code flow — no client secret.)
7. **API permissions** → **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**,
   and add the base set (below). Add optional sets only for services you will enable.
8. For admin-restricted permissions (e.g. `Sites.ReadWrite.All`), click **Grant admin consent**
   (work/school tenants only; personal accounts consent per-user on first login).
9. From **Overview**, copy the **Application (client) ID** and, for a work/school account, the
   **Directory (tenant) ID**.

**Delegated Graph permissions (exact scope names):**

| Set | Enabled by | Scopes |
|---|---|---|
| **Base** (always) | — | `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Contacts.ReadWrite`, `Files.ReadWrite`, `offline_access` |
| **Teams** | `MCP_ENABLE_TEAMS` | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `Channel.Create`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Chat.ReadWrite`, `Chat.Create`, `OnlineMeetings.ReadWrite` |
| **Teams meeting artifacts** | `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` (Teams also on) | `OnlineMeetingTranscript.Read.All`, `OnlineMeetingRecording.Read.All` |
| **Teams AI insights** | `MCP_ENABLE_TEAMS_AI_INSIGHTS` (Teams also on, + Copilot license) | `OnlineMeetingAiInsight.Read.All` |
| **SharePoint** | `MCP_ENABLE_SHAREPOINT` | `Sites.ReadWrite.All` (needs admin consent) |

### 4.2 http — confidential client

> **This is a separate registration from §4.1.** Different platform (Web, not Mobile & desktop),
> and it carries a client secret because the server performs a server-side OBO exchange for each
> user.

1. **App registrations** → **+ New registration**. Name e.g. `mcp-microsoft-remote`.
2. **Supported account types:** work/school only (*this organizational directory only* or *any
   organizational directory*). **Do not** pick a personal-account option — OBO and custom API scopes
   are not reliably supported for consumer accounts.
3. **Redirect URI:** platform **Web**, URI:
   ```
   {MCP_BASE_URL}/auth/callback
   ```
   e.g. `https://mcp.example.com/auth/callback`. `/auth/callback` is `AzureProvider`'s default
   redirect path.
4. **Register.**
5. **Expose an API:**
   - **Application ID URI** → **Add** → accept the default `api://{client_id}` → **Save**.
   - **+ Add a scope** → name `mcp-access` (this matches the server default
     `MCP_AUTH_REQUIRED_SCOPE`; if you use a different name, set that env var to match). Who can
     consent: Admins and users (or Admins only to gate access). State: **Enabled**.
6. **Manifest** → set `"requestedAccessTokenVersion": 2` → **Save**. Required so Entra issues v2.0
   tokens with the claim shapes (`scp`, `oid`, `preferred_username`, `tid`, `iss`) that
   `AzureProvider` and the audit log expect.
7. **Certificates & secrets** → **+ New client secret** → copy the **value** immediately (shown
   once). This becomes `MCP_AUTH_CLIENT_SECRET`. Azure caps expiry at 24 months.
8. **API permissions** → add the same delegated Graph permissions as §4.1 (base always; optional
   sets only for the feature flags you set — in http mode flags must be **explicit**, there is no
   auto-detect). `offline_access` is added automatically by `AzureProvider` — no need to request it.
9. **Grant admin consent** for the tenant (**API permissions** → **Grant admin consent for
   [tenant]**). Because this is a confidential, work/school-only registration, plan on tenant-wide
   consent so new users do not each hit an individual prompt. `Sites.ReadWrite.All` in particular
   will not work without it.

**Client-secret rotation:** the secret has a hard expiry — rotate before then. Create a new secret
alongside the old one (both valid simultaneously), roll `MCP_AUTH_CLIENT_SECRET` to the new value
and **restart** the server, then delete the old secret once you have confirmed the new one works. No
data migration is needed — the server holds the secret only in process memory.

**Admin-consent URL pattern** (share with your IT admin if you cannot consent yourself):

```
https://login.microsoftonline.com/{tenant}/adminconsent?client_id={client_id}
```

---

## 5. Setup: stdio (single user)

### 5.1 MCPB (Claude Desktop) — recommended

Install the `.mcpb` bundle (double-click, or `npx @anthropic-ai/mcpb install mcp-microsoft-0.8.0.mcpb`).
The installer prompts for the App Registration values, which map to `user_config` fields in
`manifest.json`:

| Prompt (`user_config`) | Env var it sets | Notes |
|---|---|---|
| **Azure Client ID** (`client_id`, required) | `MS365_CLIENT_ID` | From §4.1 Overview. |
| **Tenant ID** (`tenant_id`, default `common`) | `MS365_TENANT_ID` | `common` (personal+work), `consumers` (personal), or your tenant ID/domain. |
| **Credentials Directory** (`credentials_dir`) | `MS365_CREDENTIALS_DIR` | Defaults to `~/.microsoft-mcp/`. |
| **Enable Teams Tools** (`enable_teams`, default false) | `MCP_ENABLE_TEAMS` | Work/school only. |
| **Enable Teams Meeting Artifacts** (`enable_teams_meeting_artifacts`) | `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` | Extra Graph permissions. |
| **Enable Teams AI Insights** (`enable_teams_ai_insights`) | `MCP_ENABLE_TEAMS_AI_INSIGHTS` | Extra permissions + Copilot license. |
| **Enable SharePoint Tools** (`enable_sharepoint`, default false) | `MCP_ENABLE_SHAREPOINT` | Work/school only. |
| **Disable Permanent-Delete Tools** (`disable_deletion_tools`) | `MCP_DISABLE_DELETION_TOOLS` | Hides hard-delete tools. |

When `MS365_CLIENT_ID` is set, a `default` profile is created and persisted automatically on first
start. After saving settings, **restart Claude Desktop fully**, then ask Claude to authenticate — a
browser window opens for Microsoft sign-in.

### 5.2 From source

```bash
git clone https://github.com/guilhermeinacio/mcp-microsoft.git
cd mcp-microsoft
uv sync

export MS365_CLIENT_ID=your-client-id
export MS365_TENANT_ID=common          # or consumers, or your tenant ID/domain
# Optional overrides:
# export MCP_ENABLE_SHAREPOINT=true
# export MCP_ENABLE_TEAMS=true
# export MCP_ENABLE_TEAMS_MEETING_ARTIFACTS=true
# export MCP_ENABLE_TEAMS_AI_INSIGHTS=true
# export MCP_DISABLE_DELETION_TOOLS=1
uv run mcp-microsoft
```

Wire it into an MCP client via `claude_desktop_config.json`:

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

### 5.3 Profile management

stdio identity is organized as named **profiles**, each with its own `client_id`, `tenant_id`, and
encrypted token cache. Manage them with the `mcp-microsoft-setup` CLI:

```bash
uv run mcp-microsoft-setup add       # interactive: name, client ID, tenant, authenticate now?
uv run mcp-microsoft-setup auth      # trigger interactive OAuth for a profile
uv run mcp-microsoft-setup list      # list profiles (masked client IDs, auth status, default)
uv run mcp-microsoft-setup remove    # remove a profile and delete its token cache
uv run mcp-microsoft-setup default   # change the default profile
```

Running `mcp-microsoft-setup` with no argument shows an interactive menu with these same commands.

The same operations are also available as MCP tools (`add_ms_profile`, `authenticate_ms_profile`,
`list_ms_profiles`, `remove_ms_profile`, `set_default_ms_profile`) so you can drive them from inside
the client.

**Storage layout** (under `MS365_CREDENTIALS_DIR`, default `~/.microsoft-mcp/`, created mode `0700`
on POSIX):

| File | Contents | Permissions |
|---|---|---|
| `profiles.json` | Per-profile `client_id` + `tenant_id` (**no secrets**). | `0600` on POSIX. |
| `msal_cache_{name}.bin` | OS-encrypted MSAL token cache (refresh tokens) for each profile. | Encrypted at rest (DPAPI/Keychain/libsecret); plaintext fallback restricted to `0600`. |

> **Do not commit** `profiles.json` or `msal_cache_*.bin` to version control. `MS365_CLIENT_ID` is
> not a secret and may be committed. Legacy plaintext `msal_cache_*.json` files from versions before
> 0.7.0 are migrated to encrypted `.bin` automatically on first run and the originals deleted.

### 5.4 Feature flags & the deletion kill-switch

| Env var | Default | Effect |
|---|---|---|
| `MCP_ENABLE_TEAMS` | auto-detect (stdio) | Force Teams tools on/off. Absent → auto-enabled for corporate tenant values (`common`, `organizations`, a GUID/domain), off for `consumers`. |
| `MCP_ENABLE_SHAREPOINT` | auto-detect (stdio) | Force SharePoint tools on/off, same auto-detect logic. |
| `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` | `false` | Register Teams transcript/recording tools + request their scopes. Explicit opt-in only. |
| `MCP_ENABLE_TEAMS_AI_INSIGHTS` | `false` | Register Teams Copilot AI-insight tools + request `OnlineMeetingAiInsight.Read.All`. Explicit opt-in only. |
| `MCP_DISABLE_DELETION_TOOLS` | `false` | **Kill-switch:** when truthy, suppresses registration of all permanent-delete tools. |

The deletion kill-switch (`MCP_DISABLE_DELETION_TOOLS=1`) removes these hard-delete tools:
`delete_email`, `bulk_delete_emails`, `delete_event`, `delete_contact`, `delete_folder`,
`delete_drive_item`, `delete_list_item`, `remove_ms_profile`. Recoverable variants (`trash_email`,
`bulk_trash_emails`, `move_or_copy_item`) remain.

Truthy values for any flag: `1`, `true`, `yes`, `on` (case-insensitive).

### 5.5 Multi-account usage

Configure multiple profiles, then target one on any tool call by passing `profile`:

```
list_emails(folder="Inbox", profile="work")
search_drive(query="Q1 report", profile="personal")
```

Omitting `profile` uses the default profile. After the first interactive login per profile, MSAL
refreshes tokens silently.

---

## 6. Setup: http (multi-user server)

### 6.1 Environment-variable reference (http mode)

Every `MCP_*` variable read in http mode. stdio mode ignores all `MCP_HTTP_*` / `MCP_AUTH_*` values.

| Variable | Purpose | Default | Required in http mode? |
|---|---|---|---|
| `MCP_TRANSPORT` | Selects the mode. Must be `stdio` or `http`. | `stdio` | Set to `http`. |
| `MCP_HTTP_HOST` | Bind host **inside** the process/container. Use `0.0.0.0` behind a proxy/Docker. | `127.0.0.1` | No. |
| `MCP_HTTP_PORT` | Bind port. Must be 1–65535. | `8000` | No. |
| `MCP_BASE_URL` | The **public HTTPS URL** clients reach (your proxy's URL, not the bind address). Must start with `http://` or `https://`. | — | **Yes.** |
| `MCP_AUTH_CLIENT_ID` | Confidential app's Application (client) ID (§4.2). | — | **Yes.** |
| `MCP_AUTH_CLIENT_SECRET` | The client secret value (§4.2). | — | **Yes.** |
| `MCP_AUTH_TENANT_ID` | Your directory's **tenant GUID**. Pseudo-tenants/domains rejected at startup. | — | **Yes.** |
| `MCP_AUTH_REQUIRED_SCOPE` | Custom API scope name from "Expose an API". | `mcp-access` | No (only if you renamed it). |
| `MCP_HTTP_STATELESS` | Stateless Streamable HTTP. Single-worker constraint still applies. | `false` | No. |
| `MCP_RATE_LIMIT_RPS` | Per-user requests/second ceiling (burst = 2×). `0`/negative disables. | `10` | No. |
| `MCP_STATS_TOKEN` | Enables the observability routes (empty = disabled). | *(empty)* | No. |
| `MCP_ENABLE_FILE_UPLOAD` | Drag-drop file-upload app ([§6.8](#68-file-uploads)). Explicit value wins. | on in http | No. |
| `MCP_UPLOAD_MAX_MB` | Max size (MB) per uploaded file. Positive integer. | `10` | No. |
| `MCP_UPLOAD_GLOBAL_BUDGET_MB` | Global cap (MB) on the base64 footprint of all uploads across every user. Positive integer. | `1024` | No. |

Feature flags (`MCP_ENABLE_TEAMS`, `MCP_ENABLE_SHAREPOINT`, `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS`,
`MCP_ENABLE_TEAMS_AI_INSIGHTS`) and `MCP_DISABLE_DELETION_TOOLS` behave as in stdio — **except** the
Teams/SharePoint auto-detect fallback does not apply (no single profile to inspect), so those flags
must be set **explicitly** in http mode or the services stay off.

> **Startup is fail-fast.** In http mode the server validates config before binding and aborts with
> a clear, itemized error if `MCP_BASE_URL`, `MCP_AUTH_CLIENT_ID`, `MCP_AUTH_CLIENT_SECRET`, or
> `MCP_AUTH_TENANT_ID` is missing or malformed (see [§9.1](#91-boot-time-configuration-errors)).

### 6.2 Bare-metal quickstart

```bash
export MCP_TRANSPORT=http
export MCP_BASE_URL=https://mcp.example.com          # your public HTTPS URL (behind a proxy)
export MCP_AUTH_CLIENT_ID=your-confidential-client-id
export MCP_AUTH_CLIENT_SECRET=your-client-secret
export MCP_AUTH_TENANT_ID=your-directory-tenant-guid  # Entra > Overview > Directory (tenant) ID
# Optional: enable services (must be explicit in http mode)
# export MCP_ENABLE_TEAMS=true
# export MCP_ENABLE_SHAREPOINT=true
uv run mcp-microsoft
```

The process binds to `MCP_HTTP_HOST:MCP_HTTP_PORT` (default `127.0.0.1:8000`) and serves the MCP
endpoint at `/mcp` plus an unauthenticated `GET /health`.

### 6.3 Docker & Compose quickstart

The provided `Dockerfile` builds an http-only image (it hardcodes `MCP_TRANSPORT=http`,
`MCP_HTTP_HOST=0.0.0.0`, `MCP_HTTP_PORT=8000`, runs as a non-root user, and has a `/health`
healthcheck). Auth values are intentionally **not** baked in — supply them at run time.

```bash
cp .env.template .env
# Fill in the "Remote server (http) mode" section of .env:
#   MCP_BASE_URL, MCP_AUTH_CLIENT_ID, MCP_AUTH_CLIENT_SECRET, MCP_AUTH_TENANT_ID
docker compose up -d
curl http://localhost:8000/health      # -> {"status":"ok","transport":"http"}
```

Or without Compose:

```bash
docker build -t mcp-microsoft:0.8.0 .
docker run --rm -p 8000:8000 --env-file .env mcp-microsoft:0.8.0
```

> `.env` is gitignored — **never commit it.** Prefer a secrets manager (e.g. Azure Key Vault) that
> injects `MCP_AUTH_CLIENT_SECRET` and `MCP_STATS_TOKEN` as environment variables at deploy time
> over storing them on disk long-term.

### 6.4 Reverse proxy & TLS

The server speaks **plain HTTP** and never terminates TLS. Put a reverse proxy in front of it and
point `MCP_BASE_URL` at the proxy's **public HTTPS URL**.

> **`MCP_BASE_URL` must match exactly** — scheme, host, port, and any path prefix — what MCP clients
> connect to and what you registered as the Azure redirect-URI base. A mismatch breaks the OAuth
> redirect and the JWT audience/issuer checks.

**Caddy** (`Caddyfile`):

```
mcp.example.com {
    reverse_proxy localhost:8000
}
```

**nginx:**

```nginx
server {
    listen 443 ssl;
    server_name mcp.example.com;

    ssl_certificate     /etc/ssl/certs/mcp.example.com.pem;
    ssl_certificate_key /etc/ssl/private/mcp.example.com.key;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host              $host;
        proxy_set_header X-Forwarded-For   $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        # Streamable HTTP is long-lived; relax buffering/timeouts.
        proxy_buffering off;
        proxy_read_timeout 3600s;
    }
}
```

With either, set `MCP_BASE_URL=https://mcp.example.com`. `docker-compose.yml` also documents a
Traefik label-based example in its trailing comment block.

> **Throttle the OAuth endpoints at the proxy.** The built-in rate limiter (§8) covers **only**
> authenticated `/mcp` traffic. The unauthenticated OAuth endpoints — `/authorize`, `/token`,
> `/register`, `/auth/callback` — are **not** rate-limited by the server. Add per-IP throttling for
> those at the proxy if the server is internet-facing.

### 6.5 How clients connect

Point an OAuth-capable MCP client at `https://your-host/mcp`. The flow:

1. The client fetches `/.well-known/oauth-protected-resource` (RFC 9728) and discovers the OAuth
   endpoints — `AzureProvider` acts as an OAuth proxy in front of Entra (which lacks Dynamic Client
   Registration).
2. The user completes Microsoft sign-in in their browser and consents (or the tenant admin already
   consented for everyone).
3. The client then sends an Entra-derived **bearer token** with every request; the server validates
   it and exchanges it **On-Behalf-Of** for a Graph token scoped to that user on each call.

On **first connect** a user experiences a normal Microsoft OAuth sign-in and consent screen; after
that the client handles token refresh. Clients without OAuth or Streamable HTTP support cannot use
http mode.

### 6.6 http vs. stdio — behavior differences

| Aspect | stdio | http |
|---|---|---|
| **Tool count** (all optional services on) | **96** | **88** |
| **Profile-management tools** | Registered | **Not registered** — `add_ms_profile`, `list_ms_profiles`, `remove_ms_profile`, `authenticate_ms_profile`, `set_default_ms_profile` do not exist. |
| **`profile` argument** | Honored | **Inert** — accepted for compatibility, silently ignored; identity always comes from the bearer token. |
| **Local-disk tools** | Available | `download_file`, `download_from_site`, `teams_download_meeting_recording` are **not registered**. `upload_file`/`upload_to_site` reject `local_path` (use `content_base64`); `download_attachment` omits `save_path` and returns content inline; `read_attachment` extracts PDF/text content for the model. |
| **Feature flags** | Env or corporate auto-detect | Env **only** (explicit). |
| **Deletion kill-switch** | Works | Works identically. |
| **Rate limiting / audit logging** | None (single local caller) | On by default. |
| **Error details in responses** | Full | Masked (`mask_error_details=True`). |
| **File-upload app** ([§6.8](#68-file-uploads)) | Off (local users have `local_path`) | **On by default** — adds a drag-drop UI plus 3 model-visible tools (`file_manager`, `list_files`, `read_file`). |

The 87 vs 95 delta is exactly the **5 profile-management tools** plus the **3 local-disk download
tools** that http mode omits. (These counts exclude the optional file-upload app; when it is enabled
it adds 3 model-visible tools on top — see [§6.8](#68-file-uploads).)

### 6.7 No-delete deployment

For teams that must guarantee the assistant can never permanently delete anything, run the server
with the deletion kill-switch:

```bash
export MCP_DISABLE_DELETION_TOOLS=1
```

That is the entire recipe — it works identically in stdio and http mode and is enforced at
**registration time**: the hard-delete tools (`delete_email`, `bulk_delete_emails`, `delete_event`,
`delete_contact`, `delete_folder`, `delete_drive_item`, `delete_list_item`) simply do not exist on
the server, so no client, prompt, or model behavior can invoke them. Recoverable variants
(`trash_email`, `bulk_trash_emails`, `move_or_copy_item`) remain available, so day-to-day cleanup
still works — items land in Deleted Items / the recycle bin instead of vanishing.

> **The switch is server-wide, not per-user.** One http server has one tool set for all connected
> users. If some users need full deletion and others must not have it, run **two instances** —
> one full, one no-delete — and point each user group at the right URL.

**Side-by-side full + no-delete instances** can share a single Azure App Registration: an app
registration accepts multiple redirect URIs, so add both callback URLs (e.g.
`https://mcp.example.com/auth/callback` **and** `https://mcp-nodelete.example.com/auth/callback`)
to the same Web platform, and reuse the same client ID/secret/tenant GUID for both deployments.
Each instance needs its own `MCP_BASE_URL` (matching its public URL exactly) and, if observability
is enabled, its own `MCP_STATS_TOKEN`. A ready-to-uncomment second service for this pattern ships
in `docker-compose.yml`.

For **stdio / Claude Desktop** users the equivalent is the pre-built no-delete bundle
(`mcp-microsoft-nodelete.mcpb`) or the **Disable Permanent-Delete Tools** toggle in the MCPB
installer (see [§5.1](#51-mcpb-claude-desktop--recommended)).

### 6.8 File uploads

**What it is.** In http mode there is no shared disk between the caller and the server, so
`local_path` is rejected and passing a file as base64 forces its whole content through the model's
context window. The **file-upload app** (built on FastMCP's `FileUpload` provider, `fastmcp[apps]`)
fixes this: it exposes a drag-and-drop UI. Files the user drops travel **straight to the server**,
bypassing the model context, and land in a per-user upload area. The upload tools then consume them
**by name** via a new `uploaded_file` parameter on `upload_file` (OneDrive) and `upload_to_site`
(SharePoint) — mutually exclusive with `local_path`/`content_base64`, filename defaults to the
uploaded name.

**Client requirement.** The client must support **MCP Apps** (interactive UI resources) — e.g.
Claude Desktop. Clients without MCP Apps support see the model-visible `list_files`/`read_file`
tools but cannot open the drag-drop UI, so uploads originate elsewhere. The feature is **on by
default in http mode, off in stdio** (local users already have `local_path`); override either way
with `MCP_ENABLE_FILE_UPLOAD`.

**Quotas & limits** (per connected user; enforced in-process, bounded against abuse):

| Limit | Value | Notes |
|---|---|---|
| Max size per file | `MCP_UPLOAD_MAX_MB` (default **10 MB**) | Rejected before storage; must be a positive integer. |
| Max files per user | **20** | Over-quota upload is rejected with a clear message; overwriting a name reuses its slot. |
| Max bytes per user | **100 MB** total | Per-user quota uses the true **decoded** size, not the client-reported size. |
| Global upload budget | `MCP_UPLOAD_GLOBAL_BUDGET_MB` (default **1024 MB**) | Caps the **encoded** (base64) footprint across all users; a store over budget is rejected even if the user is under their own quota. |
| Distinct users tracked | **1000** | Whole least-recently-used upload areas are evicted past the cap. |
| Idle upload-area TTL | **2 h** | Idle areas are lazily pruned on the next upload/list. |

Uploads live **only in the server process's memory**, scoped to the caller's Entra `oid` (stable
across reconnects and stateless mode), falling back to `sub`; a request with **neither** is refused
rather than sharing a bucket. Uploads are lost on restart. File **content is never logged**.
Provider tool calls flow through the same rate-limit / audit / metrics middleware as every other
tool — including the drag-drop backend `store_files`, which is reachable by its hashed name (not
UI-only) and passes through those same middleware, per-file, and quota checks.

**Config:**

| Variable | Purpose | Default |
|---|---|---|
| `MCP_ENABLE_FILE_UPLOAD` | Enable/disable the file-upload app. Explicit value wins. | on in http, off in stdio |
| `MCP_UPLOAD_MAX_MB` | Max size (MB) of any single uploaded file. Positive integer. | `10` |
| `MCP_UPLOAD_GLOBAL_BUDGET_MB` | Global cap (MB) on the base64 footprint of all uploads across every user. Positive integer. | `1024` |

---

## 7. Observability

http mode can expose lightweight, in-process traffic/usage metrics. It is **off by default** and
turns on only when `MCP_STATS_TOKEN` is set to a non-empty secret. When unset, the routes are not
registered at all (a log line notes this at startup). `GET /health` is always open and unaffected.

### 7.1 The three routes

| Route | Returns | Use |
|---|---|---|
| `GET /metrics` | Prometheus text exposition (`text/plain; version=0.0.4`) | Prometheus/Grafana scrape target. |
| `GET /stats` | JSON snapshot: uptime/totals, per-minute traffic (last 60m), per-tool latency p50/p95/avg, per-user activity. | Programmatic dashboards, ad-hoc `curl`. |
| `GET /dashboard` | One self-contained HTML page (inline CSS/JS, no external requests) polling `/stats` every 10s. | Eyeball it in a browser. |

### 7.2 Authentication

All three require the token, presented either way:

- `Authorization: Bearer <MCP_STATS_TOKEN>`, or
- HTTP **Basic** with any username and the token as the **password** (so a browser opening
  `/dashboard` gets a native login prompt).

The comparison is timing-safe; the token is never logged. All three responses are served
`Cache-Control: no-store`.

```bash
# curl with Bearer
curl -H "Authorization: Bearer $MCP_STATS_TOKEN" https://mcp.example.com/metrics
curl -H "Authorization: Bearer $MCP_STATS_TOKEN" https://mcp.example.com/stats

# Browser: open https://mcp.example.com/dashboard and enter any username + the token as password
```

### 7.3 Prometheus scrape config

```yaml
scrape_configs:
  - job_name: mcp-microsoft
    scheme: https
    metrics_path: /metrics
    authorization:
      type: Bearer
      credentials: "<MCP_STATS_TOKEN>"      # or use basic_auth with the token as password
    static_configs:
      - targets: ["mcp.example.com"]
```

### 7.4 Emitted metrics

Process-global and per-tool series (there are **deliberately no per-user label series** —
per-identity cardinality is unbounded and a known Prometheus anti-pattern; per-user detail lives in
`/stats` and `/dashboard`):

- `mcp_uptime_seconds` (gauge)
- `mcp_calls_total`, `mcp_errors_total` (counters)
- `mcp_users_tracked` (gauge), `mcp_users_evicted_total` (counter)
- `mcp_unknown_tool_calls_total`, `mcp_tools_evicted_total` (counters)
- per-tool: `mcp_tool_calls_total{tool}`, `mcp_tool_errors_total{tool}`,
  `mcp_tool_duration_ms{tool,stat="p50|p95|avg"}`

> **Restart resets everything.** Metrics are in-memory with no persistence, and cover only the
> single worker the process runs. A restart zeroes them.

> **Keep it private.** The token is the only gate. Treat these routes as sensitive operational data
> and additionally restrict them (IP allowlist / separate auth) at your reverse proxy if the server
> is internet-facing. Leaving `MCP_STATS_TOKEN` unset disables them entirely.

---

## 8. Security checklist

| Item | What to do | Why |
|---|---|---|
| **TLS only** | Terminate TLS at a reverse proxy; set `MCP_BASE_URL` to the public HTTPS URL. | OAuth flows and every Graph call carry bearer tokens — never send them over plain HTTP. |
| **Tenant GUID** | `MCP_AUTH_TENANT_ID` must be the concrete tenant GUID. | Pseudo-tenants/domains are rejected at startup; issuer pinning means only the GUID validates real tokens. |
| **Secret storage** | Keep `MCP_AUTH_CLIENT_SECRET` / `MCP_STATS_TOKEN` in a secrets manager or env injection, not committed files. | They are passwords. `.env` is gitignored. |
| **Secret rotation** | Rotate the client secret before its (≤24-month) expiry: add a new secret, roll the env var, restart, delete the old one. | Both secrets are valid simultaneously, so no downtime. |
| **Rate limiting** | Leave `MCP_RATE_LIMIT_RPS` on (default `10`/s per user, burst `2×`). Per-user key is `tid:oid`; bucket store is LRU-capped at 10,000 keys and idle-pruned after 900s. | Prevents one user starving others; bounded memory. Over-limit raises JSON-RPC `-32000`. |
| **Proxy throttling** | Throttle `/authorize`, `/token`, `/register`, `/auth/callback` at the proxy. | The built-in limiter covers only authenticated `/mcp` traffic. |
| **Audit logs** | One line per tool call: tool name, caller `oid` + `preferred_username`, duration, outcome. | Never logs arguments, results, or the token itself. |
| **Deletion kill-switch** | Set `MCP_DISABLE_DELETION_TOOLS=1` where hard deletes should be impossible. | Removes all permanent-delete tools; recoverable variants remain. |
| **Disk-tool gating** | Automatic in http mode — download-to-disk tools not registered; `save_path`/`local_path` rejected. | The server's disk is not the caller's disk. |
| **Single worker** | Run exactly one worker/replica in http mode. | The OAuth-proxy client store and per-user OBO cache are in-process memory; multiple workers split state and break sessions. Horizontal scaling needs FastMCP's external `client_storage` (not wired up here). |
| **Non-root container** | The image already runs as a non-root user. | No reason for the process to write outside its venv/tmp. |

---

## 9. Troubleshooting

### 9.1 Boot-time configuration errors (http mode)

http mode validates config before binding. Missing/invalid values abort with:

```
Cannot start http transport — fix the following configuration problems:
  - <problem 1>
  - <problem 2>
```

Actual problem messages:

- `MCP_BASE_URL is required in http mode`
- `MCP_BASE_URL must start with http:// or https:// (got '...')`
- `MCP_AUTH_CLIENT_ID is required in http mode`
- `MCP_AUTH_CLIENT_SECRET is required in http mode`
- `MCP_AUTH_TENANT_ID is required in http mode`
- `MCP_HTTP_PORT must be between 1 and 65535 (got ...)`
- The tenant-GUID rejection (verbatim):
  > `MCP_AUTH_TENANT_ID must be your directory's tenant GUID (8-4-4-4-12 hexadecimal), not '...'. Copy the Directory (tenant) ID from Azure Portal -> Microsoft Entra ID -> Overview. Pseudo-tenants ('organizations', 'common', 'consumers') and verified domains ('contoso.onmicrosoft.com') are rejected because fastmcp's AzureProvider validates the token 'iss' claim against a single literal issuer URL built from this value, and real Entra tokens always carry the concrete tenant GUID -- so a pseudo-tenant or domain never matches and every request fails authentication.`

**Fix:** supply the missing value(s); for the last one, put the concrete tenant GUID in
`MCP_AUTH_TENANT_ID`.

### 9.2 `MCP_TRANSPORT` typo

If `MCP_TRANSPORT` is neither `stdio` nor `http`, startup aborts:

```
Cannot start mcp-microsoft — MCP_TRANSPORT must be 'stdio' or 'http' (got 'htttp')
```

**Fix:** correct the value (or unset it to default to `stdio`).

### 9.3 `AADSTS65001` — admin consent required

Graph returns `AADSTS65001: The user or administrator has not consented`. In stdio mode the server
surfaces an admin-consent URL:

```
https://login.microsoftonline.com/common/adminconsent?client_id={client_id}
```

**Fix:** a tenant admin must grant consent (share your Application (client) ID). For admin-restricted
scopes like `Sites.ReadWrite.All` this is required. In http mode, grant tenant-wide admin consent on
the confidential registration (§4.2 step 9).

### 9.4 `401` from an OAuth client (http mode)

The client's token is missing or lacks the required scope. Checks:

- The App Registration's **Expose an API** scope name matches `MCP_AUTH_REQUIRED_SCOPE` (default
  `mcp-access`), and the client requests it.
- `"requestedAccessTokenVersion": 2` is set in the app manifest (§4.2 step 6) — otherwise the token
  claim shapes (`scp`/`oid`/`iss`) will not validate.
- `MCP_BASE_URL` exactly matches the redirect-URI base and the URL the client connects to.
- `MCP_AUTH_TENANT_ID` is the concrete GUID (the issuer pinning rejects mismatches).

### 9.5 Rate-limit error (http mode)

Over-limit calls raise `RateLimitError` — an `McpError` with JSON-RPC code **`-32000`** and message
`Rate limit exceeded for client: <tid:oid>`.

**Fix:** back off, or raise/disable the limit via `MCP_RATE_LIMIT_RPS` (`0` or negative disables it).

### 9.6 stdio interactive login fails / headless

If the interactive browser flow is unavailable (MCPB, SSH, containers), the server falls back to the
**device-code flow** and logs the code + verification URL as a warning. Complete sign-in on any
device using that code.

### 9.7 Other Azure errors (stdio)

- `AADSTS50011: redirect URI does not match` → add `http://localhost` under **Mobile and desktop
  applications** (not Web, not `https://`).
- `AADSTS700016: Application not found in the directory` → tenant ID does not match the account. Use
  `consumers` for a personal account, or the Directory (tenant) ID for a work account.
- **SharePoint/Teams tools not appearing** → the corresponding flag is off, or the account is
  personal (SharePoint and Teams are work/school only).

### 9.8 Where logs go

The server logs to standard Python logging (stderr for the stdio subprocess; container stdout/stderr
for Docker — view with `docker compose logs -f`). Audit and rate-limit events appear there in http
mode. The observability `/stats` and `/dashboard` routes surface live in-process metrics.

---

## 10. Limits & FAQ

- **Personal accounts are stdio-only.** OBO and custom API scopes are not reliably supported for
  consumer Microsoft accounts, so http mode is work/school-only. Use stdio for Outlook.com / Live.
- **Multi-tenant is not supported (http).** The server targets one concrete tenant GUID; the issuer
  is pinned to a literal URL built from it. Multi-tenant deployments are future work (they need
  issuer-validation skipping plus per-tenant OBO authority).
- **Metrics reset on restart.** In-memory only, single worker, no persistence.
- **Single worker only (http).** The OAuth-proxy client store and per-user OBO cache are in-process.
  Do not run multiple workers/replicas without wiring FastMCP's external `client_storage` (not
  implemented here). `MCP_HTTP_STATELESS` does not change this.
- **MCPB updates.** The stdio MCPB bundle is updated by installing a newer `.mcpb`; settings
  (`user_config`) are preserved by Claude Desktop. Restart Claude Desktop fully after updating.
- **Teams throttling.** Teams Graph endpoints throttle around ~4 req/s and may return 429; some
  meeting-listing endpoints require an OData `$filter` and can 400 on tenants that do not support it.

---

## 11. Appendix: env vars & tool inventory

### 11.1 Complete environment-variable table

| Variable | Mode | Default | Description |
|---|---|---|---|
| `MS365_CLIENT_ID` | stdio | — | Client ID for the bootstrap `default` profile. Not a secret. |
| `MS365_TENANT_ID` | stdio | `common` | Tenant for the bootstrap profile (`common`/`consumers`/GUID/domain). |
| `MS365_CREDENTIALS_DIR` | stdio | `~/.microsoft-mcp/` | Directory for `profiles.json` and token caches. |
| `MCP_ENABLE_TEAMS` | both | auto-detect (stdio) / off (http) | Force Teams tools on/off. |
| `MCP_ENABLE_SHAREPOINT` | both | auto-detect (stdio) / off (http) | Force SharePoint tools on/off. |
| `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` | both | `false` | Teams transcript/recording tools (explicit opt-in). |
| `MCP_ENABLE_TEAMS_AI_INSIGHTS` | both | `false` | Teams Copilot AI-insight tools (explicit opt-in; needs Copilot license). |
| `MCP_DISABLE_DELETION_TOOLS` | both | `false` | Kill-switch: suppress all permanent-delete tools. |
| `MCP_TRANSPORT` | both | `stdio` | `stdio` or `http`. |
| `MCP_HTTP_HOST` | http | `127.0.0.1` | Bind host inside the process/container. |
| `MCP_HTTP_PORT` | http | `8000` | Bind port (1–65535). |
| `MCP_HTTP_STATELESS` | http | `false` | Stateless Streamable HTTP. |
| `MCP_BASE_URL` | http | — | Public HTTPS URL (required). |
| `MCP_AUTH_CLIENT_ID` | http | — | Confidential client ID (required). |
| `MCP_AUTH_CLIENT_SECRET` | http | — | Client secret (required, secret). |
| `MCP_AUTH_TENANT_ID` | http | — | Tenant GUID (required; pseudo-tenants/domains rejected). |
| `MCP_AUTH_REQUIRED_SCOPE` | http | `mcp-access` | Custom API scope name. |
| `MCP_RATE_LIMIT_RPS` | http | `10` | Per-user requests/second (burst 2×); `0`/negative disables. |
| `MCP_STATS_TOKEN` | http | *(empty)* | Enables observability routes (secret). |

### 11.2 Tool inventory by mode

Counts assume all optional services enabled.

| Group | stdio | http | Notes |
|---|---|---|---|
| Mail | 26 | 26 | Includes server-side PDF/text attachment extraction. |
| Calendar | 10 | 10 | |
| OneDrive | 8 | 7 | `download_file` omitted in http. |
| SharePoint | 13 | 12 | `download_from_site` omitted in http. |
| Contacts | 8 | 8 | `get_contact_photo` rejects `save_path` in http. |
| Teams | 25 | 24 | `teams_download_meeting_recording` omitted in http. |
| Profile management | 5 | 0 | Not registered in http. |
| Service utilities | 1 | 1 | `list_enabled_services`. |
| **Total** | **96** | **88** | |

**Tools omitted entirely in http mode:** `add_ms_profile`, `list_ms_profiles`, `remove_ms_profile`,
`authenticate_ms_profile`, `set_default_ms_profile` (profile management); `download_file`,
`download_from_site`, `teams_download_meeting_recording` (local-disk downloads).

**Tools that reject disk parameters in http mode** (but remain registered): `upload_file` /
`upload_to_site` reject `local_path` (use `content_base64`); `download_attachment`
omits `save_path` in http mode (content returned inline); `get_contact_photo` rejects
`save_path`. Use `read_attachment` when the model needs PDF or text contents directly.

---

*Guide for mcp-microsoft v0.8.0. For the Azure walkthrough with screenshots see
[`docs/azure-setup.md`](azure-setup.md); for the condensed quickstarts see the [README](../README.md).*
