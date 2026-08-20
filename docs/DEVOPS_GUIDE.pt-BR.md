# mcp-microsoft — Guia de DevOps & Operações

Uma referência completa de configuração e operação para o `mcp-microsoft` **v0.8.0**, cobrindo os
dois modos de implantação: o servidor de usuário único em **stdio** (Claude Desktop / MCPB / a
partir do código-fonte) e o servidor **HTTP remoto** multiusuário (`MCP_TRANSPORT=http`). Ele foi
escrito para ser autocontido — você não deveria precisar ler o código-fonte para colocar nenhum dos
dois modos em funcionamento.

> **Procurando a versão resumida?** O [`README`](../README.md) do projeto tem quickstarts
> condensados. Este guia é a referência de operação de longo formato: cada variável de ambiente,
> cada passo do Azure, o modelo de segurança e a solução dos erros que você realmente vai encontrar.

**English:** an English-language edition of this guide lives at
[`docs/DEVOPS_GUIDE.md`](DEVOPS_GUIDE.md) ("English version").

---

## Sumário

1. [Visão geral](#1-visão-geral)
2. [Requisitos](#2-requisitos)
3. [Entendendo as credenciais](#3-entendendo-as-credenciais)
4. [Registro de aplicativo no Azure](#4-registro-de-aplicativo-no-azure)
5. [Configuração: stdio (usuário único)](#5-configuração-stdio-usuário-único)
6. [Configuração: http (servidor multiusuário)](#6-configuração-http-servidor-multiusuário)
7. [Observabilidade](#7-observabilidade)
8. [Checklist de segurança](#8-checklist-de-segurança)
9. [Solução de problemas](#9-solução-de-problemas)
10. [Limites e FAQ](#10-limites-e-faq)
11. [Apêndice: variáveis de ambiente e inventário de ferramentas](#11-apêndice-variáveis-de-ambiente-e-inventário-de-ferramentas)

---

## 1. Visão geral

`mcp-microsoft` é um servidor [Model Context Protocol](https://modelcontextprotocol.io) que expõe o
Microsoft 365 — Mail, Calendar, OneDrive, SharePoint, Contacts e Teams — para qualquer cliente MCP
(Claude Desktop, Claude Code, VS Code etc.) por meio da **Microsoft Graph API**. Funciona tanto com
contas pessoais da Microsoft (Outlook.com / Live) quanto com contas corporativas (Microsoft Entra ID
/ Azure AD).

A partir da versão 0.8.0, o servidor é executado em um de **dois modos mutuamente exclusivos**,
escolhido na inicialização pela variável de ambiente `MCP_TRANSPORT`:

| | **stdio** (padrão) | **http** (remoto, multiusuário) |
|---|---|---|
| `MCP_TRANSPORT` | `stdio` | `http` |
| Transporte | stdio (subprocesso local) | Streamable HTTP (spec MCP 2025-11-25), servido em `/mcp` |
| Quem executa | Iniciado pelo seu cliente MCP na sua máquina | Um servidor compartilhado que você hospeda atrás de um proxy reverso |
| Identidade | **Perfis** locais nomeados; OAuth interativo/device-code; cache de tokens criptografado em disco | Cada usuário faz login com **sua própria** conta Entra; token de portador (bearer) por requisição |
| Como as chamadas ao Graph são autorizadas | Token MSAL de client público para o perfil ativo | Troca de token **On-Behalf-Of (OBO)** por usuário a partir do bearer token do chamador |
| Tipo de app no Azure | Client **público** (plataforma Mobile & desktop) | Client **confidencial** (plataforma Web + segredo) |
| Contas | Pessoais **e** corporativas/educacionais | **Somente** corporativas/educacionais, um único tenant concreto |
| Quantidade de ferramentas (com todos os serviços opcionais ativos) | **96** | **88** (ferramentas de gerenciamento de perfil e de disco local omitidas) |

**Arquitetura em cinco linhas:**

1. Construído sobre **FastMCP** (`fastmcp[azure]>=3.4.4`), MSAL e `httpx` assíncrono.
2. No modo **stdio**, um `ProfileManager` mantém perfis nomeados; cada um obtém um token do Graph via
   MSAL (navegador interativo, com fallback para device code) e o armazena em cache, criptografado
   pelo sistema operacional, em disco local.
3. No modo **http**, o **`AzureProvider`** do FastMCP implementa o padrão OAuth-proxy exigido pelo
   Entra (o Entra não possui Dynamic Client Registration), anunciando
   `/.well-known/oauth-protected-resource` (RFC 9728) para que os clientes MCP possam descobrir e
   conduzir o fluxo OAuth.
4. Toda chamada ao Graph no modo http é executada sob a identidade do chamador via uma troca
   **On-Behalf-Of** por usuário (`azure.identity` `OnBehalfOfCredential`) — sem credenciais de
   serviço compartilhadas, sem perfis no servidor.
5. Todos os corpos das ferramentas passam por um único ponto (`GraphClient`), de modo que os dois
   caminhos de identidade compartilham as mesmas implementações de ferramentas do Graph.

**Qual modo escolher?**

- **stdio** — uma única pessoa no próprio laptop, conta pessoal ou corporativa, sem servidor para
  operar. É o padrão e a experiência do MCPB / Claude Desktop.
- **http** — você quer oferecer ferramentas do Microsoft 365 para muitos usuários a partir de um
  serviço hospedado único, cada um com sua própria identidade delegada e sem instalação local.
  Requer um tenant corporativo/educacional, um endpoint HTTPS público e um Registro de Aplicativo
  confidencial no Azure.

---

## 2. Requisitos

### 2.1 Comum aos dois modos

- **Python `>=3.11`** (a imagem Docker usa 3.12).
- **[uv](https://docs.astral.sh/uv/)** para gerenciamento de dependências e execução:
  ```bash
  # macOS / Linux
  curl -LsSf https://astral.sh/uv/install.sh | sh
  # Windows (PowerShell)
  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
  ```
- **Rede de saída (outbound)** para `https://graph.microsoft.com` (Graph API) e
  `https://login.microsoftonline.com` (OAuth / endpoints de token do Entra).
- Uma **conta Microsoft** — pessoal (Outlook.com / Live) para stdio, ou um **tenant Entra** onde você
  possa criar Registros de Aplicativo para qualquer um dos modos.

### 2.2 Modo stdio / MCPB

- **Claude Desktop** (ou qualquer cliente MCP que inicie servidores stdio). Dois caminhos de
  instalação:
  - **Pacote MCPB** — dê duplo clique no arquivo `.mcpb` para instalar pelo instalador de extensões
    do Claude Desktop, **ou** instale via shell:
    ```bash
    npx @anthropic-ai/mcpb install mcp-microsoft-0.8.0.mcpb
    ```
    **Node.js só é necessário para o caminho via `npx`** — o caminho de duplo clique não precisa de
    nada extra.
  - **A partir do código-fonte** — clone o repositório e execute via `uv` (veja [§5.2](#52-a-partir-do-código-fonte)).
    Funciona com qualquer cliente MCP que aceite um comando stdio.
- **Backends de criptografia do cache de tokens do sistema operacional** (usados automaticamente
  pelo `msal-extensions` para criptografar o cache de refresh tokens em repouso):

  | SO | Backend | Observações |
  |---|---|---|
  | Windows | **DPAPI** | Sempre disponível; depende da ACL do perfil do usuário. |
  | macOS | **Keychain** | Nome de serviço `mcp-microsoft`. |
  | Linux | **libsecret** | Precisa de uma sessão de keyring/D-Bus. Linux headless sem libsecret cai
  em um cache em texto puro (plaintext) restrito ao modo `0600` (um aviso é registrado no log). |

### 2.3 Modo http (multiusuário)

- Um **host** para executar o processo — bare metal, uma VM ou um container. Para Docker, use um
  engine razoavelmente atual (Docker 20.10+ / Compose v2); a imagem fornecida é construída sobre
  `python:3.12-slim`.
- Uma **URL HTTPS pública** e um **proxy reverso** (Caddy, nginx, Traefik ou um load balancer de
  nuvem) para terminar o TLS. **O próprio servidor fala HTTP puro** e nunca termina TLS.
- Um **tenant Entra** onde você possa criar Registros de Aplicativo e (normalmente) conceder
  **consentimento do administrador** em nível de tenant. O modo http tem como alvo **um único
  tenant concreto** (seu GUID).
- Um **cliente MCP com suporte a OAuth + Streamable HTTP**. Clientes sem um dos dois não podem usar
  o modo http — esses usuários devem usar stdio.

---

## 3. Entendendo as credenciais

Esta seção é segura para ser compartilhada com qualquer pessoa que precise coletar os valores,
independentemente de sua familiaridade com o Azure.

| Termo | O que é | Segredo? |
|---|---|---|
| **Application (client) ID** | O identificador público do seu Registro de Aplicativo no Azure. Toda requisição faz referência a ele. | **Não** — seguro para versionar / compartilhar. |
| **Client secret** (segredo do cliente) | Uma senha para um app *confidencial* (somente modo http). Comprova que o servidor é o app registrado, permitindo a troca OBO. | **Sim** — trate como uma senha. |
| **Directory (tenant) ID** | O GUID do seu diretório Entra. Identifica *qual* organização. | **Não** — não é secreto, mas é identificador. |
| **Scope / permissão delegada (escopo)** | Uma capacidade específica que o app pode usar *em nome de um usuário autenticado* (ex.: `Mail.Send`, `Calendars.ReadWrite`). "Delegada" = age como o usuário, nunca mais do que o próprio usuário pode fazer. | Não. |
| **Admin consent** (consentimento do administrador) | Um administrador do tenant aprovando uma permissão de uma vez para toda a organização, de modo que usuários individuais não sejam questionados um a um (obrigatório para escopos restritos a administrador, como `Sites.ReadWrite.All`). | N/A. |
| **`MCP_STATS_TOKEN`** | Um segredo compartilhado que protege as rotas de observabilidade (modo http). | **Sim** — trate como uma senha. |

**Onde encontrar o tenant ID:** Azure Portal → **Microsoft Entra ID** → **Overview (Visão geral)** →
**Directory (tenant) ID** (ID do diretório/locatário). Copie o GUID (formato `8-4-4-4-12`
hexadecimal, ex.: `9a8b7c6d-1234-5678-90ab-cdef01234567`).

> **O modo http exige o GUID concreto do tenant (locatário).** Pseudo-tenants (`organizations`,
> `common`, `consumers`) e domínios verificados (`contoso.onmicrosoft.com`) são **rejeitados na
> inicialização**. O motivo: o `AzureProvider` do FastMCP fixa o issuer (`iss`) aceito do token a
> uma única URL literal construída a partir de `MCP_AUTH_TENANT_ID`, e tokens reais do Entra sempre
> carregam o GUID *concreto* do tenant na claim `iss`. Um pseudo-tenant ou domínio nunca
> corresponderia, então toda requisição falharia na autenticação. O servidor se recusa a iniciar em
> vez de subir em um estado quebrado — veja o erro exato em
> [§9.1](#91-erros-de-configuração-na-inicialização-modo-http). (Já os perfis stdio aceitam
> tranquilamente `common` / `consumers` / um domínio.)

**Quais valores são secretos:**

- **Secreto:** `MCP_AUTH_CLIENT_SECRET`, `MCP_STATS_TOKEN` e o conteúdo dos caches de token do MSAL
  (`msal_cache_*.bin`).
- **Não secreto:** `MS365_CLIENT_ID`, `MCP_AUTH_CLIENT_ID`, os tenant IDs, `MCP_BASE_URL` e
  `profiles.json` (contém apenas IDs de client/tenant, nenhum segredo).

---

## 4. Registro de aplicativo no Azure

Você precisa de **um Registro de Aplicativo por modo**, e eles **não são intercambiáveis**: o stdio
usa um client **público**, o http usa um client **confidencial** com uma plataforma de redirect
diferente e um segredo. Não anexe a configuração de http ao registro de client público do qual seus
usuários de desktop dependem — crie um separado.

### 4.1 stdio — cliente público

1. [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations
   (Registros de aplicativo)** → **+ New registration (Novo registro)**.
2. **Name (Nome):** qualquer valor, ex.: `mcp-microsoft`.
3. **Supported account types (Tipos de conta com suporte):**
   - Somente Outlook.com / Live pessoal → *Personal Microsoft accounts only (Somente contas
     pessoais da Microsoft)*
   - Somente corporativa/educacional → *Accounts in this organizational directory only (Contas
     somente neste diretório organizacional)*
   - Ambas → *Accounts in any organizational directory and personal Microsoft accounts (Contas em
     qualquer diretório organizacional e contas pessoais da Microsoft)*
4. **Redirect URI (URI de redirecionamento):** plataforma **Mobile and desktop applications**, URI
   `http://localhost`.
5. **Register (Registrar).**
6. **Authentication (Autenticação)** → **Advanced settings (Configurações avançadas)** → defina
   **Allow public client flows (Permitir fluxos de client público)** como **Yes (Sim)**, depois
   **Save (Salvar)**. (Necessário para o fluxo interativo de loopback / device-code — sem client
   secret.)
7. **API permissions (Permissões de API)** → **+ Add a permission (Adicionar uma permissão)** →
   **Microsoft Graph** → **Delegated permissions (Permissões delegadas)**, e adicione o conjunto
   base (abaixo). Adicione conjuntos opcionais apenas para os serviços que você for habilitar.
8. Para permissões restritas a administrador (ex.: `Sites.ReadWrite.All`), clique em **Grant admin
   consent (Conceder consentimento do administrador)** (somente tenants corporativos/educacionais;
   contas pessoais consentem individualmente no primeiro login).
9. Em **Overview (Visão geral)**, copie o **Application (client) ID** e, para conta
   corporativa/educacional, o **Directory (tenant) ID**.

**Permissões delegadas do Graph (nomes exatos dos escopos):**

| Conjunto | Habilitado por | Escopos |
|---|---|---|
| **Base** (sempre) | — | `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Contacts.ReadWrite`, `Files.ReadWrite`, `offline_access` |
| **Teams** | `MCP_ENABLE_TEAMS` | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `Channel.Create`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Chat.ReadWrite`, `Chat.Create`, `OnlineMeetings.ReadWrite` |
| **Artefatos de reunião do Teams** | `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` (Teams também ativo) | `OnlineMeetingTranscript.Read.All`, `OnlineMeetingRecording.Read.All` |
| **Insights de IA do Teams** | `MCP_ENABLE_TEAMS_AI_INSIGHTS` (Teams também ativo, + licença Copilot) | `OnlineMeetingAiInsight.Read.All` |
| **SharePoint** | `MCP_ENABLE_SHAREPOINT` | `Sites.ReadWrite.All` (precisa de consentimento do administrador) |

### 4.2 http — cliente confidencial

> **Este é um registro separado do §4.1.** Plataforma diferente (Web, não Mobile & desktop), e ele
> carrega um client secret porque o servidor executa uma troca OBO no lado servidor para cada
> usuário.

1. **App registrations (Registros de aplicativo)** → **+ New registration (Novo registro)**. Nome
   ex.: `mcp-microsoft-remote`.
2. **Supported account types (Tipos de conta com suporte):** somente corporativa/educacional (*this
   organizational directory only* ou *any organizational directory*). **Não** escolha uma opção de
   conta pessoal — OBO e escopos de API personalizados não têm suporte confiável para contas de
   consumidor.
3. **Redirect URI (URI de redirecionamento):** plataforma **Web**, URI:
   ```
   {MCP_BASE_URL}/auth/callback
   ```
   ex.: `https://mcp.example.com/auth/callback`. `/auth/callback` é o caminho de redirect padrão do
   `AzureProvider`.
4. **Register (Registrar).**
5. **Expose an API (Expor uma API):**
   - **Application ID URI** → **Add (Adicionar)** → aceite o padrão `api://{client_id}` → **Save
     (Salvar)**.
   - **+ Add a scope (Adicionar um escopo)** → nome `mcp-access` (corresponde ao padrão do servidor
     `MCP_AUTH_REQUIRED_SCOPE`; se usar um nome diferente, ajuste essa variável de ambiente para
     corresponder). Who can consent (Quem pode consentir): Admins and users, ou Admins only para
     restringir o acesso. State (Estado): **Enabled (Habilitado)**.
6. **Manifest (Manifesto)** → defina `"requestedAccessTokenVersion": 2` → **Save (Salvar)**.
   Necessário para que o Entra emita tokens v2.0 com o formato de claims (`scp`, `oid`,
   `preferred_username`, `tid`, `iss`) que o `AzureProvider` e o log de auditoria esperam.
7. **Certificates & secrets (Certificados e segredos)** → **+ New client secret (Novo client
   secret)** → copie o **value (valor)** imediatamente (exibido uma única vez). Isso se torna
   `MCP_AUTH_CLIENT_SECRET`. O Azure limita a validade a 24 meses.
8. **API permissions (Permissões de API)** → adicione as mesmas permissões delegadas do Graph do
   §4.1 (base sempre; conjuntos opcionais apenas para as feature flags que você definir — no modo
   http as flags precisam ser **explícitas**, não há auto-detecção). `offline_access` é adicionado
   automaticamente pelo `AzureProvider` — não é preciso solicitá-lo.
9. **Grant admin consent for the tenant (Conceder consentimento do administrador para o tenant)**
   (**API permissions** → **Grant admin consent for [tenant]**). Como este é um registro
   confidencial, restrito a corporativo/educacional, planeje um consentimento em nível de tenant
   para que novos usuários não sejam questionados individualmente. Em particular,
   `Sites.ReadWrite.All` não funcionará sem isso.

**Rotação do client secret:** o segredo tem uma expiração fixa — rotacione antes disso. Crie um novo
segredo em paralelo com o antigo (ambos válidos simultaneamente), atualize `MCP_AUTH_CLIENT_SECRET`
para o novo valor e **reinicie** o servidor, depois exclua o segredo antigo quando confirmar que o
novo funciona. Nenhuma migração de dados é necessária — o servidor mantém o segredo apenas em
memória do processo.

**Padrão da URL de consentimento do administrador** (compartilhe com seu administrador de TI, caso
você mesmo não possa consentir):

```
https://login.microsoftonline.com/{tenant}/adminconsent?client_id={client_id}
```

---

## 5. Configuração: stdio (usuário único)

### 5.1 MCPB (Claude Desktop) — recomendado

Instale o pacote `.mcpb` (duplo clique, ou
`npx @anthropic-ai/mcpb install mcp-microsoft-0.8.0.mcpb`). O instalador solicita os valores do
Registro de Aplicativo, que são mapeados para os campos `user_config` em `manifest.json`:

| Prompt (`user_config`) | Variável de ambiente que define | Observações |
|---|---|---|
| **Azure Client ID** (`client_id`, obrigatório) | `MS365_CLIENT_ID` | Do Overview (Visão geral) do §4.1. |
| **Tenant ID** (`tenant_id`, padrão `common`) | `MS365_TENANT_ID` | `common` (pessoal + corporativa), `consumers` (pessoal) ou o ID/domínio do seu tenant. |
| **Credentials Directory** (`credentials_dir`) | `MS365_CREDENTIALS_DIR` | Padrão `~/.microsoft-mcp/`. |
| **Enable Teams Tools** (`enable_teams`, padrão false) | `MCP_ENABLE_TEAMS` | Somente corporativa/educacional. |
| **Enable Teams Meeting Artifacts** (`enable_teams_meeting_artifacts`) | `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` | Permissões extras do Graph. |
| **Enable Teams AI Insights** (`enable_teams_ai_insights`) | `MCP_ENABLE_TEAMS_AI_INSIGHTS` | Permissões extras + licença Copilot. |
| **Enable SharePoint Tools** (`enable_sharepoint`, padrão false) | `MCP_ENABLE_SHAREPOINT` | Somente corporativa/educacional. |
| **Disable Permanent-Delete Tools** (`disable_deletion_tools`) | `MCP_DISABLE_DELETION_TOOLS` | Oculta as ferramentas de exclusão definitiva. |

Quando `MS365_CLIENT_ID` está definido, um perfil `default` é criado e persistido automaticamente no
primeiro início. Após salvar as configurações, **reinicie completamente o Claude Desktop** e depois
peça ao Claude para autenticar — uma janela do navegador abrirá para o login da Microsoft.

### 5.2 A partir do código-fonte

```bash
git clone https://github.com/guilhermeinacio/mcp-microsoft.git
cd mcp-microsoft
uv sync

export MS365_CLIENT_ID=your-client-id
export MS365_TENANT_ID=common          # ou consumers, ou o ID/domínio do seu tenant
# Overrides opcionais:
# export MCP_ENABLE_SHAREPOINT=true
# export MCP_ENABLE_TEAMS=true
# export MCP_ENABLE_TEAMS_MEETING_ARTIFACTS=true
# export MCP_ENABLE_TEAMS_AI_INSIGHTS=true
# export MCP_DISABLE_DELETION_TOOLS=1
uv run mcp-microsoft
```

Conecte a um cliente MCP via `claude_desktop_config.json`:

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

### 5.3 Gerenciamento de perfis

A identidade no modo stdio é organizada em **perfis** nomeados, cada um com seu próprio `client_id`,
`tenant_id` e cache de token criptografado. Gerencie-os com a CLI `mcp-microsoft-setup`:

```bash
uv run mcp-microsoft-setup add       # interativo: nome, client ID, tenant, autenticar agora?
uv run mcp-microsoft-setup auth      # dispara o OAuth interativo para um perfil
uv run mcp-microsoft-setup list      # lista os perfis (client IDs mascarados, status de auth, padrão)
uv run mcp-microsoft-setup remove    # remove um perfil e exclui seu cache de token
uv run mcp-microsoft-setup default   # altera o perfil padrão
```

Executar `mcp-microsoft-setup` sem argumento exibe um menu interativo com esses mesmos comandos.

As mesmas operações também estão disponíveis como ferramentas MCP (`add_ms_profile`,
`authenticate_ms_profile`, `list_ms_profiles`, `remove_ms_profile`, `set_default_ms_profile`), para
que você possa conduzi-las de dentro do cliente.

**Layout de armazenamento** (sob `MS365_CREDENTIALS_DIR`, padrão `~/.microsoft-mcp/`, criado no modo
`0700` em POSIX):

| Arquivo | Conteúdo | Permissões |
|---|---|---|
| `profiles.json` | `client_id` + `tenant_id` por perfil (**sem segredos**). | `0600` em POSIX. |
| `msal_cache_{name}.bin` | Cache de token do MSAL criptografado pelo SO (refresh tokens) para cada perfil. | Criptografado em repouso (DPAPI/Keychain/libsecret); fallback em texto puro restrito a `0600`. |

> **Não faça commit** de `profiles.json` ou `msal_cache_*.bin` no controle de versão.
> `MS365_CLIENT_ID` não é um segredo e pode ser versionado. Arquivos legados em texto puro
> `msal_cache_*.json` de versões anteriores à 0.7.0 são migrados automaticamente para `.bin`
> criptografado na primeira execução, e os originais são excluídos.

### 5.4 Flags de recursos e o interruptor de exclusão (kill-switch)

| Variável de ambiente | Padrão | Efeito |
|---|---|---|
| `MCP_ENABLE_TEAMS` | auto-detecção (stdio) | Força as ferramentas do Teams a ativadas/desativadas. Se ausente → auto-habilitado para valores de tenant corporativo (`common`, `organizations`, um GUID/domínio), desativado para `consumers`. |
| `MCP_ENABLE_SHAREPOINT` | auto-detecção (stdio) | Força as ferramentas do SharePoint a ativadas/desativadas, mesma lógica de auto-detecção. |
| `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` | `false` | Registra as ferramentas de transcrição/gravação do Teams + solicita seus escopos. Somente opt-in explícito. |
| `MCP_ENABLE_TEAMS_AI_INSIGHTS` | `false` | Registra as ferramentas de insight de IA do Copilot no Teams + solicita `OnlineMeetingAiInsight.Read.All`. Somente opt-in explícito. |
| `MCP_DISABLE_DELETION_TOOLS` | `false` | **Kill-switch:** quando verdadeiro (truthy), suprime o registro de todas as ferramentas de exclusão definitiva. |

O interruptor de exclusão (`MCP_DISABLE_DELETION_TOOLS=1`) remove estas ferramentas de exclusão
definitiva (hard-delete): `delete_email`, `bulk_delete_emails`, `delete_event`, `delete_contact`,
`delete_folder`, `delete_drive_item`, `delete_list_item`, `remove_ms_profile`. As variantes
recuperáveis (`trash_email`, `bulk_trash_emails`, `move_or_copy_item`) permanecem.

Valores verdadeiros (truthy) para qualquer flag: `1`, `true`, `yes`, `on` (sem diferenciar
maiúsculas/minúsculas).

### 5.5 Uso multiconta

Configure múltiplos perfis e depois direcione um deles em qualquer chamada de ferramenta passando
`profile`:

```
list_emails(folder="Inbox", profile="work")
search_drive(query="Q1 report", profile="personal")
```

Omitir `profile` usa o perfil padrão. Após o primeiro login interativo por perfil, o MSAL renova os
tokens silenciosamente.

---

## 6. Configuração: http (servidor multiusuário)

### 6.1 Referência de variáveis de ambiente (modo http)

Toda variável `MCP_*` lida no modo http. O modo stdio ignora todos os valores `MCP_HTTP_*` /
`MCP_AUTH_*`.

| Variável | Finalidade | Padrão | Obrigatória no modo http? |
|---|---|---|---|
| `MCP_TRANSPORT` | Seleciona o modo. Deve ser `stdio` ou `http`. | `stdio` | Defina como `http`. |
| `MCP_HTTP_HOST` | Host de bind **dentro** do processo/container. Use `0.0.0.0` atrás de um proxy/Docker. | `127.0.0.1` | Não. |
| `MCP_HTTP_PORT` | Porta de bind. Deve ser 1–65535. | `8000` | Não. |
| `MCP_BASE_URL` | A **URL HTTPS pública** que os clientes acessam (a URL do seu proxy, não o endereço de bind). Deve começar com `http://` ou `https://`. | — | **Sim.** |
| `MCP_AUTH_CLIENT_ID` | Application (client) ID do app confidencial (§4.2). | — | **Sim.** |
| `MCP_AUTH_CLIENT_SECRET` | O valor do client secret (§4.2). | — | **Sim.** |
| `MCP_AUTH_TENANT_ID` | O **GUID do tenant** do seu diretório. Pseudo-tenants/domínios são rejeitados na inicialização. | — | **Sim.** |
| `MCP_AUTH_REQUIRED_SCOPE` | Nome do escopo de API personalizado de "Expose an API". | `mcp-access` | Não (somente se você renomeou). |
| `MCP_HTTP_STATELESS` | Streamable HTTP sem estado (stateless). A restrição de worker único ainda se aplica. | `false` | Não. |
| `MCP_RATE_LIMIT_RPS` | Limite de requisições/segundo por usuário (burst = 2×). `0`/negativo desativa. | `10` | Não. |
| `MCP_STATS_TOKEN` | Habilita as rotas de observabilidade (vazio = desativado). | *(vazio)* | Não. |
| `MCP_ENABLE_FILE_UPLOAD` | App de upload de arquivos por arrastar-e-soltar ([§6.8](#68-upload-de-arquivos)). O valor explícito vence. | ligado no http | Não. |
| `MCP_UPLOAD_MAX_MB` | Tamanho máx. (MB) por arquivo enviado. Inteiro positivo. | `10` | Não. |
| `MCP_UPLOAD_GLOBAL_BUDGET_MB` | Limite global (MB) da pegada base64 de todos os uploads entre todos os usuários. Inteiro positivo. | `1024` | Não. |

As feature flags (`MCP_ENABLE_TEAMS`, `MCP_ENABLE_SHAREPOINT`, `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS`,
`MCP_ENABLE_TEAMS_AI_INSIGHTS`) e `MCP_DISABLE_DELETION_TOOLS` se comportam como no stdio —
**exceto** que o fallback de auto-detecção de Teams/SharePoint não se aplica (não há um único
perfil para inspecionar), então essas flags precisam ser definidas **explicitamente** no modo http,
ou os serviços permanecem desativados.

> **A inicialização falha rapidamente (fail-fast).** No modo http o servidor valida a configuração
> antes de fazer o bind e aborta com um erro claro e detalhado se `MCP_BASE_URL`,
> `MCP_AUTH_CLIENT_ID`, `MCP_AUTH_CLIENT_SECRET` ou `MCP_AUTH_TENANT_ID` estiver ausente ou malformado
> (veja [§9.1](#91-erros-de-configuração-na-inicialização-modo-http)).

### 6.2 Início rápido em bare-metal

```bash
export MCP_TRANSPORT=http
export MCP_BASE_URL=https://mcp.example.com          # sua URL HTTPS pública (atrás de um proxy)
export MCP_AUTH_CLIENT_ID=your-confidential-client-id
export MCP_AUTH_CLIENT_SECRET=your-client-secret
export MCP_AUTH_TENANT_ID=your-directory-tenant-guid  # Entra > Overview > Directory (tenant) ID
# Opcional: habilitar serviços (precisa ser explícito no modo http)
# export MCP_ENABLE_TEAMS=true
# export MCP_ENABLE_SHAREPOINT=true
uv run mcp-microsoft
```

O processo faz bind em `MCP_HTTP_HOST:MCP_HTTP_PORT` (padrão `127.0.0.1:8000`) e serve o endpoint
MCP em `/mcp`, além de um `GET /health` não autenticado.

### 6.3 Início rápido com Docker e Compose

O `Dockerfile` fornecido constrói uma imagem exclusiva para http (ele fixa `MCP_TRANSPORT=http`,
`MCP_HTTP_HOST=0.0.0.0`, `MCP_HTTP_PORT=8000`, executa como usuário não root, e tem um healthcheck em
`/health`). Os valores de autenticação intencionalmente **não** são embutidos na imagem — forneça-os
em tempo de execução.

```bash
cp .env.template .env
# Preencha a seção "Remote server (http) mode" do .env:
#   MCP_BASE_URL, MCP_AUTH_CLIENT_ID, MCP_AUTH_CLIENT_SECRET, MCP_AUTH_TENANT_ID
docker compose up -d
curl http://localhost:8000/health      # -> {"status":"ok","transport":"http"}
```

Ou sem o Compose:

```bash
docker build -t mcp-microsoft:0.8.0 .
docker run --rm -p 8000:8000 --env-file .env mcp-microsoft:0.8.0
```

> `.env` está no `.gitignore` — **nunca faça commit dele.** Prefira um gerenciador de segredos (ex.:
> Azure Key Vault) que injete `MCP_AUTH_CLIENT_SECRET` e `MCP_STATS_TOKEN` como variáveis de
> ambiente no momento do deploy, em vez de armazená-los em disco por longo prazo.

### 6.4 Proxy reverso e TLS

O servidor fala **HTTP puro** e nunca termina TLS. Coloque um proxy reverso na frente dele e aponte
`MCP_BASE_URL` para a **URL HTTPS pública** do proxy.

> **`MCP_BASE_URL` precisa corresponder exatamente** — esquema, host, porta e qualquer prefixo de
> caminho — ao que os clientes MCP acessam e ao que você registrou como base da redirect-URI no
> Azure. Uma divergência quebra o redirect do OAuth e as verificações de audience/issuer do JWT.

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
        # Streamable HTTP é de longa duração; relaxe buffering/timeouts.
        proxy_buffering off;
        proxy_read_timeout 3600s;
    }
}
```

Em ambos os casos, defina `MCP_BASE_URL=https://mcp.example.com`. O `docker-compose.yml` também
documenta um exemplo baseado em labels do Traefik no bloco de comentários final.

> **Limite a taxa (throttle) dos endpoints OAuth no proxy.** O limitador de taxa embutido (§8) cobre
> **apenas** o tráfego autenticado em `/mcp`. Os endpoints OAuth não autenticados — `/authorize`,
> `/token`, `/register`, `/auth/callback` — **não** têm limite de taxa aplicado pelo servidor.
> Adicione throttling por IP para eles no proxy se o servidor estiver exposto à internet.

### 6.5 Como os clientes se conectam

Direcione um cliente MCP com suporte a OAuth para `https://your-host/mcp`. O fluxo:

1. O cliente busca `/.well-known/oauth-protected-resource` (RFC 9728) e descobre os endpoints OAuth
   — o `AzureProvider` atua como um proxy OAuth na frente do Entra (que não tem Dynamic Client
   Registration).
2. O usuário completa o login da Microsoft no navegador e consente (ou o administrador do tenant já
   consentiu para todos).
3. O cliente então envia um **bearer token** derivado do Entra em cada requisição; o servidor o
   valida e o troca **On-Behalf-Of** por um token do Graph com o escopo daquele usuário, a cada
   chamada.

Na **primeira conexão**, o usuário passa por um login e consentimento padrão do OAuth da Microsoft;
depois disso, o cliente lida com a renovação do token. Clientes sem suporte a OAuth ou Streamable
HTTP não podem usar o modo http.

### 6.6 http vs. stdio — diferenças de comportamento

| Aspecto | stdio | http |
|---|---|---|
| **Quantidade de ferramentas** (todos os serviços opcionais ativos) | **96** | **88** |
| **Ferramentas de gerenciamento de perfil** | Registradas | **Não registradas** — `add_ms_profile`, `list_ms_profiles`, `remove_ms_profile`, `authenticate_ms_profile`, `set_default_ms_profile` não existem. |
| **Argumento `profile`** | Respeitado | **Inerte** — aceito por compatibilidade, mas silenciosamente ignorado; a identidade sempre vem do bearer token. |
| **Ferramentas de disco local** | Disponíveis | `download_file`, `download_from_site`, `teams_download_meeting_recording` **não são registradas**. `upload_file`/`upload_to_site` rejeitam `local_path` (use `content_base64`); `download_attachment`/`get_contact_photo` rejeitam `save_path` (o conteúdo é retornado inline). |
| **Feature flags** | Env ou auto-detecção corporativa | Env **apenas** (explícito). |
| **Kill-switch de exclusão** | Funciona | Funciona identicamente. |
| **Rate limiting / audit logging** | Nenhum (chamador local único) | Ativo por padrão. |
| **Detalhes de erro nas respostas** | Completos | Mascarados (`mask_error_details=True`). |
| **App de upload de arquivos** ([§6.8](#68-upload-de-arquivos)) | Desligado (usuários locais têm `local_path`) | **Ligado por padrão** — adiciona uma UI de arrastar-e-soltar mais 3 ferramentas visíveis ao modelo (`file_manager`, `list_files`, `read_file`). |

A diferença de 87 vs. 95 é exatamente as **5 ferramentas de gerenciamento de perfil** mais as **3
ferramentas de download em disco local** que o modo http omite. (Esses números excluem o app
opcional de upload; quando habilitado, ele adiciona mais 3 ferramentas visíveis ao modelo — veja a
[§6.8](#68-upload-de-arquivos).)

### 6.7 Implantação sem exclusão (no-delete)

Para equipes que precisam garantir que o assistente jamais consiga excluir nada permanentemente,
execute o servidor com o kill-switch de exclusão:

```bash
export MCP_DISABLE_DELETION_TOOLS=1
```

Essa é a receita inteira — funciona de forma idêntica nos modos stdio e http e é aplicada no
**momento do registro**: as ferramentas de exclusão permanente (`delete_email`,
`bulk_delete_emails`, `delete_event`, `delete_contact`, `delete_folder`, `delete_drive_item`,
`delete_list_item`) simplesmente não existem no servidor, então nenhum cliente, prompt ou
comportamento do modelo consegue invocá-las. As variantes recuperáveis (`trash_email`,
`bulk_trash_emails`, `move_or_copy_item`) permanecem disponíveis, então a limpeza do dia a dia
continua funcionando — os itens vão para Itens Excluídos / a lixeira em vez de desaparecer.

> **O switch vale para o servidor inteiro, não por usuário.** Um servidor http tem um único
> conjunto de ferramentas para todos os usuários conectados. Se alguns usuários precisam de
> exclusão completa e outros não podem tê-la, execute **duas instâncias** — uma completa e uma
> no-delete — e direcione cada grupo de usuários para a URL correta.

**Instâncias lado a lado (completa + no-delete)** podem compartilhar um único App Registration do
Azure: um registro de aplicativo aceita múltiplas URIs de redirecionamento, então adicione as duas
URLs de callback (por exemplo, `https://mcp.example.com/auth/callback` **e**
`https://mcp-nodelete.example.com/auth/callback`) à mesma plataforma Web e reutilize o mesmo client
ID/segredo/GUID de tenant nas duas implantações. Cada instância precisa do seu próprio
`MCP_BASE_URL` (correspondendo exatamente à sua URL pública) e, se a observabilidade estiver
habilitada, do seu próprio `MCP_STATS_TOKEN`. Um segundo serviço pronto para descomentar, seguindo
esse padrão, acompanha o `docker-compose.yml`.

Para usuários de **stdio / Claude Desktop**, o equivalente é o bundle no-delete pré-compilado
(`mcp-microsoft-nodelete.mcpb`) ou a opção **Disable Permanent-Delete Tools** (Desativar
ferramentas de exclusão permanente) no instalador MCPB (veja a [§5.1](#51-mcpb-claude-desktop--recomendado)).

### 6.8 Upload de arquivos

**O que é.** No modo http não há disco compartilhado entre o chamador e o servidor, então
`local_path` é rejeitado e passar um arquivo como base64 força todo o seu conteúdo pela janela de
contexto do modelo. O **app de upload de arquivos** (baseado no provider `FileUpload` do FastMCP,
`fastmcp[apps]`) resolve isso: expõe uma UI de arrastar-e-soltar. Os arquivos que o usuário solta
vão **direto para o servidor**, contornando o contexto do modelo, e ficam em uma área de upload por
usuário. As ferramentas de upload então os consomem **por nome** via um novo parâmetro
`uploaded_file` em `upload_file` (OneDrive) e `upload_to_site` (SharePoint) — mutuamente exclusivo
com `local_path`/`content_base64`; o nome do arquivo assume por padrão o nome enviado.

**Requisito do cliente.** O cliente precisa suportar **MCP Apps** (recursos de UI interativa) — por
exemplo, o Claude Desktop. Clientes sem suporte a MCP Apps enxergam as ferramentas visíveis ao
modelo `list_files`/`read_file`, mas não conseguem abrir a UI de arrastar-e-soltar. O recurso vem
**ligado por padrão no modo http e desligado no stdio** (usuários locais já têm `local_path`);
sobrescreva nos dois sentidos com `MCP_ENABLE_FILE_UPLOAD`.

**Cotas e limites** (por usuário conectado; aplicados em memória, limitados contra abuso):

| Limite | Valor | Observações |
|---|---|---|
| Tamanho máximo por arquivo | `MCP_UPLOAD_MAX_MB` (padrão **10 MB**) | Rejeitado antes do armazenamento; deve ser um inteiro positivo. |
| Máx. de arquivos por usuário | **20** | Upload acima da cota é rejeitado com mensagem clara; sobrescrever um nome reutiliza o slot. |
| Máx. de bytes por usuário | **100 MB** no total | A cota por usuário usa o tamanho **decodificado** real, não o informado pelo cliente. |
| Orçamento global de upload | `MCP_UPLOAD_GLOBAL_BUDGET_MB` (padrão **1024 MB**) | Limita a pegada **codificada** (base64) entre todos os usuários; um armazenamento acima do orçamento é rejeitado mesmo que o usuário esteja dentro da própria cota. |
| Usuários distintos rastreados | **1000** | Áreas de upload menos usadas recentemente são removidas ao exceder o limite. |
| TTL de área ociosa | **2 h** | Áreas ociosas são podadas de forma preguiçosa no próximo upload/listagem. |

Os uploads vivem **apenas na memória do processo do servidor**, escopados pelo `oid` do Entra do
chamador (estável entre reconexões e no modo stateless), com fallback para o `sub`; uma requisição
sem **nenhum** dos dois é recusada em vez de compartilhar um bucket. Os uploads são perdidos ao
reiniciar. O **conteúdo dos arquivos nunca é registrado em log**. As chamadas às ferramentas do
provider passam pela mesma middleware de rate-limit / auditoria / métricas que qualquer outra
ferramenta — incluindo a ferramenta de backend `store_files`, que é acessível pelo seu nome com hash
(não é exclusiva da UI) e passa pelas mesmas verificações de middleware, por arquivo e de cota.

**Configuração:**

| Variável | Propósito | Padrão |
|---|---|---|
| `MCP_ENABLE_FILE_UPLOAD` | Habilita/desabilita o app de upload. O valor explícito vence. | ligado no http, desligado no stdio |
| `MCP_UPLOAD_MAX_MB` | Tamanho máx. (MB) de cada arquivo enviado. Inteiro positivo. | `10` |
| `MCP_UPLOAD_GLOBAL_BUDGET_MB` | Limite global (MB) da pegada base64 de todos os uploads entre todos os usuários. Inteiro positivo. | `1024` |

---

## 7. Observabilidade

O modo http pode expor métricas leves de tráfego/uso, em processo. Isso é **desativado por padrão**
e só é ativado quando `MCP_STATS_TOKEN` é definido com um segredo não vazio. Quando não definido, as
rotas nem são registradas (uma linha de log observa isso na inicialização). `GET /health` está
sempre aberto e não é afetado.

### 7.1 As três rotas

| Rota | Retorna | Uso |
|---|---|---|
| `GET /metrics` | Exposição em texto do Prometheus (`text/plain; version=0.0.4`) | Alvo de scrape do Prometheus/Grafana. |
| `GET /stats` | Snapshot em JSON: uptime/totais, tráfego por minuto (últimos 60 min), latência p50/p95/média por ferramenta, atividade por usuário. | Dashboards programáticos, `curl` ad-hoc. |
| `GET /dashboard` | Uma única página HTML autocontida (CSS/JS inline, sem requisições externas) que consulta `/stats` a cada 10s. | Visualizar em um navegador. |

### 7.2 Autenticação

Todas as três rotas exigem o token, apresentado de uma das duas formas:

- `Authorization: Bearer <MCP_STATS_TOKEN>`, ou
- HTTP **Basic**, com qualquer nome de usuário e o token como **senha** (assim, um navegador abrindo
  `/dashboard` recebe um prompt de login nativo).

A comparação é resistente a timing attacks (timing-safe); o token nunca é logado. As três respostas
são servidas com `Cache-Control: no-store`.

```bash
# curl com Bearer
curl -H "Authorization: Bearer $MCP_STATS_TOKEN" https://mcp.example.com/metrics
curl -H "Authorization: Bearer $MCP_STATS_TOKEN" https://mcp.example.com/stats

# Navegador: abra https://mcp.example.com/dashboard e informe qualquer usuário + o token como senha
```

### 7.3 Configuração de scrape do Prometheus

```yaml
scrape_configs:
  - job_name: mcp-microsoft
    scheme: https
    metrics_path: /metrics
    authorization:
      type: Bearer
      credentials: "<MCP_STATS_TOKEN>"      # ou use basic_auth com o token como senha
    static_configs:
      - targets: ["mcp.example.com"]
```

### 7.4 Métricas emitidas

Séries globais do processo e por ferramenta (deliberadamente **não há séries com label por usuário**
— a cardinalidade por identidade é ilimitada e é um antipadrão conhecido do Prometheus; o detalhe
por usuário vive em `/stats` e `/dashboard`):

- `mcp_uptime_seconds` (gauge)
- `mcp_calls_total`, `mcp_errors_total` (counters)
- `mcp_users_tracked` (gauge), `mcp_users_evicted_total` (counter)
- `mcp_unknown_tool_calls_total`, `mcp_tools_evicted_total` (counters)
- por ferramenta: `mcp_tool_calls_total{tool}`, `mcp_tool_errors_total{tool}`,
  `mcp_tool_duration_ms{tool,stat="p50|p95|avg"}`

> **Reiniciar zera tudo.** As métricas estão em memória, sem persistência, e cobrem apenas o worker
> único em que o processo é executado. Um restart zera tudo.

> **Mantenha isso privado.** O token é a única barreira. Trate essas rotas como dados operacionais
> sensíveis e restrinja-as adicionalmente (allowlist de IP / autenticação separada) no seu proxy
> reverso, se o servidor estiver exposto à internet. Deixar `MCP_STATS_TOKEN` indefinido as
> desativa por completo.

---

## 8. Checklist de segurança

| Item | O que fazer | Por quê |
|---|---|---|
| **Somente TLS** | Termine o TLS em um proxy reverso; defina `MCP_BASE_URL` como a URL HTTPS pública. | Os fluxos OAuth e toda chamada ao Graph carregam bearer tokens — nunca os envie por HTTP puro. |
| **GUID do tenant** | `MCP_AUTH_TENANT_ID` precisa ser o GUID concreto do tenant. | Pseudo-tenants/domínios são rejeitados na inicialização; o pinning do issuer significa que somente o GUID valida tokens reais. |
| **Armazenamento de segredos** | Mantenha `MCP_AUTH_CLIENT_SECRET` / `MCP_STATS_TOKEN` em um gerenciador de segredos ou injeção via variáveis de ambiente, não em arquivos versionados. | São senhas. `.env` está no `.gitignore`. |
| **Rotação de segredos** | Rotacione o client secret antes de sua expiração (≤24 meses): adicione um novo segredo, atualize a variável de ambiente, reinicie, exclua o antigo. | Ambos os segredos são válidos simultaneamente, então não há downtime. |
| **Rate limiting** | Mantenha `MCP_RATE_LIMIT_RPS` ativo (padrão `10`/s por usuário, burst `2×`). A chave por usuário é `tid:oid`; o armazenamento de buckets é limitado por LRU a 10.000 chaves e podado por ociosidade após 900s. | Evita que um usuário monopolize os demais; memória limitada. Exceder o limite gera JSON-RPC `-32000`. |
| **Throttling no proxy** | Aplique throttling em `/authorize`, `/token`, `/register`, `/auth/callback` no proxy. | O limitador embutido cobre apenas o tráfego autenticado em `/mcp`. |
| **Logs de auditoria** | Uma linha por chamada de ferramenta: nome da ferramenta, `oid` + `preferred_username` do chamador, duração, resultado. | Nunca registra argumentos, resultados ou o próprio token. |
| **Kill-switch de exclusão** | Defina `MCP_DISABLE_DELETION_TOOLS=1` onde exclusões definitivas devem ser impossíveis. | Remove todas as ferramentas de exclusão definitiva; as variantes recuperáveis permanecem. |
| **Restrição das ferramentas de disco** | Automática no modo http — ferramentas de download para disco não são registradas; `save_path`/`local_path` são rejeitados. | O disco do servidor não é o disco do chamador. |
| **Worker único** | Execute exatamente um worker/réplica no modo http. | O armazenamento de clients do proxy OAuth e o cache OBO por usuário estão em memória no processo; múltiplos workers dividem o estado e quebram sessões. Escalonamento horizontal precisa do `client_storage` externo do FastMCP (não implementado aqui). |
| **Container não root** | A imagem já é executada como usuário não root. | Não há motivo para o processo escrever fora de seu venv/tmp. |

---

## 9. Solução de problemas

### 9.1 Erros de configuração na inicialização (modo http)

O modo http valida a configuração antes de fazer o bind. Valores ausentes/inválidos abortam com:

```
Cannot start http transport — fix the following configuration problems:
  - <problem 1>
  - <problem 2>
```

Mensagens de erro reais:

- `MCP_BASE_URL is required in http mode`
- `MCP_BASE_URL must start with http:// or https:// (got '...')`
- `MCP_AUTH_CLIENT_ID is required in http mode`
- `MCP_AUTH_CLIENT_SECRET is required in http mode`
- `MCP_AUTH_TENANT_ID is required in http mode`
- `MCP_HTTP_PORT must be between 1 and 65535 (got ...)`
- A rejeição do GUID do tenant (na íntegra, em inglês):
  > `MCP_AUTH_TENANT_ID must be your directory's tenant GUID (8-4-4-4-12 hexadecimal), not '...'. Copy the Directory (tenant) ID from Azure Portal -> Microsoft Entra ID -> Overview. Pseudo-tenants ('organizations', 'common', 'consumers') and verified domains ('contoso.onmicrosoft.com') are rejected because fastmcp's AzureProvider validates the token 'iss' claim against a single literal issuer URL built from this value, and real Entra tokens always carry the concrete tenant GUID -- so a pseudo-tenant or domain never matches and every request fails authentication.`
  >
  > (Tradução livre: `MCP_AUTH_TENANT_ID` precisa ser o GUID do tenant do seu diretório
  > (hexadecimal `8-4-4-4-12`), não '...'. Copie o Directory (tenant) ID em Azure Portal ->
  > Microsoft Entra ID -> Overview. Pseudo-tenants ('organizations', 'common', 'consumers') e
  > domínios verificados ('contoso.onmicrosoft.com') são rejeitados porque o `AzureProvider` do
  > fastmcp valida a claim 'iss' do token contra uma única URL de issuer literal construída a
  > partir desse valor, e tokens reais do Entra sempre carregam o GUID concreto do tenant — logo,
  > um pseudo-tenant ou domínio nunca corresponde e toda requisição falha na autenticação.)

**Correção:** forneça o(s) valor(es) ausente(s); para o último caso, coloque o GUID concreto do
tenant em `MCP_AUTH_TENANT_ID`.

### 9.2 Erro de digitação em `MCP_TRANSPORT`

Se `MCP_TRANSPORT` não for nem `stdio` nem `http`, a inicialização aborta:

```
Cannot start mcp-microsoft — MCP_TRANSPORT must be 'stdio' or 'http' (got 'htttp')
```

**Correção:** corrija o valor (ou remova a variável para usar o padrão `stdio`).

### 9.3 `AADSTS65001` — consentimento do administrador necessário

O Graph retorna `AADSTS65001: The user or administrator has not consented` (o usuário ou
administrador não consentiu). No modo stdio, o servidor exibe uma URL de consentimento do
administrador:

```
https://login.microsoftonline.com/common/adminconsent?client_id={client_id}
```

**Correção:** um administrador do tenant precisa conceder o consentimento (compartilhe seu
Application (client) ID). Para escopos restritos a administrador, como `Sites.ReadWrite.All`, isso
é obrigatório. No modo http, conceda o consentimento do administrador em nível de tenant no registro
confidencial (§4.2, passo 9).

### 9.4 `401` de um cliente OAuth (modo http)

O token do cliente está ausente ou não tem o escopo exigido. Verificações:

- O nome do escopo em **Expose an API** do Registro de Aplicativo corresponde a
  `MCP_AUTH_REQUIRED_SCOPE` (padrão `mcp-access`), e o cliente o solicita.
- `"requestedAccessTokenVersion": 2` está definido no manifesto do app (§4.2, passo 6) — caso
  contrário, o formato das claims do token (`scp`/`oid`/`iss`) não será validado.
- `MCP_BASE_URL` corresponde exatamente à base da redirect-URI e à URL que o cliente acessa.
- `MCP_AUTH_TENANT_ID` é o GUID concreto (o pinning do issuer rejeita divergências).

### 9.5 Erro de limite de taxa (modo http)

Chamadas acima do limite geram `RateLimitError` — um `McpError` com código JSON-RPC **`-32000`** e a
mensagem `Rate limit exceeded for client: <tid:oid>`.

**Correção:** reduza a frequência das chamadas, ou aumente/desative o limite via
`MCP_RATE_LIMIT_RPS` (`0` ou negativo desativa).

### 9.6 Falha no login interativo do stdio / ambiente headless

Se o fluxo interativo de navegador não estiver disponível (MCPB, SSH, containers), o servidor recorre
ao fluxo de **device code** e registra o código + a URL de verificação como um aviso no log. Conclua
o login em qualquer dispositivo usando esse código.

### 9.7 Outros erros do Azure (stdio)

- `AADSTS50011: redirect URI does not match` → adicione `http://localhost` em **Mobile and desktop
  applications** (não em Web, nem em `https://`).
- `AADSTS700016: Application not found in the directory` → o tenant ID não corresponde à conta. Use
  `consumers` para uma conta pessoal, ou o Directory (tenant) ID para uma conta corporativa.
- **Ferramentas de SharePoint/Teams não aparecem** → a flag correspondente está desativada, ou a
  conta é pessoal (SharePoint e Teams são exclusivos de contas corporativas/educacionais).

### 9.8 Para onde vão os logs

O servidor registra logs usando o logging padrão do Python (stderr para o subprocesso stdio;
stdout/stderr do container para Docker — visualize com `docker compose logs -f`). Eventos de
auditoria e de rate-limit aparecem ali no modo http. As rotas de observabilidade `/stats` e
`/dashboard` exibem métricas em processo, em tempo real.

---

## 10. Limites e FAQ

- **Contas pessoais são somente stdio.** OBO e escopos de API personalizados não têm suporte
  confiável para contas Microsoft de consumidor, então o modo http é exclusivo para
  corporativa/educacional. Use stdio para Outlook.com / Live.
- **Multi-tenant não tem suporte (http).** O servidor tem como alvo um único GUID de tenant
  concreto; o issuer é fixado (pinned) a uma URL literal construída a partir dele. Implantações
  multi-tenant são trabalho futuro (precisariam pular a validação de issuer, além de uma authority
  OBO por tenant).
- **Métricas são zeradas ao reiniciar.** Somente em memória, worker único, sem persistência.
- **Somente worker único (http).** O armazenamento de clients do proxy OAuth e o cache OBO por
  usuário estão em processo. Não execute múltiplos workers/réplicas sem conectar o `client_storage`
  externo do FastMCP (não implementado aqui). `MCP_HTTP_STATELESS` não altera isso.
- **Atualizações do MCPB.** O pacote MCPB do stdio é atualizado instalando um `.mcpb` mais recente;
  as configurações (`user_config`) são preservadas pelo Claude Desktop. Reinicie completamente o
  Claude Desktop após atualizar.
- **Throttling do Teams.** Os endpoints do Graph para Teams limitam a taxa em torno de ~4 req/s e
  podem retornar 429; alguns endpoints de listagem de reuniões exigem um `$filter` OData e podem
  retornar 400 em tenants sem esse suporte.

---

## 11. Apêndice: variáveis de ambiente e inventário de ferramentas

### 11.1 Tabela completa de variáveis de ambiente

| Variável | Modo | Padrão | Descrição |
|---|---|---|---|
| `MS365_CLIENT_ID` | stdio | — | Client ID para o perfil `default` de bootstrap. Não é um segredo. |
| `MS365_TENANT_ID` | stdio | `common` | Tenant para o perfil de bootstrap (`common`/`consumers`/GUID/domínio). |
| `MS365_CREDENTIALS_DIR` | stdio | `~/.microsoft-mcp/` | Diretório para `profiles.json` e caches de token. |
| `MCP_ENABLE_TEAMS` | ambos | auto-detecção (stdio) / desativado (http) | Força as ferramentas do Teams a ativadas/desativadas. |
| `MCP_ENABLE_SHAREPOINT` | ambos | auto-detecção (stdio) / desativado (http) | Força as ferramentas do SharePoint a ativadas/desativadas. |
| `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` | ambos | `false` | Ferramentas de transcrição/gravação do Teams (opt-in explícito). |
| `MCP_ENABLE_TEAMS_AI_INSIGHTS` | ambos | `false` | Ferramentas de insight de IA do Copilot no Teams (opt-in explícito; precisa de licença Copilot). |
| `MCP_DISABLE_DELETION_TOOLS` | ambos | `false` | Kill-switch: suprime todas as ferramentas de exclusão definitiva. |
| `MCP_TRANSPORT` | ambos | `stdio` | `stdio` ou `http`. |
| `MCP_HTTP_HOST` | http | `127.0.0.1` | Host de bind dentro do processo/container. |
| `MCP_HTTP_PORT` | http | `8000` | Porta de bind (1–65535). |
| `MCP_HTTP_STATELESS` | http | `false` | Streamable HTTP sem estado. |
| `MCP_BASE_URL` | http | — | URL HTTPS pública (obrigatória). |
| `MCP_AUTH_CLIENT_ID` | http | — | Client ID confidencial (obrigatório). |
| `MCP_AUTH_CLIENT_SECRET` | http | — | Client secret (obrigatório, secreto). |
| `MCP_AUTH_TENANT_ID` | http | — | GUID do tenant (obrigatório; pseudo-tenants/domínios rejeitados). |
| `MCP_AUTH_REQUIRED_SCOPE` | http | `mcp-access` | Nome do escopo de API personalizado. |
| `MCP_RATE_LIMIT_RPS` | http | `10` | Requisições/segundo por usuário (burst 2×); `0`/negativo desativa. |
| `MCP_STATS_TOKEN` | http | *(vazio)* | Habilita as rotas de observabilidade (secreto). |

### 11.2 Inventário de ferramentas por modo

As contagens assumem todos os serviços opcionais habilitados.

| Grupo | stdio | http | Observações |
|---|---|---|---|
| Mail | 26 | 26 | Inclui extração de texto de anexos PDF/texto no servidor. |
| Calendar | 10 | 10 | |
| OneDrive | 8 | 7 | `download_file` omitido no http. |
| SharePoint | 13 | 12 | `download_from_site` omitido no http. |
| Contacts | 8 | 8 | `get_contact_photo` rejeita `save_path` no http. |
| Teams | 25 | 24 | `teams_download_meeting_recording` omitido no http. |
| Gerenciamento de perfil | 5 | 0 | Não registrado no http. |
| Utilitários de serviço | 1 | 1 | `list_enabled_services`. |
| **Total** | **96** | **88** | |

**Ferramentas totalmente omitidas no modo http:** `add_ms_profile`, `list_ms_profiles`,
`remove_ms_profile`, `authenticate_ms_profile`, `set_default_ms_profile` (gerenciamento de perfil);
`download_file`, `download_from_site`, `teams_download_meeting_recording` (downloads em disco
local).

**Ferramentas que rejeitam parâmetros de disco no modo http** (mas permanecem registradas):
`upload_file` / `upload_to_site` rejeitam `local_path` (use `content_base64`);
`download_attachment` / `get_contact_photo` rejeitam `save_path` (o conteúdo é retornado inline).

---

*Guia do mcp-microsoft v0.8.0. Para o passo a passo do Azure com capturas de tela, veja
[`docs/azure-setup.md`](azure-setup.md); para os quickstarts condensados, veja o
[README](../README.md).*
