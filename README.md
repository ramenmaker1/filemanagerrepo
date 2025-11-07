# Elion SharePoint Agent — Node Microservice

This service provisions and manages the SharePoint/OneDrive workspace for **Elion Studio**. It exposes REST endpoints and an OpenAI tool handler that can orchestrate deterministic Graph actions aligned with the PRD.

## Features

- Provision or reconcile the "Elion Studio" SharePoint workspace as either a Microsoft 365 team site or a communication site.
- Ensure document libraries exist: `Projects`, `Assets`, `Data`, `Deliverables`, `Templates`, `Legal & Finance`.
- Create dedicated security groups (`Elion-Editors`, `Elion-Viewers`) and lock down the Legal & Finance library.
- Create and refresh a modern **Catalog** page with repository, Base44, and data bucket links.
- Generate expiring share links (Deliverables only).
- List folders/files and surface the SharePoint URLs for syncing back to GitHub/Base44 metadata.
- Provide an OpenAI Responses API integration with the `ms_graph_ops` tool schema for agent workflows.
- Multi-tenant aware: map API keys to tenants, honour tenant-specific SharePoint hosts/paths, and support key rotation with multiple secret versions.

## Getting Started

### 1. Configure environment

Copy `.env.example` to `.env` and fill in the required secrets:

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`
- `OPENAI_API_KEY`
- Optional overrides: `SITE_DISPLAY_NAME`, `SITE_TYPE` (`team` or `communication`), `SHAREPOINT_HOST` (required for communication sites), `PORT`, `OPENAI_MODEL`
- Redis-backed token caching: `REDIS_URL`, `REDIS_TLS`, `REDIS_KEY_PREFIX`
- Token policy: `AUTH_TOKEN_CACHE_TTL_MINUTES` (<=55), `AUTH_TOKEN_FALLBACK_MS`
- Runtime override file (JSON): `RUNTIME_SECRETS_PATH`
- Tenant & API key sources: `TENANTS_JSON`/`TENANTS_CONFIG_PATH`, `API_KEYS_JSON`/`API_KEYS_CONFIG_PATH`

> Secrets can also be provided by path using the `*_FILE` convention (for example `AZURE_CLIENT_SECRET_FILE=/mnt/secrets/client-secret`). This is the recommended pattern when mounting Azure Key Vault secrets with the CSI driver.

### 1.1 Multi-tenant & API key configuration

Define tenants and API keys via JSON files or inline environment variables. Each tenant entry may override SharePoint site settings, and every API key is associated with a tenant and one or more secret versions for rotation.

`tenants.json`

```json
[
  {
    "id": "tenant-elion",
    "name": "Elion Studio",
    "active": true,
    "sharepoint": {
      "siteDisplayName": "Elion Studio",
      "siteType": "team",
      "host": "https://contoso.sharepoint.com",
      "sitePath": "/sites/elion-studio"
    }
  }
]
```

`api-keys.json`

```json
[
  {
    "id": "elion-primary",
    "name": "Elion Production Key",
    "tenantId": "tenant-elion",
    "roles": ["admin", "share-manager"],
    "active": true,
    "secrets": [
      { "env": "ELION_API_KEY", "active": true },
      { "file": "/mnt/rotation/elion-api-key-prev", "active": false, "expiresAt": "2024-12-01T00:00:00Z" }
    ]
  }
]
```

At least one active secret must be defined per key. During rotation, mark the new secret material as `active: true` and keep the previous secret active until clients have swapped. When rotation is complete, flip the old entry to `active: false` or remove it entirely.

Authentication accepts either `Authorization: Bearer <key>` or `X-API-Key: <key>` headers. Failures return an error envelope with an informative `code` and include `WWW-Authenticate: Bearer realm="elion-studio-api"` on 401 responses.

### 1.2 Secrets via Azure Key Vault (production)

- Mount secrets into the container with the Azure Key Vault CSI driver or Workload Identity init containers.
- Point `AZURE_CLIENT_SECRET_FILE`, `OPENAI_API_KEY_FILE`, and any custom secret references inside the API key definitions to the mounted paths.
- The service reads mounted files on startup and hashes API key material immediately; raw values are never logged.

> Grant the Azure AD app the app-only Graph permissions: `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Group.ReadWrite.All`, `User.Read.All` (plus `Team.Create` if you plan to team-enable the group).

### 2. Install dependencies

```bash
npm install
```

### 3. Run locally

```bash
npm run dev
```

The server exposes:

- `GET /healthz` — liveness health check (status, version, uptime)
- `GET /readyz` — readiness probe with Graph, SharePoint, and Redis checks
- `POST /provision` — ensure site, libraries, and permissions
- `POST /share` — generate expiring links from Deliverables
- `POST /catalog/ensure` — publish/update the Catalog page
- `GET /list` — enumerate folder contents by library/path
- `POST /agent/complete` — OpenAI tool-enabled completion endpoint

### Authentication

All non-health endpoints require an API key. Supply credentials via either `Authorization: Bearer <key>` or `X-API-Key: <key>`. Keys are scoped to a tenant and carry one or more roles; inactive tenants or secrets produce `403` responses, while unknown keys return `401` with a `WWW-Authenticate` header. The middleware also emits `X-Tenant-Id` and `X-Api-Key-Id` headers on successful responses for observability.

### OpenAI Tool Contract

The `ms_graph_ops` tool accepts the following JSON payloads:

```json
{
  "action": "ensure_site" | "ensure_libraries" | "ensure_groups_permissions" |
             "create_catalog_page" | "share_deliverable" | "list_folder" |
             "link_repo_and_base44",
  "siteType": "team" | "communication",
  "siteName": "Elion Studio",
  "libraryName": "Deliverables",
  "driveItemPath": "Projects/Alpha",
  "shareType": "view" | "edit",
  "expiresAt": "2024-05-01T00:00:00Z",
  "catalogLinks": {
    "repos": ["https://github.com/org/repo"],
    "base44": ["https://base44.app/records/123"],
    "dataBuckets": ["s3://bucket/path"]
  },
  "repoUrl": "https://github.com/org/repo",
  "base44Url": "https://base44.app/records/123",
  "sharepointUrl": "https://contoso.sharepoint.com/sites/ElionStudio/Projects/Alpha"
}
```

The agent enforces Deliverables-only sharing and will respond with human guidance for device hygiene (Files On-Demand/selective sync) per the PRD.

> Omit any of the `catalogLinks` arrays to leave a section empty; the page renders a placeholder when no links are provided.

### Notes & Limitations

- Communication site creation leverages SharePoint's `_api/SPSiteManager/Create`; configure `SHAREPOINT_HOST` with your tenant hostname (for example `https://contoso.sharepoint.com`).
- Share link expiration is subject to tenant policies; the service requests the provided `expiresAt`, but SharePoint may override it.
- Catalog page composition uses a simplified text web part — extend `catalog.ts` for richer layouts.
- OAuth client credentials flow is implemented manually with Redis-backed caching (55 minute TTL) and an automatic fallback circuit breaker after repeated failures. Redis is optional but strongly recommended for multi-instance deployments.
- Never upload Git repositories or `.env` files to SharePoint/OneDrive; keep code in GitHub and sync references via the Catalog page.
