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

## Getting Started

### 1. Install dependencies

```bash
npm install
```

### 2. Configure environment

Copy `.env.example` to `.env` and fill in the required secrets:

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`
- `OPENAI_API_KEY`
- Optional overrides: `SITE_DISPLAY_NAME`, `SITE_TYPE` (`team` or `communication`), `SHAREPOINT_HOST` (required for communication sites), `PORT`, `OPENAI_MODEL`

> Grant the Azure AD app the app-only Graph permissions: `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Group.ReadWrite.All`, `User.Read.All` (plus `Team.Create` if you plan to team-enable the group).

### 3. Run locally

```bash
npm run dev
```

The server exposes:

- `GET /health` — health check
- `POST /provision` — ensure site, libraries, and permissions
- `POST /share` — generate expiring links from Deliverables
- `POST /catalog/ensure` — publish/update the Catalog page
- `GET /list` — enumerate folder contents by library/path
- `POST /agent/complete` — OpenAI tool-enabled completion endpoint

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
- Never upload Git repositories or `.env` files to SharePoint/OneDrive; keep code in GitHub and sync references via the Catalog page.
