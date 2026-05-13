# Wordembedded

A React + .NET app for document collaboration backed by **SharePoint Embedded (SPE)**, the
headless file-storage partition that lives inside a Microsoft 365 tenant. The React frontend owns
the entire UX; Microsoft 365 owns storage, identity, permissions, and Office co-authoring.

The product brief is in `docs/project/PRD.md`. The brand guardrails are in `PRODUCT.md`. Read those
first — many design choices below are downstream of those documents.

## Layout

```
src/
  infra/           Provisioning artifacts: outputs.json, sync-config.ps1, README.md
  backend/         .NET Aspire AppHost + ASP.NET Core 9 minimal API (OBO → Microsoft Graph)
    Wordembedded.sln
    Wordembedded.AppHost/
    Wordembedded.ServiceDefaults/
    Wordembedded.Api/
  frontend/        Next.js 16 (App Router) + React 19 + TypeScript + MSAL.js + Fluent UI v9
```

## Prerequisites

- **.NET 9 SDK** (`dotnet --version` ≥ 9.0)
- **Aspire workload** (`dotnet workload install aspire`)
- **Node.js 20+** and **npm 10+**
- **Azure CLI** (`az --version`), signed in to the M365 tenant where SPE is enabled
- **PowerShell 7+** for the provisioning helpers
- **PnP.PowerShell** module (`Install-Module PnP.PowerShell`) for container-type management
- **Microsoft.Graph PowerShell** (`Install-Module Microsoft.Graph`) for container creation

## First-time setup (one-time)

The SharePoint Embedded resources, Entra app registration, certificate, and tenant consent must be
provisioned **once per environment**. Step-by-step instructions are in
[`src/infra/README.md`](./src/infra/README.md).

After provisioning, `src/infra/outputs.json` contains the IDs the apps need.

Push those IDs into the backend and frontend config:

```powershell
pwsh src/infra/sync-config.ps1
```

That writes:
- `src/backend/Wordembedded.Api/appsettings.Development.json` (tenantId, clientId, audience, cert thumbprint, container-type id, webhook secret)
- `src/frontend/.env.local` (NEXT_PUBLIC_* values)

Both files are gitignored.

## Running locally

Two terminals.

### Backend (Aspire AppHost)

```powershell
cd src/backend
dotnet run --project Wordembedded.AppHost
```

The Aspire dashboard opens; the API listens on `https://localhost:7071`. Swagger is at
`https://localhost:7071/swagger`.

### Frontend (Next.js dev server)

```powershell
cd src/frontend
npm install
npm run dev
```

Open `http://localhost:3000`. MSAL redirects to Entra ID; sign in with a user in the consuming
tenant who is a member of the SPE container.

## What's wired

| Capability | Status |
|---|---|
| MSAL.js sign-in (auth-code + PKCE) | ✅ |
| OBO token exchange backend → Graph | ✅ |
| Container list + create + drill-in | ✅ |
| File browser (list/grid, breadcrumb, multi-select) | ✅ |
| Inline read-only preview (iframe + `nb=true`, refresh-on-error) | ✅ |
| Edit handoff via `window.open(webUrl, '_blank', 'noopener')` | ✅ |
| Small upload (multipart, ≤4 MB through backend) | ✅ |
| Large upload (upload session, frontend PUTs directly to Graph) | ✅ |
| Rename / move / delete / create folder | ✅ |
| Sharing: createLink (organization/specific people, view/edit) | ✅ |
| Sharing: invite (recipients + roles + message) | ✅ |
| Sharing: list + revoke permissions | ✅ |
| Versions: list + restore | ✅ |
| Microsoft Search across container content | ✅ |
| Webhook receiver (validation handshake + change notifications) | ✅ |
| Webhook subscription create + delete + auto-renew every 12h | ✅ |
| Fluent UI v9 light + dark themes; `prefers-color-scheme` + `prefers-reduced-motion` | ✅ |
| WCAG 2.2 AA focus rings | ✅ (Fluent default) |

Out of scope for phase 1 (per PRD §3 Non-Goals and §12 open questions):

- Anonymous sharing links
- Sensitivity-label management UI (Purview handles it server-side)
- Native mobile apps
- AI / Copilot grounding (PRD phase 3)
- Multi-tenant onboarding flows (PRD phase 2)

## Architectural rules — don't break these

These come straight from the PRD; ignoring them will get you stuck:

1. **No editable Word iframe.** Office for the web sets `frame-ancestors` CSP that blocks our
   domain. Editing always opens in a new tab via `driveItem.webUrl`. Don't propose, sketch, or
   build anything that puts an editable Office surface inside our iframe.
2. **Preview URLs are short-lived.** Refresh on iframe error; never cache them long-term.
3. **Webhook subscriptions expire ≤ 3 days.** `WebhookRenewalService` extends them every 12h.
   If you change that interval, make sure it stays well below the 3-day ceiling.
4. **Large uploads need upload sessions.** Single `PUT` is fine ≤ 4 MB; anything larger uses
   `/upload-session` and frontend-direct chunked PUTs to the returned `uploadUrl`.
5. **Hidden permission required.** The Entra app needs the SharePoint Online `Container.Selected`
   app role (resourceAppId `00000003-0000-0ff1-ce00-000000000000`). It doesn't show in the portal
   UI — it's added via manifest. See `src/infra/README.md`.
6. **OBO scope is `FileStorageContainer.Selected`.** Don't confuse this with the SPE-management
   scope `FileStorageContainerType.Manage.All` which is app-only and only ever used by
   provisioning, not by user-facing flows.
7. **Container permissions ≠ driveItem permissions.** A user must be a member of the container
   AND have a `driveItem` permission to access a file. Both layers are real.

## Testing the auth flow end-to-end

After `sync-config.ps1` succeeds with real values:

1. Visit `http://localhost:3000` → redirect to `/signin`.
2. Click "Sign in with Microsoft" → MSAL redirect to Entra → consent (if first time) → callback.
3. Land on `/files` showing the provisioned `Workspace` container.
4. Open it → empty file list → drag-drop a `.docx` → it uploads.
5. Click the file → preview iframe loads.
6. Click "Edit in Word" → Office for the web opens in a new tab → edit → save → close.
7. Switch back to the React app → metadata refreshes (lastModifiedDateTime ticks).

## Webhook development

To receive Graph change notifications during local dev, you need a public HTTPS URL because Graph
can't reach `localhost`. Use a tunnel (e.g. `dev tunnel`, `ngrok`, `cloudflared`):

```powershell
devtunnel host -p 7071 --allow-anonymous
```

Then update `Webhooks:NotificationUrl` in `appsettings.Development.json` to the tunnel URL (still
ending in `/api/webhooks/graph`) and re-create any active subscriptions via the API.

## Repo conventions

- Source code lives under `src/`. Nothing outside `src/` is build output.
- Don't commit secrets. `src/infra/secrets/` is gitignored; the cert PFX, its password, and
  `outputs.json` (when filled in) are local-only.
- Documentation lives under `docs/`. Update `docs/project/PRD.md` whenever a decision contradicts
  it.

## References

- PRD: [`docs/project/PRD.md`](./docs/project/PRD.md)
- Brand guardrails: [`PRODUCT.md`](./PRODUCT.md)
- Repo agent rules: [`.github/copilot-instructions.md`](./.github/copilot-instructions.md)
- SPE app architecture: <https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/app-architecture>
- SPE auth: <https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth>
- driveItem: preview: <https://learn.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-1.0>
- Microsoft.Identity.Web OBO: <https://learn.microsoft.com/en-us/entra/identity-platform/scenario-web-api-call-api-overview>
- Graph webhooks: <https://learn.microsoft.com/en-us/graph/webhooks>
