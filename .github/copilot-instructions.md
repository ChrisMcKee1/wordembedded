# Copilot instructions — wordembedded

## Repository status

This repo now contains the product specification plus the initial .NET Aspire backend scaffold. `docs/project/PRD.md` remains the source of truth for product scope and architecture decisions.

Current code layout:
- Backend: `src\backend\Wordembedded.sln`
  - AppHost: `src\backend\Wordembedded.AppHost\`
  - API: `src\backend\Wordembedded.Api\`
  - Aspire service defaults: `src\backend\Wordembedded.ServiceDefaults\`
- Frontend: not scaffolded yet.

Current commands:
- Build backend: `cd src\backend; dotnet build Wordembedded.sln`
- Run API directly: `cd src\backend; dotnet run --project Wordembedded.Api\Wordembedded.Api.csproj --launch-profile https`
- Run Aspire AppHost: `cd src\backend; dotnet run --project Wordembedded.AppHost\Wordembedded.AppHost.csproj`
- API health check: `Invoke-RestMethod http://localhost:5071/api/health`
- Tests/lint: none configured yet.

Local setup notes:
- The API listens on `https://localhost:7071` and `http://localhost:5071` in Development.
- CORS currently allows the future frontend origin `http://localhost:3000`.
- Graph/SPE minimal API endpoints live in `src\backend\Wordembedded.Api\Endpoints\` and are mounted under `/api`.
- Webhook subscription renewal is in `src\backend\Wordembedded.Api\Services\WebhookRenewalService.cs`; set `Webhooks:NotificationUrl` to a publicly reachable dev tunnel URL before creating Graph subscriptions locally.
- Fill Entra ID, backend API audience, SharePoint Embedded container type, and webhook client-state values with user secrets or environment variables; do not commit real tenant IDs or secrets to `appsettings*.json`.

## What the project is

A React app for document collaboration backed by **SharePoint Embedded (SPE)**. SPE is a headless, API-only file partition inside a customer's Microsoft 365 tenant. The React app provides the entire UX; Microsoft 365 owns storage, identity, permissions, and Office co-authoring.

Read `docs/project/PRD.md` end-to-end before making non-trivial decisions — many architecture details below have important caveats spelled out there (especially §4 Feasibility, §9 Auth, §10 Tenant Setup).

## The single most important architectural constraint

**Word/Excel/PowerPoint editing cannot be hosted inside an iframe on your domain.** Office for the web sets a `frame-ancestors` CSP that blocks third-party domains. The supported pattern is:

- **Read-only preview** → inline iframe using `POST /drives/{driveId}/items/{itemId}/preview` and the returned `getUrl` (append `&nb=true` to hide the banner). Preview URLs are short-lived; treat them as one-shot.
- **Editing** → open `driveItem.webUrl` in a new tab or popup via `window.open()`. Co-authoring, AutoSave, presence, comments, and @mentions all happen in Microsoft's window, not yours.

Do not propose, sketch, or build flows that put an editable Word/Excel/PowerPoint surface inside the React app's own iframe. The beta `allowEdit` parameter on `driveItem: preview` is not promoted to v1.0 and must not be a build dependency.

## Intended stack (per PRD §8)

| Layer | Choice |
|---|---|
| Frontend | Next.js 16 (App Router) + React 19, TypeScript |
| Frontend auth | MSAL.js (`@azure/msal-browser`, `@azure/msal-react`) |
| Frontend Graph client | `@microsoft/microsoft-graph-client` |
| Data fetching | TanStack Query |
| UI | shadcn/ui or Fluent UI v9 |
| Backend | ASP.NET Core 9 minimal API |
| Backend orchestration | .NET Aspire (OpenTelemetry built-in) |
| Backend Graph | Microsoft Graph SDK for .NET v5 |
| Backend auth | `Microsoft.Identity.Web` (handles OBO) |
| Hosting | Azure Static Web Apps / App Service (frontend); Azure Container Apps (backend) |
| Secrets | Azure Key Vault + Managed Identity |

If you scaffold a different stack, first confirm with the user and update the PRD.

## Auth flow (don't get this wrong)

Three distinct token flows — they are easy to conflate:

1. **Frontend → Backend API**: MSAL.js auth code + PKCE. Frontend gets a token whose audience is *your backend API*, not Graph.
2. **Backend → Graph (per-user)**: On-Behalf-Of (OBO) exchange. `Microsoft.Identity.Web` does this in a few lines. Required scope: `FileStorageContainer.Selected`.
3. **Backend → Graph (app-only)**: Client credentials. Used for container type registration, webhook subscription management, and admin tasks — not for user data access.

The Entra app needs a **hidden** SharePoint Online permission (`Container.Selected`, resourceAppId `00000003-0000-0ff1-ce00-000000000000`) that does not appear in the portal UI — it must be added via manifest and consented via the explicit admin-consent URL. See PRD §9.

Permission layering: app registration scopes → container type registration (per consuming tenant) → container membership → driveItem permissions. A user only sees containers they are a member of, regardless of what the app is allowed to do.

## Operational gotchas to keep in mind

- **Graph webhook subscriptions expire every 3 days max.** Any webhook implementation needs a renewal job (cron / timer-triggered function) that refreshes before expiry.
- **Preview URLs are short-lived.** Refresh on iframe error; do not cache long-term.
- **Large file uploads** require the Graph upload-session pattern, not a single `PUT`.
- **Billing model is immutable at container type creation** — Standard (you pay) vs. Passthrough (consuming tenant pays). Choosing requires a product decision, not just a dev decision.
- **Multi-tenant onboarding** has a real flow: admin consent → `containerTypeRegistration` API call on the consuming tenant → (Passthrough only) customer admin sets up PAYG → app creates initial containers. Don't skip steps.
- **@mentions** require the target user to have an M365 license and be in the same tenant. Cross-tenant @mentions are not supported.

## Doc conventions

- Product/architecture docs live under `docs/`. `docs/project/` for product-facing specs (PRD, scope), `docs/architecture/` for technical design docs.
- Treat `PRD.md` as a living document — when you make a decision that contradicts or extends it, update the PRD in the same change.
- Reference links in PRD §14 are curated and worth preserving; if you add references, group them under the same headings.

## When code lands

Frontend scaffold lives in `src/frontend`.

Frontend commands:
- Install/apply dependencies: `cd src\frontend; npm install`
- Run locally: `cd src\frontend; npm run dev`
- Build: `cd src\frontend; npm run build`
- Lint: `cd src\frontend; npm run lint`
- Tests: none configured yet.

Frontend local setup:
- Copy `src\frontend\.env.local.example` to `src\frontend\.env.local` and set `NEXT_PUBLIC_AAD_CLIENT_ID`, `NEXT_PUBLIC_AAD_TENANT_ID`, `NEXT_PUBLIC_API_BASE_URL`, and `NEXT_PUBLIC_API_SCOPE`.
- MSAL redirect URI expected locally: `http://localhost:3000/auth/callback`.
- Frontend requests a token for the backend API `access_as_user` scope; backend OBO/Graph integration is not scaffolded yet.

When more code lands, keep this section updated with backend paths, Aspire wiring, single-test commands, and any webhook/dev-tunnel setup.
