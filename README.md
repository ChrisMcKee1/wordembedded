# Wordembedded

A React + .NET reference app for collaborative document workflows on **SharePoint Embedded
(SPE)**. It's a working demonstration of every officially supported SPE capability — container
management, preview, sharing, versioning, webhooks, recycle bin, thumbnails — plus an ETL pattern
that pulls files from **regular SharePoint document libraries**, lets a team work on them inside an
SPE container, and publishes the result back to SharePoint with full version history.

Everything in this README is grounded in the Microsoft Learn docs cited inline. Where the
documentation contradicts third-party tutorials or older blog posts, the docs win.

---

## 1. What SharePoint Embedded actually is

> "SharePoint Embedded is a cloud-based file and document management system suitable for use in any
> application. SharePoint Embedded is a new API-only solution that enables app developers to
> harness the power of the Microsoft 365 file and document storage platform for any app."
> — [SPE Overview](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/overview)

The shape of it:

- When an app uses SPE inside a Microsoft 365 tenant, SPE creates **a separate storage partition**
  inside that tenant. It has **no built-in UI** — your app is the entire user experience. Content
  is reached only through Microsoft Graph. ([Overview](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/overview))
- Within that partition, your app creates one or more **File Storage Containers**. A container is
  the basic unit of storage and the security/compliance boundary. Think of it as "an API-only
  document library that your app fully owns." ([Overview](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/overview))
- Every container belongs to exactly one **Container Type** — a tenant-level resource that defines
  the relationship, access privileges, billing accountability, and default behaviors for all
  containers of that type. There is a strict **1:1 relationship between an owning application and a
  container type**. ([Container types](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/getting-started/containertypes), [App architecture](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/app-architecture))
- The files inside a container inherit all the M365 features: Office co-authoring, AutoSave,
  version history, Microsoft Search, Purview compliance, sensitivity labels, retention, DLP, and
  Copilot grounding (via Microsoft Foundry SPE knowledge source). ([Overview](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/overview), [Security and compliance](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/compliance/security-and-compliance))

### Owning tenant vs consuming tenant

Two concepts the docs use throughout ([App architecture](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/app-architecture)):

- **Owning tenant** — the Entra ID tenant where the container type is *created*. Often the tenant
  where the app itself is registered.
- **Consuming tenant** — any Entra ID tenant where the container type is *used*. Only consuming
  tenants hold actual container instances. The same tenant can be both owning and consuming.

For Wordembedded both roles are the same tenant (`<your-tenant>.onmicrosoft.com` / id
`<your-tenant-guid>`). The single-tenant LOB scenario.

### What's in your "consumer view" of SharePoint

Containers do not show up in regular SharePoint search, the SharePoint Admin Center's
**Active sites**, OneDrive MRU, etc. **By design**:

- Container types ship with `DiscoverabilityDisabled = $true` so files in containers do not appear
  in cross-surface discovery. The developer admin can flip this via
  `Set-SPOContainerTypeConfiguration -DiscoverabilityDisabled $false`. ([Developer admin](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/administration/developer-admin/dev-admin))
- Microsoft Search treats SPE content as "hidden content" — to surface it tenant-wide a search
  request must set `sharePointOneDriveOptions.includeHiddenContent = true`. ([Search SharePoint Embedded content](https://learn.microsoft.com/en-us/microsoftsearch/sharepoint-embedded-content))
- Admin-only views of SPE containers exist in the **SharePoint Admin Center → SharePoint Embedded**
  pages (Active / Deleted / Archived containers) for users with the **SharePoint Embedded
  Administrator** role (Global Admin inherits this). ([Consuming tenant admin](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/administration/consuming-tenant-admin/cta))

### Container vs regular SharePoint document library — the boundary

| | SharePoint Embedded container | Regular SharePoint document library |
|---|---|---|
| Owner | Your app (the owning application of its container type) | A site collection |
| Access | Container membership + driveItem permissions | Site permissions + driveItem permissions |
| Browseable in SharePoint UI | No (API-only) | Yes |
| Browseable in Microsoft Graph | Yes, but only by apps registered to its container type | Yes, with `Sites.Read.All` / `Files.Read.All` |
| Search via `Microsoft Search` | Hidden by default; needs `includeHiddenContent: true` | Visible by default |
| Billing | Container-type-level PAYG (Standard or Passthrough) | Counts against tenant's pooled SharePoint storage |

**A SharePoint Embedded application CANNOT directly access files in a regular SharePoint document
library through the SPE APIs.** SPE's authorization model is strictly `app → container type →
containers`. ([App architecture](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/app-architecture))

You CAN access regular SharePoint libraries from the same Entra ID app by *also* requesting
**`Sites.Read.All`** or **`Sites.ReadWrite.All`** Graph permissions and calling the standard
`/sites/.../drives` Graph endpoints. That's a separate surface — that's how Wordembedded
implements the ETL pattern (see §6).

---

## 2. Capabilities and hard limitations

Read this section before you build anything new on SPE.

### What's fully supported

| Capability | Mechanism |
|---|---|
| Custom React UX on your domain | App-controlled, talks to Graph |
| File storage inside the customer's M365 tenant | SPE container, isolated partition |
| **Read-only inline preview in an iframe** | [`POST /drives/{driveId}/items/{itemId}/preview`](https://learn.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-1.0) returns embeddable `getUrl` |
| **Edit Word/Excel/PowerPoint** | Launch `driveItem.webUrl` in a popup or new tab → Office for the web |
| Real-time co-authoring + AutoSave | Native in Office for the web |
| Version history | Automatic; configurable retention via container type settings |
| Comments + @mentions | Native in Office. @mentions require the target user to have an M365 license and be in the same tenant. |
| Sharing (link + invite) | [`createLink`](https://learn.microsoft.com/en-us/graph/api/driveitem-createlink) + [`invite`](https://learn.microsoft.com/en-us/graph/api/driveitem-invite) |
| Permissions model | Native M365 permissions on the driveItem |
| Microsoft Search across container content | Microsoft Search API + `Files.Read.All` delegated permission |
| Compliance / DLP / sensitivity labels | Native via Purview ([Security and compliance](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/compliance/security-and-compliance)) |
| Copilot grounding | SharePoint Embedded knowledge source in [Microsoft Foundry](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/declarative-agent/sharepoint-embedded-knowledge-source) |
| Cross-drive copy (e.g., SPE ↔ regular SharePoint) | [`POST /drives/{srcDrive}/items/{srcItem}/copy`](https://learn.microsoft.com/en-us/graph/api/driveitem-copy) with `parentReference` targeting the destination drive |
| Recycle bin (list, restore, purge) | [`/storage/fileStorage/containers/{id}/recycleBin`](https://learn.microsoft.com/en-us/graph/api/recyclebin-list-items?view=graph-rest-1.0) |
| Custom properties on containers | [`/customProperties`](https://learn.microsoft.com/en-us/graph/api/filestoragecontainer-post-customproperty?view=graph-rest-1.0) (Wordembedded uses this to store per-container publish targets) |

### Editing flows — the truth, two paths

There are two officially supported ways to edit Word / Excel / PowerPoint stored in an SPE
container, and Wordembedded implements both.

**Path A — Inline editing in our iframe via the beta `driveItem: preview` API.**
The Microsoft Graph **beta** endpoint accepts three parameters the v1.0 endpoint does not
([driveItem: preview (beta)](https://learn.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-beta)):

| Parameter | Effect |
|---|---|
| `viewer: "office"` | Render with Office for the web (instead of OneDrive previewer) |
| `chromeless: false` | Show Office's toolbars and editing chrome |
| `allowEdit: true` | "The file can be edited from the embedded UI." |

When all three are set, the returned `getUrl` is an embeddable URL that **does** open a fully
editable Office for the web session inside our iframe — co-authoring, AutoSave, comments,
@mentions, version history all keep working. Wordembedded exposes this behind an **"Edit inline
(beta)"** Fluent `Switch` on the preview page (`src/frontend/src/app/files/[containerId]/preview/[itemId]/page.tsx`).
When the switch is on, the backend (`src/backend/Wordembedded.Api/Endpoints/PreviewEndpoints.cs`)
routes the preview call to `https://graph.microsoft.com/beta/...` and forwards those flags. When
the switch is off we use `v1.0` for a clean read-only preview with `nb=true` to hide the banner.

**Caveat the docs are explicit about.** Each of `viewer`, `chromeless`, `allowEdit` is marked in the
docs as:
> "This parameter is deprecated and will not be made available on the v1.0 endpoint."

It works today, the official samples and Microsoft demos use it, and there's no announced
timeline for removal — but Microsoft has been clear they won't promote it to v1.0. Treat inline
edit as a great UX option that you should keep behind a feature flag, with a graceful fallback to
Path B.

**Path B — Edit in Office for the web (new tab / popup) via `driveItem.webUrl`.**
This is the path Microsoft documents as the long-term-supported pattern. Office for the web
serves the document on `*.officeapps.live.com` / `*.sharepoint.com` and sets a
`Content-Security-Policy: frame-ancestors` that **blocks third-party iframe embedding**
([Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/5669186/is-there-a-supported-way-to-embed-editable-sharepo)).
So for the popup pattern we never iframe; we call `window.open(driveItem.webUrl, "_blank", "noopener")`.
This is the "Edit in Word / Excel / PowerPoint" button in our sidebar — it works even for users
whose browser blocks the beta iframe (some browser CSP configurations, some tenant-level policies
that lock down where Office for the web can render), and it's the safest bet for editing flows
that need to survive Microsoft eventually removing the beta `allowEdit` flag.

Both paths produce the same writes — the document, version history, presence, and comments are all
on the same Office service. The only difference is whether the editor renders inside our `<iframe>`
or in a separate browser tab.

### What's NOT supported, and the workaround

| Wanted | Reality (with citation) | Workaround |
|---|---|---|
| **Editable Office iframe with only the v1.0 API** | The v1.0 `driveItem: preview` endpoint accepts only `page` and `zoom`. It always returns a read-only viewer. ([driveItem: preview (v1.0)](https://learn.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-1.0)) | Use the **beta** `driveItem: preview` with `allowEdit + viewer:"office" + chromeless:false` (we do this behind the Edit-inline switch), and/or fall back to the new-tab `webUrl` popup. |
| Long-term-stable in-iframe editing without flags | Beta `allowEdit`/`viewer`/`chromeless` are marked **deprecated** in the docs and "will not be made available on the v1.0 endpoint." ([beta preview](https://learn.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-beta)) | Keep the inline-edit toggle behind a feature flag so it can be flipped off if Microsoft removes it; the new-tab `webUrl` flow is the production-safe fallback. |
| Native sticky-note / annotation tools inside a PDF in our app | The March 2026 SPE PDF viewer is read-only with respect to authoring — it can *render* existing sticky notes, comments, and add **in-file search + print**. It does NOT provide an "add sticky note" UI. ([What's new — March 2026](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/whats-new)) | Open the PDF in Microsoft Edge's native annotator (round-trips via SharePoint), or use the Adobe Acrobat for SharePoint integration. |
| Embed Word editor on an unauthenticated public page | Editing requires M365 identity | Anonymous **view-only** share links allow embedded view |
| Cross-tenant @mentions | M365 platform restriction | @mentions only work within the consuming tenant |
| Programmatic merge of conflicting Word versions | Office handles internally | Out of scope — rely on Office co-authoring conflict resolution |
| `replace` conflict on cross-drive copy via the body | `@microsoft.graph.conflictBehavior` is a **query parameter**, not a body field. ([driveItem: copy](https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0)) | Wordembedded sets it via the URL query string. |
| Copy a file while keeping its custom metadata | "Metadata isn't retained when a driveItem is copied, including system metadata and custom metadata." ([driveItem: copy](https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0)) | Wordembedded mirrors per-file workflow state on the source side and surfaces a "View published version" link on completion. |
| Combine `includeAllVersionHistory: true` and `name` parameter on copy | Known issue: `includeAllVersionHistory` is **silently ignored** if `name` is also set. ([driveItem: copy](https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0)) | Wordembedded omits `name` from publish bodies; destination keeps the source filename. |
| Webhook subscription that lives indefinitely | Subscription expirations capped at ~3 days for SharePoint resources. ([Graph webhooks](https://learn.microsoft.com/en-us/graph/webhooks)) | Wordembedded runs an `IHostedService` that re-PATCHes subscription expirations every 12 h. |
| Large file upload through a single PUT | Files > 4 MB must use upload session. | Wordembedded creates an upload session on the backend and the frontend PUTs chunks directly to Graph. |

### Trial container type limits

We currently use a trial container type ([Container types](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/getting-started/containertypes)):

- Up to **5 containers** per tenant (active + recycle bin)
- **1 GB per container**
- **30-day** lifetime, then access to the containers is removed
- Restricted to the developer tenant; cannot be deployed to other consuming tenants

Standard container types are billed via Azure (PAYG) — Standard (developer pays) or `directToCustomer` (consuming tenant pays). The model is chosen at container-type creation and is **immutable**.

---

## 3. Architecture

```
┌──────────────────────────────────────────────────────────────────────────┐
│                        User's Browser                                     │
│   ┌────────────────────────────────────────────────────────────────┐     │
│   │  Next.js 16 App Router + React 19 + Fluent UI v9               │     │
│   │  • MSAL.js (PKCE) — sign-in to Entra ID                        │     │
│   │  • TanStack Query — server state                               │     │
│   │  • Inline iframe preview from driveItem:preview getUrl         │     │
│   │  • Edit-in-Office popup → driveItem.webUrl                     │     │
│   └────────────────────────────────────────────────────────────────┘     │
│        │ Bearer (access_as_user)              │ HTTPS                     │
└────────┼──────────────────────────────────────┼───────────────────────────┘
         ▼                                      ▼
 ┌──────────────────────────────────┐    ┌──────────────────────────────┐
 │  ASP.NET Core 9 minimal API      │    │  Microsoft Office for the    │
 │  (Wordembedded.Api)              │◀───│  web — co-authoring, AutoSave│
 │  • Microsoft.Identity.Web (OBO)  │    │  (Microsoft-hosted, popup)   │
 │  • Microsoft.Graph v5 SDK        │    └──────────────────────────────┘
 │  • Cert-auth client credentials  │
 │    for webhook subscriptions     │
 │  • Webhook receiver + 12h        │
 │    auto-renewal IHostedService   │
 └─────────────┬────────────────────┘
               │ Microsoft Graph (delegated OBO + app-only)
               ▼
 ┌──────────────────────────────────────────────────────────────────────┐
 │                Microsoft 365 Tenant (<your-tenant>)                │
 │  ┌─────────────────────────────────────────────────────────────┐    │
 │  │  SharePoint Embedded partition                              │    │
 │  │   Container type: Wordembedded.Documents (trial)            │    │
 │  │   Containers: Workspace, test, …                            │    │
 │  └─────────────────────────────────────────────────────────────┘    │
 │  ┌─────────────────────────────────────────────────────────────┐    │
 │  │  Regular SharePoint sites + document libraries              │    │
 │  │   (e.g. /sites/AIDeveloper/Shared Documents)                │    │
 │  │   reached via Sites.Read.All / Sites.ReadWrite.All          │    │
 │  └─────────────────────────────────────────────────────────────┘    │
 │  ┌─────────────────────────────────────────────────────────────┐    │
 │  │  Microsoft Entra ID  •  Microsoft Purview                   │    │
 │  └─────────────────────────────────────────────────────────────┘    │
 └──────────────────────────────────────────────────────────────────────┘
```

### Source layout (everything under `src/`)

```
src/
├── README.md             ← developer run book (this README is the product overview)
├── infra/                Entra app registration + certificate + container type provisioning
│   ├── outputs.json        Generated config consumed by sync-config.ps1
│   ├── sync-config.ps1     Writes infra IDs into backend appsettings + frontend .env.local
│   └── README.md           Step-by-step manual provisioning fallback
├── backend/              .NET 9 Aspire AppHost + ASP.NET Core 9 minimal API
│   ├── Wordembedded.sln
│   ├── Wordembedded.AppHost/        Aspire orchestrator. Adds the API project + Next.js app
│   ├── Wordembedded.ServiceDefaults/ OpenTelemetry, health checks, service discovery
│   └── Wordembedded.Api/
│       ├── Program.cs              MSAL JWT + OBO + Graph SDK + CORS + Swagger
│       ├── Endpoints/
│       │   ├── ContainerEndpoints.cs           List/get/create SPE containers
│       │   ├── FileEndpoints.cs                CRUD on driveItems, upload-session, multipart upload
│       │   ├── PreviewEndpoints.cs             v1.0 + beta driveItem:preview
│       │   ├── SharingEndpoints.cs             createLink, invite, list/revoke permissions
│       │   ├── VersionEndpoints.cs             List + restore driveItem versions
│       │   ├── SearchEndpoints.cs              Microsoft Search across the container type
│       │   ├── WebhookEndpoints.cs             Change-notification receiver + subscription mgmt
│       │   ├── RecycleBinEndpoints.cs          List/restore/purge SPE recycle bin
│       │   ├── ThumbnailEndpoints.cs           Per-driveItem thumbnails
│       │   ├── SharePointBrowseEndpoints.cs    /sites, /drives, /search across REGULAR SharePoint
│       │   └── WorkflowEndpoints.cs            Workflow state (Draft→InReview→Approved→Published),
│       │                                       import (SP→container) and publish (container→SP)
│       ├── Services/                  AppOnlyGraphClient, GraphSubscriptionStore,
│       │                              WebhookRenewalService (IHostedService, 12h cadence)
│       ├── Middleware/                GraphErrorMappingMiddleware (ODataError → ProblemDetails)
│       └── Contracts/                 Strongly-typed DTOs for the API
└── frontend/             Next.js 16 (App Router) + React 19 + TypeScript
    └── src/
        ├── app/                       Routes (signin, files, files/[containerId]/…, search, …)
        ├── components/                Fluent UI v9 components, grouped by feature
        ├── lib/api/                   Typed Graph/API wrappers
        └── lib/hooks/                 TanStack Query hooks
```

### Aspire dev orchestration

Aspire 13.3 with `Aspire.Hosting.JavaScript` for the Next.js resource. The AppHost
(`Wordembedded.AppHost/Program.cs`) launches:

| Resource | URL |
|---|---|
| `api` (ASP.NET Core) | `https://localhost:7071` + `http://localhost:5071` |
| `frontend` (Next.js dev via `AddNextJsApp`) | `http://localhost:3000` (Aspire proxy) |
| Aspire dashboard | `https://localhost:17089/login?t=...` |

Aspire injects `NEXT_PUBLIC_API_BASE_URL` into the Next.js process from the live API endpoint.
Pinned ports because Entra redirect URIs and the backend CORS allowlist are fixed at
`http://localhost:3000` / `https://localhost:7071`.

The Aspire CLI + Aspire MCP server provide live console logs, structured logs (ILogger), and
distributed traces (OpenTelemetry) for both processes — invaluable for debugging the OBO flow,
Graph calls, and the publish monitor URL.

### Auth flows (three of them)

1. **Frontend → Backend API.** MSAL.js authorization-code + PKCE in the browser. The frontend
   acquires an access token with audience `api://{clientId}/access_as_user`. ([MSAL.js for React](https://learn.microsoft.com/en-us/entra/identity-platform/tutorial-v2-react))
2. **Backend → Microsoft Graph (per user).** `Microsoft.Identity.Web` performs the
   [On-Behalf-Of (OBO) flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow):
   exchanges the user's bearer for a Graph token scoped to whatever Graph permission the endpoint
   needs (`FileStorageContainer.Selected`, `Files.Read.All`, `Sites.Read.All`,
   `Sites.ReadWrite.All`). The certificate from `secrets/wordembedded-dev.pfx` (CurrentUser\My
   store, thumbprint in `outputs.json`) is used as the confidential-client credential.
3. **Backend → Microsoft Graph (app-only).** Client credentials with the same cert. Used **only**
   for Graph **webhook subscription** management (create + 12h renewal). See `AppOnlyGraph.cs` +
   `WebhookRenewalService.cs`.

---

## 4. Prerequisites

### Local toolchain

- .NET 9 SDK (`dotnet --version` ≥ 9.0)
- .NET Aspire CLI 13.3+ (`irm https://aspire.dev/install.ps1 | iex` then `aspire --version`)
- Node 20+ and npm 10+
- Azure CLI signed in to the M365 tenant (`az login`)
- PowerShell 7+
- `PnP.PowerShell`, `Microsoft.Online.SharePoint.PowerShell`, and `MSAL.PS` modules for one-time
  provisioning (the `src/infra/README.md` walks through `Install-Module` for each)
- Docker Desktop (Aspire 13.x uses it for container-based dev resources; not strictly required for
  our current resources, but Aspire's environment doctor expects it)

### Tenant prerequisites

- A Microsoft 365 tenant with SharePoint enabled. ([Container types — Tenant requirements](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/getting-started/containertypes))
- An admin user with **SharePoint Embedded Administrator** or **Global Administrator** role to
  create the container type and grant admin consent. ([Admin role](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/administration/adminrole))
- For Standard / passthrough container types, an Azure subscription for billing. We use **trial**
  so this is not required.

### Entra ID app registration — required Graph delegated scopes

These are the scopes the app actually consents — they correspond to the operations Wordembedded
performs. Lookup IDs via `az ad sp show --id 00000003-0000-0000-c000-000000000000`.

| Scope | Why we need it | Doc |
|---|---|---|
| `openid`, `profile`, `User.Read` | Sign-in | [Permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference) |
| `FileStorageContainer.Selected` | All container/driveItem operations inside our SPE container type | [SPE auth](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth#access-on-behalf-of-a-user) |
| `FileStorageContainerType.Manage.All` | Create the container type (owning-tenant only) | [SPE auth](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth#application-permissions) |
| `FileStorageContainerTypeReg.Selected` | Register the container type on consuming tenants | [SPE auth](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth#application-permissions) |
| `Files.Read.All` | Microsoft Search across SPE + read driveItems on regular SharePoint (source for `import`) | [Permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference#filesreadall) |
| `Sites.Read.All` | Browse SharePoint sites + libraries for the ETL search | [Permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference#sitesreadall) |
| `Sites.ReadWrite.All` | Cross-drive copy *to* a SharePoint document library (publish) | [driveItem: copy permissions](https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0#permissions) |

### Entra ID app registration — required application (app-only) scopes

| Scope | Why | Doc |
|---|---|---|
| `FileStorageContainer.Selected` | Container management + webhook subscription create/renew | [SPE auth — Access without a user](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth#access-without-a-user) |
| `FileStorageContainerTypeReg.Selected` | Container type registration (one-time, during provisioning) | [SPE auth](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth) |

### Container type permissions

These permissions are granted to the app *on the container type* via the **container type
registration** API. The owning application sets them when calling
[`PUT /storage/fileStorage/containerTypeRegistrations/{containerTypeId}`](https://learn.microsoft.com/en-us/graph/api/filestorage-post-containertyperegistrations).
Wordembedded grants itself **`full`** delegated and **`full`** application permissions on its
container type during provisioning.

The full taxonomy ([SPE auth — Container type application permissions](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth#container-type-application-permissions)):
`None`, `ReadContent`, `WriteContent`, `Create`, `Delete`, `Read`, `Write`, `EnumeratePermissions`,
`AddPermissions`, `UpdatePermissions`, `DeletePermissions`, `DeleteOwnPermission`,
`ManagePermissions`, `ManageContent`, `Full`.

### Entra ID app registration — redirect URIs

| Platform | URI | Why |
|---|---|---|
| Single-page application (SPA) | `http://localhost:3000/auth/callback` | MSAL.js auth-code + PKCE in the browser |
| Web | `https://localhost:7071/signin-oidc` | OBO acquisition target; required even though we don't use the interactive web flow |
| Public client | `http://localhost` | PnP PowerShell / MSAL desktop flows during provisioning |
| API exposed | `api://{clientId}` with scope `access_as_user` (User-consentable) | Frontend → backend audience |

### Certificate

Wordembedded uses a self-signed cert (2-year validity) stored in **CurrentUser\My** plus a PFX in
`src/infra/secrets/wordembedded-dev.pfx`. The public key (`.cer`) is uploaded to the app registration
as a `keyCredentials` entry. This is what `Microsoft.Identity.Web` uses to acquire tokens during the
OBO flow and the app-only client-credentials flow.

### Admin consent

Some of the consents — including the SPO hidden `Container.Selected` permission and the
`FileStorageContainerType.Manage.All` scope — require **admin consent**. Wordembedded uses the
direct `oauth2PermissionGrants` Graph API to grant `AllPrincipals` consent at all-tenant scope.

A consent URL also works:
```
https://login.microsoftonline.com/{tenantId}/v2.0/adminconsent
   ?client_id={clientId}
   &scope=https://graph.microsoft.com/.default
```

### Backend appsettings

`src/backend/Wordembedded.Api/appsettings.Development.json` (gitignored, populated by
`sync-config.ps1`):

```jsonc
{
  "AzureAd": {
    "Instance": "https://login.microsoftonline.com/",
    "TenantId": "...",
    "ClientId": "...",
    "Audience": "api://{clientId}",
    "ClientCertificates": [{
      "SourceType": "StoreWithThumbprint",
      "CertificateStorePath": "CurrentUser/My",
      "CertificateThumbprint": "..."
    }]
  },
  "DownstreamApi": {
    "BaseUrl": "https://graph.microsoft.com/v1.0",
    "Scopes": [
      "https://graph.microsoft.com/FileStorageContainer.Selected",
      "https://graph.microsoft.com/FileStorageContainerTypeReg.Selected",
      "https://graph.microsoft.com/Files.Read.All",
      "https://graph.microsoft.com/User.Read"
    ]
  },
  "SharePointEmbedded": { "ContainerTypeId": "..." },
  "Webhooks": {
    "NotificationUrl": "https://localhost:7071/api/webhooks/graph",
    "ClientStateSecret": "..."
  }
}
```

**Common pitfalls** (we hit each of these — caught via the Aspire MCP logs):

- `CertificateStorePath` must be `"CurrentUser/My"` (single string). Older docs show
  `CertificateStoreLocation` + `CertificateStoreName` separately; that schema NREs in current
  `Microsoft.Identity.Web`.
- `DownstreamApi:Scopes` must be **specific delegated scopes**, not `.default` — AADSTS70011
  forbids combining `.default` with resource-specific scopes.
- Don't request scopes the user/admin hasn't consented to; AADSTS65001 surfaces as a 500 on the
  endpoint.

### Frontend `.env.local`

`src/frontend/.env.local` (gitignored, also populated by `sync-config.ps1`):

```
NEXT_PUBLIC_AAD_CLIENT_ID=...
NEXT_PUBLIC_AAD_TENANT_ID=...
NEXT_PUBLIC_API_BASE_URL=https://localhost:7071
NEXT_PUBLIC_API_SCOPE=api://{clientId}/access_as_user
```

---

## 5. Running locally

```powershell
# One-time: provision Entra app, cert, container type, container.
# Follow src/infra/README.md exactly. After it succeeds, src/infra/outputs.json is populated.

pwsh src/infra/sync-config.ps1   # propagates IDs into backend + frontend config

# Day-to-day: start everything via Aspire
cd src/backend/Wordembedded.AppHost
aspire run
```

The Aspire CLI starts the dashboard, the API, and the Next.js dev server. Open the dashboard URL
from the console for live structured logs and traces. The frontend is at <http://localhost:3000>.

---

## 6. The ETL workflow — pulling files from SharePoint into a container and publishing them back

### Why this pattern exists

SPE containers are **isolated by design** — they're a security and compliance boundary owned by the
app. Files in regular SharePoint document libraries are NOT directly accessible through SPE APIs
([App architecture — Access Model](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/app-architecture)).
That's the security feature, not a bug.

But many real-world flows need to:

1. **Pull a document from a regular SharePoint library into the app**, where a defined group of
   reviewers can collaborate in an isolated space.
2. **Publish the approved result back to a regular SharePoint library** (or a different one) where
   the rest of the organisation expects to find it.

Wordembedded implements this as an "import + workflow + publish" pipeline using the standard
Microsoft Graph `driveItem: copy` action both ways, with the same Entra app holding both the
`FileStorageContainer.Selected` permission (for SPE) and `Sites.ReadWrite.All` (for regular
SharePoint).

### Stage 1 — Discover (search regular SharePoint)

The container explorer has a **"Pull files from SharePoint or OneDrive"** panel that calls
`POST /search/query` via Microsoft Graph with `entityTypes: ["driveItem"]`. The search runs
against everything the signed-in user has access to in the tenant.

Backend: `SharePointBrowseEndpoints.cs` → `GET /api/sharepoint/search?q=...`
([Microsoft Search API](https://learn.microsoft.com/en-us/graph/search-concept-files)).

### Stage 2 — Import (copy from SharePoint → SPE container)

Frontend "Import" button calls `POST /api/sharepoint/import` with the source `driveId` + `itemId`
and the target container's `driveId`. The backend issues a Graph cross-drive copy:

```http
POST https://graph.microsoft.com/v1.0/drives/{sourceDriveId}/items/{sourceItemId}/copy
     ?@microsoft.graph.conflictBehavior=rename
Content-Type: application/json

{
  "parentReference": { "driveId": "{targetSpeDriveId}", "id": "root" }
}
```

Doc: [driveItem: copy](https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0).

Notes:

- **`@microsoft.graph.conflictBehavior` is a query parameter**, not a body field. Putting it in the
  body causes Graph to silently default to `fail` and the operation returns `nameAlreadyExists`.
- Import defaults to `rename` so a re-import of the same source file doesn't silently clobber work
  in progress in the container.
- Graph returns `202 Accepted` with a `Location` header pointing to a monitor URL. The actual copy
  is asynchronous. Wordembedded **does not** poll on import (the file simply appears in the
  container shortly after — TanStack invalidation refetches on a 30s window).

### Stage 3 — Work (in the container)

Each driveItem in the container carries a workflow status stored in its **`description`** field
with the prefix `[wf] <Status> {publishedReference}`:

- `Draft` — initial state after upload or import
- `InReview` — under review
- `Approved` — ready to publish
- `Published` — published; the embedded JSON holds the destination `driveId`, `itemId`, `webUrl`,
  `name`, and `publishedAt` timestamp

Why the description field? Container `customProperties` are container-level, not per-driveItem; and
the standard driveItem schema doesn't expose user-defined columns the way a regular SharePoint
listItem does. The description field is the cheapest universally-supported metadata slot on a
driveItem, and we never overwrite user-authored description text — our parser strips the `[wf]`
line and preserves the rest.

Workflow transitions are buttons on the file preview page (`WorkflowControls.tsx`). The Fluent
status badge color-codes the state. Edit-in-Office continues to work normally — co-authoring,
AutoSave, version history all keep functioning while the file is in any state.

### Stage 4 — Configure the publish destination (per container)

Each container has its own publish target — a SharePoint site + document library + optional folder
path — stored as a JSON blob in the container's **customProperties** (the `publishTarget` key,
non-searchable). Doc:
[Custom properties on fileStorageContainer](https://learn.microsoft.com/en-us/graph/api/filestoragecontainer-post-customproperty?view=graph-rest-1.0).

The "Configure" button on the file preview opens a dialog that lists sites the user has access to
(via `GET /api/sharepoint/sites` → `GET /v1.0/sites?search=*`) and the document libraries on the
selected site (`GET /v1.0/sites/{siteId}/drives`).

### Stage 5 — Publish (copy from SPE container → SharePoint)

Pressing **Publish to SharePoint** (on an `Approved` file) or **Republish (overwrite)** (on a
`Published` file) calls `POST /api/drives/{driveId}/items/{itemId}/publish`. The backend:

1. Issues `POST /drives/{speDriveId}/items/{itemId}/copy?@microsoft.graph.conflictBehavior=replace`
   with body `{ parentReference, includeAllVersionHistory: true }`.
   - **`replace`** ([per the docs](https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0)):
     "The preexisting file item is deleted and replaced with the new item when a conflict occurs.
     This option is only supported for file items. The new item has the same name as the old one.
     **The old item's history is deleted.**"
   - **`includeAllVersionHistory: true`**: brings the **source's** full version history along.
     Without this, only the latest version is copied.
   - We deliberately **omit `name`** from the body. The docs warn:
     > "A known issue occurs when the **includeAllVersionHistory** request parameter is ignored if
     > the **name** request parameter is also passed."
2. Polls the monitor URL until terminal status (`completed`, `succeeded`, `failed`, `cancelled`,
   or timeout). Only on success does the source file get marked `Published`.
3. On success, fetches the destination driveItem's `webUrl` + `name` and stores them on the source
   file's description JSON so the preview UI can render a **"View published version"** link.

### What the user needs to understand about the result

**Publishing creates a copy.** The source file in the SPE container and the destination file in the
SharePoint library are **two separate driveItems with two separate URLs**. Edits made in the
container after publishing do NOT automatically propagate. To overwrite the published version, the
user clicks **Republish (overwrite)** — this re-runs the copy with `replace`, which the
[docs confirm](https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0)
"deletes the old item's history" (history of the *destination*) before placing the new copy with
the source's history. The destination's `webUrl` stays the same (same library + filename); the
published-version link in the sidebar continues to work.

**Custom metadata is not preserved by Graph `copy`.** From the docs:

> "Metadata isn't retained when a driveItem is copied, including system metadata and custom
> metadata. An entirely new driveItem is created in the target location instead."

This is why the workflow state lives on the **source** file (which stays in the container) and is
mirrored as a sidebar link rather than as metadata on the destination.

---

## 7. References

### SharePoint Embedded core
- [Overview](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/overview)
- [What's new](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/whats-new) (March 2026 PDF viewer updates, beta `permissions` + `informationBarrier` properties)
- [App architecture](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/app-architecture)
- [Authentication and authorization](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth)
- [Container types](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/getting-started/containertypes)
- [Container type registration](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/getting-started/register-api-documentation)
- [VS Code extension](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/getting-started/spembedded-for-vscode)
- [Limits and calling patterns](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/limits-calling)
- [Security and compliance](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/compliance/security-and-compliance)

### Microsoft Graph SPE APIs
- [`fileStorageContainer` resource](https://learn.microsoft.com/en-us/graph/api/resources/filestoragecontainer)
- [`fileStorageContainerType` resource](https://learn.microsoft.com/en-us/graph/api/resources/filestoragecontainertype)
- [`fileStorageContainerTypeRegistration` resource](https://learn.microsoft.com/en-us/graph/api/resources/filestoragecontainertyperegistration)
- [Container `customProperties`](https://learn.microsoft.com/en-us/graph/api/filestoragecontainer-post-customproperty?view=graph-rest-1.0)
- [Container `recycleBin`](https://learn.microsoft.com/en-us/graph/api/recyclebin-list-items?view=graph-rest-1.0)
- [Container columns (custom metadata schema)](https://learn.microsoft.com/en-us/graph/api/filestoragecontainer-list-columns)

### Microsoft Graph driveItem APIs we use
- [`driveItem`](https://learn.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0)
- [`driveItem: preview` (v1.0)](https://learn.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-1.0)
- [`driveItem: preview` (beta — `allowEdit`, `viewer`, `chromeless`)](https://learn.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-beta)
- [`driveItem: copy`](https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0) (conflictBehavior + includeAllVersionHistory)
- [`driveItem: invite`](https://learn.microsoft.com/en-us/graph/api/driveitem-invite?view=graph-rest-1.0)
- [`driveItem: createLink`](https://learn.microsoft.com/en-us/graph/api/driveitem-createlink?view=graph-rest-1.0)
- [`driveItem: createUploadSession`](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0)
- [Thumbnails](https://learn.microsoft.com/en-us/graph/api/driveitem-list-thumbnails)
- [Versions](https://learn.microsoft.com/en-us/graph/api/driveitem-list-versions)

### Identity, search, and webhooks
- [Microsoft Search API](https://learn.microsoft.com/en-us/graph/search-concept-files)
- [Search SharePoint Embedded content (incl. `includeHiddenContent`)](https://learn.microsoft.com/en-us/microsoftsearch/sharepoint-embedded-content)
- [Graph webhooks (change notifications)](https://learn.microsoft.com/en-us/graph/webhooks) — note the ≤3-day subscription expiration cap
- [Microsoft.Identity.Web — OBO patterns](https://learn.microsoft.com/en-us/entra/msal/dotnet/microsoft-identity-web/)
- [Permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [On-Behalf-Of flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow)
- [Client credentials with certificate](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow)

### Microsoft Q&A confirmation that an editable Word iframe is not supported
- [Q&A — Embedding editable SharePoint Embedded content](https://learn.microsoft.com/en-us/answers/questions/5669186/is-there-a-supported-way-to-embed-editable-sharepo)

### Samples and tutorials
- [Official SPE samples repo](https://github.com/microsoft/SharePoint-Embedded-Samples)
- [Preview tutorial (`getUrl` + `nb=true`; PDF `embed=` params)](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/tutorials/using-file-preview)

### Aspire
- [Aspire docs](https://aspire.dev)
- [Aspire 13.0 — multi-language platform (JavaScript first-class)](https://aspire.dev/whats-new/whats-new-in-aspire-13/)
- [Aspire 13.3 — `AddNextJsApp`, JavaScript publishing](https://aspire.dev/whats-new/whats-new-in-aspire-133/)
- [Deploy JavaScript apps (Next.js, Vite, etc.)](https://aspire.dev/integrations/frameworks/javascript/)
