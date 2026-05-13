# SharePoint Embedded Document Collaboration App
## Product Requirements Document

**Version:** 0.1 (Draft)
**Author:** Chris McKee
**Date:** May 2026
**Status:** For review

---

## 1. Executive Summary

This PRD covers a hypothetical React-based document collaboration app built on **SharePoint Embedded (SPE)**. The app would give users a custom file management UX (browsing, organizing, sharing, AI-assisted workflows) while delegating all document storage, identity, permissions, and Office co-authoring to Microsoft 365.

**The headline verdict:** the architecture is supported and proven. Microsoft uses the same pattern internally for Loop and Designer. You can build a fully custom React app that stores files in a SharePoint Embedded container, lets users open Word/Excel/PowerPoint documents, and supports real-time multi-user co-authoring with AutoSave and version history.

**The one important caveat:** Microsoft does **not** support hosting a fully editable Word iframe inside your own domain. The supported pattern for editing is to launch Office for the web in a new tab or popup. Inline editing inside your React app via iframe is not on the public roadmap. Read-only inline preview via iframe is fully supported and clean.

---

## 2. Background

### What is SharePoint Embedded?

SharePoint Embedded is an API-only, headless file storage platform that lets third-party apps store and manage content inside a customer's Microsoft 365 tenant. Each app gets its own partition (one or more "containers") that's isolated from the regular SharePoint sites and OneDrive content the user normally sees. Files in the container inherit Microsoft 365 features: Office co-authoring, AutoSave, versioning, Purview compliance, sensitivity labels, eDiscovery, and Microsoft Search.

The headless model means SPE has no UI of its own. Your app provides the entire user experience and talks to containers through Microsoft Graph.

### Why use it for this scenario?

The alternative paths each have problems. Building a Word-compatible editor (Slate, ProseMirror, or similar) and re-implementing co-authoring with CRDTs is years of work and won't be a faithful Word experience. Becoming a WOPI host through the Cloud Storage Partner Program is heavyweight and aimed at storage providers. Using regular SharePoint document libraries pulls users into the SharePoint UX and doesn't give you the per-app isolation. SPE is the path Microsoft actually built for this use case.

---

## 3. Goals & Non-Goals

### Goals

- React app with custom UX for browsing, organizing, and sharing documents stored in a SharePoint Embedded container
- Inline read-only preview of Office docs and PDFs inside the React app
- Edit and co-author Word/Excel/PowerPoint documents (via Office for the web, launched from the React app)
- Real-time multi-user co-authoring with presence indicators
- AutoSave, version history, comments, and @mentions
- Native Microsoft 365 permissions and sharing (internal and external)
- Compliance with Purview, sensitivity labels, and tenant DLP policies
- Multi-tenant support if this becomes an ISV product

### Non-Goals

- Inline editable Word iframe inside the React app's domain (not supported by Microsoft)
- Custom merge / conflict resolution logic (Office handles this)
- Custom permissions model layered over the SPE permissions
- Replacing Office for the web with a custom editor
- Mobile native apps in phase 1 (web responsive only)

---

## 4. Feasibility Analysis

### What's fully supported

| Capability | Mechanism |
|---|---|
| Custom React UX hosted on your domain | App-controlled, talks to Graph |
| File storage in customer's M365 tenant | SPE container, isolated partition |
| Read-only inline preview in iframe | `POST /drives/{driveId}/items/{itemId}/preview` returns embeddable `getUrl` |
| Open Word for editing | Launch `driveItem.webUrl` in popup or new tab |
| Real-time co-authoring | Native in Office for the web on the document |
| AutoSave | Automatic on Office docs in SPE containers |
| Version history | Automatic, configurable retention |
| Comments and @mentions | Native in Office, @mentions require M365 license on target user |
| Sharing internal + external | Graph `createLink` API, respects tenant sharing settings |
| Permissions model | Native M365 permissions on driveItem |
| Search across container content | Microsoft Search, `Files.Read.All` delegated permission |
| Compliance, DLP, sensitivity labels | Native via Purview on the consuming tenant |
| Copilot grounding on container content | Via SharePoint Embedded knowledge source in Microsoft Foundry |

### What's not supported (and the workaround)

| Wanted | Reality | Workaround |
|---|---|---|
| Editable Word iframe inside React app domain | Office for the web sets frame-ancestors CSP that blocks third-party domains | Launch edit in popup or new tab. Users return to React app when done. AutoSave keeps state synced. |
| Embed Word editor on a public, unauthenticated page | Editing requires M365 identity | Anonymous share links allow view-only embed |
| Cross-tenant @mentions | M365 platform restriction | @mentions limited to users in the consuming tenant |
| Programmatic merge of conflicting versions | Office handles this internally | Out of scope; rely on Office co-authoring conflict resolution |

### Beta-only capabilities to monitor

The Graph beta endpoint for `driveItem: preview` accepts an `allowEdit: true` parameter that documents as "the file can be edited from the embedded UI." In practice this has not been promoted to v1.0 and Microsoft's public guidance still directs developers to launch edit in Office for the web. Treat this as a future possibility, not a build dependency.

Reference: https://learn.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-beta

---

## 5. User Stories

### Reader
- As a reader, I can browse a folder of documents inside the React app
- As a reader, I can click a Word doc and see a read-only preview inline without leaving the app
- As a reader, I can search across all documents I have access to
- As a reader, I can request edit access from a document owner

### Editor
- As an editor, I can click "Edit" on a Word doc and it opens in Office for the web in a new tab
- As an editor, I can co-author the document in real time with other editors and see their presence
- As an editor, I can leave comments and @mention teammates
- As an editor, my changes save automatically and version history is kept
- As an editor, I can return to the React app and see updated metadata (last modified, last editor)

### Owner / Admin
- As an owner, I can share a document with specific people, my organization, or anyone with the link
- As an owner, I can revoke access at any time
- As an owner, I can see who has accessed the document and when
- As an admin, I can apply sensitivity labels and DLP rules through Purview

### Tenant Admin (Consumer)
- As a tenant admin, I can review what data my org's SPE app has access to
- As a tenant admin, I can monitor storage and transaction costs in Azure Cost Management
- As a tenant admin, I can revoke the app's container type registration

---

## 6. Functional Requirements

### F1. File Management
- List containers the signed-in user has access to
- List files and folders within a container with paging
- Create, rename, move, copy, and delete files and folders
- Upload files (small files via single PUT, large files via upload session)
- Download files
- Search within a container

### F2. Document Preview
- Inline iframe preview for Office docs, PDFs, images, and supported file types
- Preview banner can be hidden via `nb=true` URL parameter
- Preview URLs are short-lived; refresh on expiration
- Preview honors the calling user's permissions

### F3. Document Editing
- "Edit in Word" button launches `driveItem.webUrl` in a new tab or popup
- Office for the web handles co-authoring, AutoSave, comments
- React app detects when the editing window closes and refreshes metadata
- Breadcrumb properties on the container surface the app name and a deep link back to the React app from inside Office's UI

### F4. Sharing & Permissions
- Generate sharing links: anyone, organization, specific people, existing access
- Set link permissions: view or edit
- Set expiration and password on links where tenant policy allows
- Send invitations with optional email message
- List, modify, and revoke existing permissions on a file

### F5. Versioning
- Automatic versioning on Office files
- List versions, restore a version, delete a version
- Configurable major version limit per container type

### F6. Notifications
- Subscribe to change notifications on a drive or container via Graph webhooks
- React app gets push notifications when files change so the UI can refresh

### F7. Search
- Microsoft Search API across container content
- Requires `Files.Read.All` delegated permission in addition to `FileStorageContainer.Selected`

### F8. Compliance & Governance
- Sensitivity labels applied via Purview are honored
- DLP policies on the consuming tenant apply
- eDiscovery and audit logs work out of the box
- Documents in archived containers return appropriate error states

### F9. Authentication
- MSAL-based sign-in on the React frontend
- Backend uses On-Behalf-Of (OBO) flow to call Graph on behalf of the user
- Multi-tenant Entra app registration for ISV scenarios
- Admin consent flow for new consuming tenants

### F10. Multi-Tenancy & Onboarding
- New customer admin signs in, app detects no container type registration
- App walks admin through admin consent
- App calls `containerTypeRegistration` API to register on the new tenant
- For pass-through billing, admin sets up PAYG billing in their M365 Admin Center
- App creates initial container(s) for the customer

---

## 7. Architecture Overview

```
┌────────────────────────────────────────────────────────────────────┐
│                       User's Browser                                │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │  React App (your domain)                                     │  │
│  │  - File browser, search, share dialogs, custom UX            │  │
│  │  - MSAL.js for sign-in                                       │  │
│  │  - Inline iframe for read-only preview                       │  │
│  │  - Launches popup → Office for the web for editing           │  │
│  └──────────────────────────────────────────────────────────────┘  │
│         │                                          │                │
│         │ HTTPS                                    │ HTTPS           │
│         ▼                                          ▼                │
└─────────┼──────────────────────────────────────────┼────────────────┘
          │                                          │
          ▼                                          ▼
┌──────────────────────────┐         ┌──────────────────────────────┐
│  Your Backend API        │         │  Office for the Web          │
│  (ASP.NET Core / Azure   │         │  (Microsoft-hosted)          │
│  Functions / Container   │◄────────┤  - Word, Excel, PowerPoint    │
│  Apps)                   │         │  - Co-authoring, AutoSave    │
│  - OBO token exchange    │         │  - Comments, @mentions       │
│  - Graph SDK calls       │         └──────────────────────────────┘
│  - Webhook receiver      │                       │
└──────────┬───────────────┘                       │
           │                                       │
           │ Microsoft Graph                       │ Reads/writes
           ▼                                       ▼
┌──────────────────────────────────────────────────────────────────┐
│                      Microsoft 365 Tenant (Consumer)              │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │  SharePoint Embedded Partition                             │  │
│  │  ┌──────────────────────────────────────────────────────┐  │  │
│  │  │  Container Type: YourApp.Documents                   │  │  │
│  │  │  ┌────────────────┐  ┌────────────────┐              │  │  │
│  │  │  │ Container A    │  │ Container B    │  ...         │  │  │
│  │  │  │ - file1.docx   │  │ - file2.xlsx   │              │  │  │
│  │  │  │ - file3.pdf    │  │ - file4.pptx   │              │  │  │
│  │  │  └────────────────┘  └────────────────┘              │  │  │
│  │  └──────────────────────────────────────────────────────┘  │  │
│  └────────────────────────────────────────────────────────────┘  │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │  Microsoft Entra ID (identity, consent, OAuth)             │  │
│  └────────────────────────────────────────────────────────────┘  │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │  Purview (DLP, sensitivity labels, eDiscovery)             │  │
│  └────────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────────┘
```

### Key flows

**Sign-in:** MSAL.js handles interactive sign-in on the frontend, gets an access token for your backend API.

**Graph call:** Frontend calls your backend with its bearer token. Backend uses OBO flow to exchange that token for a Graph token scoped to `FileStorageContainer.Selected`. Backend calls Graph.

**Preview:** Backend calls `POST /drives/{driveId}/items/{itemId}/preview`, returns the `getUrl` to the frontend, frontend drops it into an `<iframe>`.

**Edit:** Frontend opens `driveItem.webUrl` in `window.open()`. User edits in Office for the web. When they close it, frontend refreshes metadata via Graph.

**Co-authoring:** Multiple users open the same `webUrl` simultaneously. Office for the web handles all real-time sync, presence, and conflict resolution. Your React app doesn't see this traffic.

---

## 8. Technical Stack

### Recommended (matches your existing patterns)

| Layer | Choice | Rationale |
|---|---|---|
| Frontend framework | Next.js 16 (App Router) + React 19 | Matches Mosaic Wealth, you already know it |
| Frontend auth | MSAL.js (`@azure/msal-browser`, `@azure/msal-react`) | Standard for Entra ID in React |
| Frontend Graph client | `@microsoft/microsoft-graph-client` | Lightweight, well-maintained |
| State / data fetching | TanStack Query | Caching, optimistic updates |
| UI components | shadcn/ui or Fluent UI v9 | Fluent for Microsoft-native feel, shadcn if you want consistency with Mosaic Wealth |
| Backend | ASP.NET Core 9 minimal API | Your default; OBO flow is well-supported |
| Backend orchestration | .NET Aspire | Matches Mosaic Wealth, gives you OpenTelemetry out of the box |
| Backend Graph client | Microsoft Graph SDK for .NET v5 | First-party |
| Backend auth | Microsoft.Identity.Web | Handles OBO transparently |
| Hosting (frontend) | Azure Static Web Apps or App Service | SWA if SPA, App Service for SSR |
| Hosting (backend) | Azure Container Apps | Aligns with your existing infra |
| Webhook receiver | Azure Functions or Container App | Simple HTTP endpoint with validation |
| Observability | OpenTelemetry → Aspire dashboard locally, App Insights in cloud | Matches your existing setup |
| Secrets | Azure Key Vault + Managed Identity | Standard |

### Alternative (lighter weight)

If you want a Microsoft-canonical starting point and prefer to keep things minimal: clone the **`boilerplate-typescript-react`** sample from `microsoft/SharePoint-Embedded-Samples`, which gives you React + TypeScript + Azure Functions OBO proxy. Then layer Next.js or your preferred frontend over it once you understand the auth dance.

### Cross-stack considerations

- The Word/Excel/PowerPoint editing happens entirely in Microsoft's domain, so your bundle stays small. You don't ship a Word editor.
- Webhook subscriptions on Graph expire every 3 days max; you need a renewal job.
- Preview URLs from `driveItem: preview` are short-lived. Treat them as one-shot.
- Large file uploads need the upload session pattern.

---

## 9. Authentication & Authorization

### Entra ID App Registration (Owning Tenant)

Create a multi-tenant Entra app in your tenant. Required Microsoft Graph permissions for the **owning tenant**:

| Permission | Type | Purpose |
|---|---|---|
| `FileStorageContainerType.Manage.All` | Application | Create and manage container types |
| `FileStorageContainerTypeReg.Selected` | Application + Delegated | Register container type on consuming tenants |
| `FileStorageContainer.Selected` | Application + Delegated | Access containers and contents |
| `User.Read` | Delegated | Sign-in basics |

For the **consuming tenant**, drop `FileStorageContainerType.Manage.All` from the manifest. Tenants don't need to grant management permissions to your app.

### Hidden Permissions

You also need the SharePoint Online resource permission **`Container.Selected`** (resourceAppId `00000003-0000-0ff1-ce00-000000000000`). This is a hidden permission and won't appear in the Entra UI. You must edit the manifest directly and grant admin consent via the explicit consent URL:

```
https://login.microsoftonline.com/{tenantId}/v2.0/adminconsent?client_id={clientId}&scope=https://graph.microsoft.com/.default
```

### Token Flows

- **Frontend (React) → Backend API:** MSAL.js authorization code flow with PKCE. Frontend gets a token for your backend API audience.
- **Backend → Microsoft Graph:** OBO (On-Behalf-Of) flow. Backend exchanges the user's token for a Graph token. `Microsoft.Identity.Web` handles this in three lines.
- **Backend → Graph (app-only operations):** Client credentials flow for things like container type registration, webhook subscriptions, or admin tasks.

### Permission Hierarchy

Worth understanding the layered model:

1. **App registration permissions** (Entra): broad scopes your app can request
2. **Container type permissions** (set at container type registration time on each consuming tenant): scopes per container type
3. **Container permissions** (per-container): owners, writers, readers
4. **DriveItem permissions** (per-file): standard M365 sharing model

A user can only access a container they're a member of, even if the app has broad permissions. This is a feature, not a bug.

---

## 10. Tenant Setup Requirements

### Owning Tenant (Your Tenant)

| Step | What | How |
|---|---|---|
| 1 | Enable SharePoint Embedded on tenant | SharePoint Admin Center → Settings |
| 2 | Create Entra app registration | Entra admin center |
| 3 | Add required Graph + SharePoint permissions | App manifest |
| 4 | Grant admin consent (including hidden permissions) | Admin consent URL |
| 5 | Create container type | `New-SPOContainerType` PowerShell or VS Code extension |
| 6 | Choose billing model: Standard or Passthrough | At container type creation; cannot be changed later |
| 7 | Set up Azure billing profile (Standard billing) | Azure Cost Management + Billing |
| 8 | Register container type on owning tenant | `Register-SPOContainerType` |

### Consuming Tenant (Each Customer Tenant)

| Step | What | Who |
|---|---|---|
| 1 | Enable SharePoint Embedded on tenant | Customer M365 admin |
| 2 | Grant admin consent to your app | Customer M365 admin via your onboarding flow |
| 3 | Register your container type on their tenant | Your app calls `containerTypeRegistration` API |
| 4 | Set up PAYG billing (Passthrough billing only) | Customer in their M365 Admin Center |
| 5 | Create initial containers | Your app via Graph |

### What to "Turn On" Checklist

The literal list of things that need to be enabled or configured:

- [ ] SharePoint Embedded enabled in owning tenant
- [ ] SharePoint Embedded enabled in each consuming tenant
- [ ] Multi-tenant Entra app registration
- [ ] `FileStorageContainer.Selected` (delegated + app)
- [ ] `FileStorageContainerType.Manage.All` (app, owning tenant only)
- [ ] `FileStorageContainerTypeReg.Selected` (delegated + app)
- [ ] `Container.Selected` SharePoint Online hidden permission
- [ ] `User.Read` (delegated)
- [ ] `Files.Read.All` (delegated, only if using Microsoft Search)
- [ ] Admin consent granted via consent URL
- [ ] Container type created and configured
- [ ] Container type registered on each consuming tenant
- [ ] Azure subscription + resource group for billing
- [ ] Container type billing profile set (Standard) or PAYG configured on consuming tenant (Passthrough)
- [ ] Container type `urlTemplate` set for breadcrumb deep links back to your app
- [ ] Optional: configure sharing capability, item versioning, max storage per container

---

## 11. Billing & Cost Model

SPE uses an Azure pay-as-you-go billing model with four meters:

| Meter | Measures |
|---|---|
| Storage | GB stored per month, active and archived |
| Transactions | API calls to read/modify content |
| Egress | Data leaving the SPE platform |
| Active users | Per-user activity in some scenarios |

### Two Billing Models (chosen at container type creation, immutable)

- **Standard billing:** All charges flow to the developer (owning) tenant's Azure subscription. Simpler for the customer; you eat the bill and pass costs in your pricing.
- **Passthrough billing:** Charges flow to each consuming tenant's Azure subscription. Customer pays Microsoft directly. Cleaner for enterprise sales but requires customer M365 admin to set up PAYG in their tenant.

### Cost Planning

Use the **SharePoint cost calculator** to model expected usage before launch. Storage is usually the dominant line item for document-heavy apps. Transactions can spike during bulk operations (migration, batch processing, frequent preview generation).

Through June 2026, tenants with PAYG billing set up get included capacity per month at no cost for selected services. Verify whether this covers your scenario.

### Trial Container Types

For dev/POC work, trial container types are free for 30 days and capped at 100MB. Useful for proving out the integration before standing up production billing.

---

## 12. Risks & Open Questions

### Risks

| Risk | Severity | Mitigation |
|---|---|---|
| Microsoft never delivers inline editable iframe support | Medium | Design around popup/new-tab edit pattern from day one. Don't promise inline edit to users. |
| Preview URL expiration causes broken iframes after long idle | Low | Refresh on iframe error, regenerate preview on demand |
| Webhook subscription expiry (3 days) causes missed updates | Medium | Cron job to renew all active subscriptions every 24h |
| Customer tenant doesn't enable PAYG (Passthrough model) | Medium | Onboarding flow checks billing setup, blocks container creation with a clear error |
| Cross-tenant @mentions don't work | Low | Document the limitation; if needed, fall back to email invitations |
| @mentions require M365 license on target user | Low | Detect license, show alternate flow for unlicensed users |
| Large file uploads need upload session pattern | Low | Standard Graph pattern, well documented |
| Beta APIs you depend on get removed | Medium | Don't build core flows on beta endpoints; only use beta for nice-to-haves |
| Tenant admin revokes consent | Low | Detect 401/403, prompt re-consent |
| Microsoft Foundry knowledge-source integration churn | Medium | The old SPE agent SDK was deprecated April 2026 in favor of Foundry. Track this carefully if AI is in scope. |

### Open Questions for Architecture Review

1. **Edit experience UX:** popup vs. new tab vs. an iframe with a polite "this would have been inline but Microsoft doesn't allow it" overlay? Recommend new tab.
2. **Container model:** one container per customer org, one per project, or one per user? This affects permission model significantly.
3. **Multi-tenancy:** ISV multi-tenant from day one, or start single-tenant LOB and refactor later? Multi-tenant from day one is materially harder but avoids a rebuild.
4. **AI integration:** if Copilot grounding on container content is in scope, route through SharePoint Embedded knowledge source in Microsoft Foundry. The old SPE agent SDK is deprecated.
5. **Frontend hosting:** Static Web Apps (cheap, simple) or App Service (SSR, more flexible)?
6. **External sharing:** allowed? If yes, anonymous links or B2B guest only?
7. **Sensitivity labels:** required at launch or phase 2?
8. **Existing tenant data migration:** if customers have existing SharePoint or OneDrive content, do we migrate it into SPE containers, or treat SPE as new-content-only?

---

## 13. Implementation Phases

### Phase 0: Validate (1-2 weeks)
- Clone `boilerplate-typescript-react` sample
- Create trial container type via VS Code extension
- Run sample end-to-end: sign in, create container, upload Word doc, preview, edit in Office for the web
- Confirm co-authoring works with two test accounts
- Verify the "edit in popup" UX feels acceptable

### Phase 1: Core MVP (4-6 weeks)
- Production Entra app registration with consent flow
- Standard or Passthrough billing decision and setup
- React + Next.js frontend with MSAL
- ASP.NET Core backend with OBO and Graph SDK
- File browser, upload, download, delete
- Preview iframe
- Edit launch flow
- Sharing (link generation, permission management)
- Single-tenant deployment

### Phase 2: Multi-Tenant + Compliance (3-4 weeks)
- Multi-tenant Entra app, admin consent onboarding
- Container type registration API integration
- Customer admin onboarding flow with billing setup check
- Webhook subscriptions for change notifications
- Sensitivity label awareness
- Audit logging

### Phase 3: Search + AI (3-4 weeks)
- Microsoft Search integration
- Copilot grounding via Microsoft Foundry SPE knowledge source
- AI-assisted document workflows (summarization, classification, etc.)

### Phase 4: Polish (ongoing)
- Versioning UI
- Bulk operations
- Mobile responsive refinement
- Performance tuning
- Cost monitoring dashboards

---

## 14. Reference Links

### Core SharePoint Embedded Docs
- Overview: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/overview
- What's new (most recent: April 2026): https://learn.microsoft.com/en-us/sharepoint/dev/embedded/whats-new
- App architecture: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/app-architecture
- Office file experiences: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/content-experiences/office-experience
- Authentication and authorization: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth

### Getting Started
- VS Code extension: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/getting-started/spembedded-for-vscode
- Container type registration: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/getting-started/register-api-documentation
- Vendor install on customer tenant: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/tutorials/vendor-install-app-customer

### Tutorials
- File preview tutorial: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/tutorials/using-file-preview
- SPE agent + VS Code tutorial: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/tutorials/spe-da-vscode

### Microsoft Graph API References
- `driveItem` resource: https://learn.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0
- `driveItem: preview`: https://learn.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-1.0
- `driveItem: preview` (beta with `allowEdit`): https://learn.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-beta
- Create permission on driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-post-permissions?view=graph-rest-beta
- Create sharing link: https://learn.microsoft.com/en-us/graph/api/driveitem-createlink

### Samples
- Official samples repo: https://github.com/microsoft/SharePoint-Embedded-Samples
- `boilerplate-typescript-react`: React + TypeScript + Azure Functions OBO proxy
- `boilerplate-aspnet-webservice`: C# ASP.NET Core reference
- `legal-docs`: React + Fluent UI + Copilot SDK
- `project-management`: React + Vite + Tailwind + shadcn-ui
- `webhook`: Node.js Graph change notification listener

### Billing
- Billing overview: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/administration/billing/billing
- Billing meters: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/administration/billing/meters
- Billing management: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/administration/billing/billingmanagement
- M365 PAYG pricing: https://learn.microsoft.com/en-us/sharepoint/microsoft-365-pay-as-you-go-pricing

### Connectors & Integrations
- SharePoint Embedded connector (Power Platform): https://learn.microsoft.com/en-us/connectors/sharepointembedded/

### Background / Confirmation that Inline Editable Iframe is Not Supported
- Microsoft Q&A: https://learn.microsoft.com/en-us/answers/questions/5669186/is-there-a-supported-way-to-embed-editable-sharepo

### MSAL & Identity
- MSAL.js for React: https://learn.microsoft.com/en-us/entra/identity-platform/tutorial-v2-react
- Microsoft.Identity.Web (backend OBO): https://learn.microsoft.com/en-us/entra/identity-platform/scenario-web-api-call-api-overview
- On-Behalf-Of flow: https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow

### Microsoft Graph SDK
- Graph SDK for .NET: https://learn.microsoft.com/en-us/graph/sdks/sdks-overview
- Graph SDK for JavaScript: https://github.com/microsoftgraph/msgraph-sdk-javascript

---

## Appendix A: Sample Code Sketches

### Frontend: Preview iframe in React

```tsx
async function getPreviewUrl(driveId: string, itemId: string): Promise<string> {
  const response = await fetch(`/api/files/${driveId}/${itemId}/preview`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${await getToken()}` }
  });
  const { getUrl } = await response.json();
  return `${getUrl}&nb=true`; // nb=true removes the banner
}

function DocumentPreview({ driveId, itemId }: Props) {
  const { data: url } = useQuery({
    queryKey: ['preview', driveId, itemId],
    queryFn: () => getPreviewUrl(driveId, itemId),
    staleTime: 60_000, // refresh every 60s for safety
  });

  if (!url) return <Loading />;
  return <iframe src={url} className="w-full h-full" title="Preview" />;
}
```

### Frontend: Launch edit in new window

```tsx
function EditButton({ driveItem }: { driveItem: DriveItem }) {
  const handleEdit = () => {
    const editWindow = window.open(driveItem.webUrl, '_blank', 'noopener');
    if (!editWindow) {
      toast.error('Pop-up blocked. Allow pop-ups for this site to edit.');
    }
  };

  return <Button onClick={handleEdit}>Edit in Word</Button>;
}
```

### Backend: OBO + Graph preview call (ASP.NET Core)

```csharp
[HttpPost("api/files/{driveId}/{itemId}/preview")]
public async Task<IActionResult> GetPreview(string driveId, string itemId)
{
    var preview = await _graphServiceClient
        .Drives[driveId]
        .Items[itemId]
        .Preview
        .PostAsync(new() { /* page, zoom, etc. */ });

    return Ok(new { getUrl = preview?.GetUrl });
}
```

### Backend: Get drive item with webUrl

```csharp
var driveItem = await _graphServiceClient
    .Drives[driveId]
    .Items[itemId]
    .GetAsync();

// driveItem.WebUrl is what the frontend opens for editing
return Ok(new { webUrl = driveItem.WebUrl, name = driveItem.Name });
```

---

## Appendix B: What This PRD Does Not Cover

- Detailed UI/UX wireframes (separate design doc)
- Specific data model for app-side metadata (you'll want your own DB for app-specific data: tags, workspaces, projects, etc.)
- Disaster recovery and backup strategy (SPE has tenant-side backup story, but app-side state needs its own plan)
- Pricing model for your end customers
- Go-to-market and sales motion
- Specific Copilot/AI features beyond the integration point

---

*End of PRD.*