# Security Policy

## Reporting a vulnerability

Wordembedded is a reference implementation that integrates with **Microsoft Entra ID** and
**Microsoft Graph (SharePoint Embedded)**. If you discover a security issue that affects this
project's code, please **do not open a public GitHub issue**. Instead, report it privately so it
can be triaged before disclosure:

1. Use [GitHub's private vulnerability reporting](https://docs.github.com/en/code-security/security-advisories/guidance-on-reporting-and-writing-information-about-vulnerabilities/privately-reporting-a-security-vulnerability)
   on this repository (the **Security** tab → **Report a vulnerability**).
2. If you can't use that workflow, email the repository owner directly through the contact
   information on their GitHub profile.

When reporting, please include:

- A clear description of the issue and its impact
- Steps to reproduce (a proof of concept is ideal)
- The commit hash or release version where the issue was observed
- Any suggested remediation

We aim to acknowledge reports within **3 business days** and to publish a fix (or a documented
mitigation) within **30 days** for high-severity findings. Lower-severity findings may take longer
to address.

## Scope

In scope for this repository:

- `src/backend/` — the ASP.NET Core 9 minimal API and its auth/OBO + webhook handling
- `src/frontend/` — the Next.js 16 + React 19 + MSAL.js client
- `src/infra/` — provisioning scripts and helpers

Out of scope (please report to the upstream owner):

- **Microsoft Graph**, **SharePoint Embedded**, **Microsoft Entra ID**, **Office for the web** — see
  the [Microsoft Security Response Center](https://msrc.microsoft.com/) (MSRC) and the
  [Microsoft Bug Bounty Programs](https://www.microsoft.com/en-us/msrc/bounty).
- **.NET Aspire** — see the [Aspire security policy](https://github.com/dotnet/aspire/security).
- **MSAL.js**, **Microsoft.Identity.Web**, **Microsoft Graph SDK** — see the
  [Microsoft identity-platform repos](https://github.com/AzureAD).
- **Fluent UI**, **TanStack Query**, **Next.js**, **React** — see their respective project security
  policies on GitHub.

## Hardening guidance for deployers

This project is structured as a development reference. If you adapt it for production, please:

1. **Rotate the development certificate** and use a managed identity or Key Vault-backed cert for
   any non-local environment. Do not reuse the local `wordembedded-dev.pfx` outside of a developer
   machine.
2. **Never commit** `outputs.json`, `appsettings.Development.json`, `.env.local`, or anything under
   `src/infra/secrets/`. The repository's `.gitignore` covers them; if you add new files, audit
   their contents before committing.
3. **Limit the Graph permissions** to the minimum your scenario needs. The reference scopes (see
   `README.md` §4) include `Sites.ReadWrite.All` because the ETL flow publishes to regular
   SharePoint libraries. If you don't need that, drop the scope.
4. **Use the SPE container type permission system** — when registering the container type on a
   consuming tenant, set the minimum container type permissions per guest application. Don't grant
   `full` to every app.
5. **Validate the `clientState`** on every webhook notification. The project ships a per-instance
   `Webhooks:ClientStateSecret`; rotate it if you suspect leakage.
6. **Set a content security policy** in your hosting environment. The Next.js app uses the
   Microsoft preview iframe; restrict `frame-src` and `connect-src` to the Microsoft origins your
   tenant relies on.
7. **Review Microsoft's [SPE security and compliance docs](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/compliance/security-and-compliance)**
   and apply tenant-level controls (Purview DLP, retention, sensitivity labels) appropriate to
   your data.

## Disclosure

Once a fix lands, we'll publish a GitHub security advisory and credit reporters by name unless
they request otherwise. Coordinated disclosure timelines are at our discretion but we will not sit
on a fix once it's available.
