# Contributing

Thanks for your interest in Wordembedded. The project is a reference implementation, but we
welcome bug fixes, doc improvements, and well-scoped enhancements.

## Before you open a PR

1. **Read `README.md`** for the architecture and the Microsoft-documented constraints we follow.
2. **Reproduce the issue or proposed change locally**. Provisioning is described in
   `src/infra/README.md`; configuration is propagated by `src/infra/sync-config.ps1`.
3. **Make sure both builds pass**:
   - `cd src/backend && dotnet build` — should report 0 warnings and 0 errors
   - `cd src/frontend && npm run build` — should report a clean Turbopack build
4. **Don't commit secrets.** `outputs.json`, `appsettings.Development.json`, `.env.local`, and
   anything under `src/infra/secrets/` are gitignored. Use the `.example.json` files instead.

## Code conventions

- Backend: minimal API endpoints grouped under `src/backend/Wordembedded.Api/Endpoints/`. DTOs in
  `Contracts/`. Don't return raw Graph SDK types from public endpoints — map through DTOs.
- Frontend: Fluent UI v9 only (`@fluentui/react-components`). No global CSS — use `makeStyles` and
  Fluent design tokens. TanStack Query for all server state.
- Microsoft Graph behavior should match the cited docs. If you're tempted to add a workaround,
  link to the relevant Microsoft Learn page in the code comment or PR description.

## Pull request checklist

- [ ] Targets `main`
- [ ] Both builds pass locally
- [ ] No secrets in the diff
- [ ] Microsoft Graph endpoints / scopes / conflict-behaviors are linked to the official docs
- [ ] README / SECURITY / docs updated if behavior changed
