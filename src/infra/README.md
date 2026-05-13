# Infrastructure provisioning

This folder holds the artifacts the app needs from Entra ID and SharePoint Embedded:

- `outputs.json` — IDs the backend/frontend read at runtime (clientId, tenantId, containerTypeId, driveId, etc.)
- `secrets/wordembedded-dev.pfx` — self-signed certificate for app-only Graph auth (gitignored)
- `secrets/cert-password.txt` — password for the PFX (gitignored)

## Status

`outputs.json` is fully populated:

- **Tenant**: `<your-tenant-guid>` (`<your-tenant>.onmicrosoft.com`)
- **App**: `Wordembedded Dev` (clientId `<entra-app-client-id>`)
- **Certificate**: thumbprint `<cert-thumbprint>`, PFX at `secrets/wordembedded-dev.pfx`
- **Container type**: `Wordembedded.Documents` (trial, id `<container-type-guid>`, expires 2026-06-12)
- **Initial container**: `Workspace` (id+driveId `<container-id-starts-with-b!>`)

Run `pwsh sync-config.ps1` to push these into the backend and frontend config files.

If you ever need to **re-provision** (different tenant, lost cert, expired trial), follow the steps below.

> Prerequisites: Owner of the M365 tenant `<your-tenant>.onmicrosoft.com`, `PnP.PowerShell`,
> `Microsoft.Graph` PowerShell module, and Azure CLI installed.

### 1. Refresh the Azure CLI session

```powershell
az logout
az login --tenant <your-tenant-guid>
```

Sign in as the user that owns the SPO admin site (`<your-admin@your-tenant.onmicrosoft.com>`).

### 2. Create the Entra app registration

```powershell
$tenantId = "<your-tenant-guid>"
$app = az ad app create --display-name "Wordembedded Dev" --sign-in-audience AzureADMultipleOrgs | ConvertFrom-Json
$clientId = $app.appId
$objectId = $app.id
az ad sp create --id $clientId | Out-Null
```

### 3. Add redirect URIs and expose an API scope

Patch the app via Microsoft Graph because `az ad app update` doesn't fully cover SPA URIs or
custom API scopes:

```powershell
$body = @{
  spa = @{ redirectUris = @("http://localhost:3000/auth/callback") }
  web = @{ redirectUris = @("https://localhost:7071/signin-oidc") }
  identifierUris = @("api://$clientId")
  api = @{
    oauth2PermissionScopes = @(@{
      id = [guid]::NewGuid().ToString()
      adminConsentDescription = "Allows the app to access Wordembedded API on behalf of the signed-in user."
      adminConsentDisplayName = "Access Wordembedded API"
      isEnabled = $true
      type = "User"
      value = "access_as_user"
      userConsentDescription = "Allow this app to access Wordembedded API as you."
      userConsentDisplayName = "Access Wordembedded API as you"
    })
  }
} | ConvertTo-Json -Depth 10

az rest --method PATCH `
  --uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
  --headers "Content-Type=application/json" `
  --body $body
```

### 4. Add required Graph + SharePoint Online permissions

Resource IDs:

| Resource | App ID |
|---|---|
| Microsoft Graph | `00000003-0000-0000-c000-000000000000` |
| SharePoint Online | `00000003-0000-0ff1-ce00-000000000000` |

Look up the permission GUIDs once:

```powershell
$graphSp = az ad sp show --id 00000003-0000-0000-c000-000000000000 | ConvertFrom-Json
$spoSp   = az ad sp show --id 00000003-0000-0ff1-ce00-000000000000 | ConvertFrom-Json

function Find-Scope($sp, $value)   { ($sp.oauth2PermissionScopes | Where-Object value -eq $value).id }
function Find-AppRole($sp, $value) { ($sp.appRoles                | Where-Object value -eq $value).id }

$perms = @{
  requiredResourceAccess = @(
    @{ resourceAppId = "00000003-0000-0000-c000-000000000000"; resourceAccess = @(
       @{ id = (Find-Scope $graphSp "FileStorageContainer.Selected"); type = "Scope" },
       @{ id = (Find-AppRole $graphSp "FileStorageContainer.Selected"); type = "Role"  },
       @{ id = (Find-AppRole $graphSp "FileStorageContainerType.Manage.All"); type = "Role" },
       @{ id = (Find-Scope $graphSp "FileStorageContainerTypeReg.Selected"); type = "Scope" },
       @{ id = (Find-AppRole $graphSp "FileStorageContainerTypeReg.Selected"); type = "Role" },
       @{ id = (Find-Scope $graphSp "User.Read"); type = "Scope" },
       @{ id = (Find-Scope $graphSp "Files.Read.All"); type = "Scope" }
    )},
    @{ resourceAppId = "00000003-0000-0ff1-ce00-000000000000"; resourceAccess = @(
       @{ id = (Find-AppRole $spoSp "Container.Selected"); type = "Role" }
    )}
  )
} | ConvertTo-Json -Depth 10

az rest --method PATCH `
  --uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
  --headers "Content-Type=application/json" `
  --body $perms
```

### 5. Upload the existing certificate to the app registration

```powershell
$certPath = "secrets\wordembedded-dev.cer"  # public key only — DO NOT upload the PFX
$certBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes((Resolve-Path $certPath)))

$kc = @{
  keyCredentials = @(@{
    type = "AsymmetricX509Cert"
    usage = "Verify"
    key = $certBytes
    displayName = "Wordembedded Dev cert"
  })
} | ConvertTo-Json -Depth 10

az rest --method PATCH `
  --uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
  --headers "Content-Type=application/json" `
  --body $kc
```

The thumbprint is already stored in `outputs.json` (look at `appRegistration.certificateThumbprint`).
If you regenerate the cert, also rerun `New-SelfSignedCertificate` and re-export `.cer` and `.pfx`.

### 6. Grant admin consent

```powershell
az ad app permission admin-consent --id $clientId
# Or open the explicit consent URL in a browser as a tenant admin:
Start-Process "https://login.microsoftonline.com/$tenantId/v2.0/adminconsent?client_id=$clientId&scope=https://graph.microsoft.com/.default"
```

### 7. Create the SharePoint Embedded container type

Use the official SharePoint Online PowerShell module (the PnP module's container-type cmdlets
require admin scopes that PnP's default token doesn't carry):

```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force -AllowClobber
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url "https://<your-tenant>-admin.sharepoint.com"
# Sign in as a Global Administrator or SharePoint Embedded Administrator

$ct = New-SPOContainerType -ContainerTypeName "Wordembedded.Documents" `
                           -OwningApplicationId $clientId `
                           -TrialContainerType
$ct.ContainerTypeId
```

Confirm the Y prompt when asked.

### 8. Register the container type on the consuming tenant

The modern (and **only working**) registration path is the **Microsoft Graph** endpoint at
`PUT /storage/fileStorage/containerTypeRegistrations/{ctId}` — **not** the older SPO REST
`/_api/v2.1/storageContainerTypes/{ctId}/applicationPermissions` endpoint, which is being
deprecated and returns `accessDenied` for tokens that don't have the hidden SPO `Container.Selected`
role (a permission that no longer appears in modern SPO tenants).

This call uses an **app-only** token from the owning app:

```powershell
Install-Module MSAL.PS -Scope CurrentUser -Force -AcceptLicense -AllowClobber
Import-Module MSAL.PS

$cert = Get-Item "Cert:\CurrentUser\My\$thumb"
$graphToken = (Get-MsalToken -ClientId $clientId -TenantId $tenantId `
              -ClientCertificate $cert `
              -Scopes "https://graph.microsoft.com/.default").AccessToken

$registrationBody = @{
  applicationPermissionGrants = @(@{
    appId = $clientId
    delegatedPermissions   = @("full")
    applicationPermissions = @("full")
  })
} | ConvertTo-Json -Depth 10

Invoke-RestMethod `
  -Uri "https://graph.microsoft.com/v1.0/storage/fileStorage/containerTypeRegistrations/$($ct.ContainerTypeId)" `
  -Method Put `
  -Headers @{ Authorization = "Bearer $graphToken"; Accept = "application/json" } `
  -ContentType "application/json" `
  -Body $registrationBody
```

Reference: <https://learn.microsoft.com/en-us/graph/api/filestorage-post-containertyperegistrations>

### 9. Create the initial container

The owning app calls Graph directly with app-only auth:

```powershell
$gToken = (Get-MsalToken -ClientId $clientId -TenantId $tenantId `
          -ClientCertificate $cert `
          -Scopes "https://graph.microsoft.com/.default").AccessToken

$body = @{
  displayName = "Workspace"
  description = "Wordembedded primary workspace"
  containerTypeId = $ct.ContainerTypeId
} | ConvertTo-Json

$container = Invoke-RestMethod -Method POST `
  -Uri "https://graph.microsoft.com/v1.0/storage/fileStorage/containers" `
  -Headers @{ Authorization = "Bearer $gToken"; Accept = "application/json" } `
  -ContentType "application/json" -Body $body

# Activate (a no-op return for trial CTs, but required call)
Invoke-RestMethod -Method POST `
  -Uri "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/$($container.id)/activate" `
  -Headers @{ Authorization = "Bearer $gToken" } -ContentType "application/json" -Body "{}"

$drive = Invoke-RestMethod -Method GET `
  -Uri "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/$($container.id)/drive" `
  -Headers @{ Authorization = "Bearer $gToken" }
```

### 10. Fill in `outputs.json`

Update these fields:

```json
{
  "appRegistration": { "clientId": "<from step 2>", "objectId": "<from step 2>", "apiScope": "api://<clientId>/access_as_user" },
  "containerType": { "id": "<from step 7>", "name": "Wordembedded.Documents", "billingType": "trial" },
  "initialContainer": { "id": "<container.id>", "driveId": "<drive.id>", "displayName": "Workspace" },
  "blockers": []
}
```

Then run `pwsh src\infra\sync-config.ps1` (created during integration) to push the IDs into the
backend `appsettings.Development.json` and the frontend `.env.local`.

## References

- App architecture: <https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/app-architecture>
- Authentication: <https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth>
- Container type registration: <https://learn.microsoft.com/en-us/sharepoint/dev/embedded/getting-started/register-api-documentation>
- Vendor install on customer tenant: <https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/tutorials/vendor-install-app-customer>
- Permission reference: <https://learn.microsoft.com/en-us/graph/permissions-reference>
