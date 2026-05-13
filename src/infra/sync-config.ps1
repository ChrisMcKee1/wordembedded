#requires -Version 7.0
<#
.SYNOPSIS
  Sync src/infra/outputs.json into the backend appsettings.Development.json and the frontend .env.local.

.DESCRIPTION
  Run this once you have completed the manual provisioning steps in src/infra/README.md and filled in
  outputs.json (clientId, containerType.id, initialContainer.id + driveId, certificate thumbprint, etc.).

  - Backend: writes appsettings.Development.json with TenantId / ClientId / Audience / ClientCertificates
    and SharePointEmbedded.ContainerTypeId. Webhook ClientStateSecret is generated if empty.
  - Frontend: writes src/frontend/.env.local with NEXT_PUBLIC_AAD_* + API base + API scope.

.NOTES
  Sensitive values (cert password, webhook client state secret) are kept in user-secrets / env files
  that are gitignored.
#>

[CmdletBinding()]
param(
  [string]$OutputsPath = (Join-Path $PSScriptRoot "outputs.json"),
  [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $OutputsPath)) {
  throw "outputs.json not found at $OutputsPath. Run the provisioning steps in src/infra/README.md first."
}

$outputs = Get-Content $OutputsPath -Raw | ConvertFrom-Json

if ($outputs.PSObject.Properties.Match('blockers').Count -and $outputs.blockers.Count -gt 0) {
  Write-Warning "outputs.json reports blockers:"
  $outputs.blockers | ForEach-Object { Write-Warning "  - $_" }
  Write-Warning "Continuing anyway — null values will be written as REPLACE_IN_DEV placeholders."
}

function Resolve-Or-Placeholder($value, [string]$placeholder = "REPLACE_IN_DEV") {
  if ([string]::IsNullOrWhiteSpace($value)) { return $placeholder }
  return $value
}

$tenantId   = Resolve-Or-Placeholder $outputs.tenantId
$clientId   = Resolve-Or-Placeholder $outputs.appRegistration.clientId
$thumb      = Resolve-Or-Placeholder $outputs.appRegistration.certificateThumbprint
$ctId       = Resolve-Or-Placeholder $outputs.containerType.id
$audience   = if ($clientId -eq "REPLACE_IN_DEV") { "api://REPLACE_IN_DEV" } else { "api://$clientId" }

# --- Backend ---
$backendCfg = Join-Path $RepoRoot "src\backend\Wordembedded.Api\appsettings.Development.json"
$cfg = [ordered]@{
  Logging = [ordered]@{
    LogLevel = [ordered]@{
      Default                 = "Information"
      "Microsoft.AspNetCore"  = "Warning"
    }
  }
  AzureAd = [ordered]@{
    Instance           = "https://login.microsoftonline.com/"
    TenantId           = $tenantId
    ClientId           = $clientId
    Audience           = $audience
    ClientCertificates = @(
      [ordered]@{
        SourceType            = "StoreWithThumbprint"
        CertificateStorePath  = "CurrentUser/My"
        CertificateThumbprint = $thumb
      }
    )
  }
  DownstreamApi = [ordered]@{
    BaseUrl = "https://graph.microsoft.com/v1.0"
    Scopes  = @(
      "https://graph.microsoft.com/FileStorageContainer.Selected",
      "https://graph.microsoft.com/FileStorageContainerTypeReg.Selected",
      "https://graph.microsoft.com/Files.Read.All",
      "https://graph.microsoft.com/User.Read"
    )
  }
  SharePointEmbedded = [ordered]@{
    ContainerTypeId = $ctId
  }
  Webhooks = [ordered]@{
    NotificationUrl     = "https://localhost:7071/api/webhooks/graph"
    ClientStateSecret   = ([Guid]::NewGuid().ToString("N"))
  }
}
$cfg | ConvertTo-Json -Depth 10 | Set-Content -Path $backendCfg -Encoding UTF8
Write-Host "Backend config written:" $backendCfg

# --- Frontend ---
$frontendEnv = Join-Path $RepoRoot "src\frontend\.env.local"
$apiBase   = "https://localhost:7071"
$apiScope  = if ($clientId -eq "REPLACE_IN_DEV") { "api://REPLACE_IN_DEV/access_as_user" } else { "api://$clientId/access_as_user" }

@(
  "NEXT_PUBLIC_AAD_CLIENT_ID=$clientId"
  "NEXT_PUBLIC_AAD_TENANT_ID=$tenantId"
  "NEXT_PUBLIC_API_BASE_URL=$apiBase"
  "NEXT_PUBLIC_API_SCOPE=$apiScope"
) | Set-Content -Path $frontendEnv -Encoding UTF8
Write-Host "Frontend env written:" $frontendEnv

Write-Host ""
Write-Host "Done. If any value above is REPLACE_IN_DEV, finish the provisioning steps in src/infra/README.md and re-run sync-config.ps1."
