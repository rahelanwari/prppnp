#requires -Modules PnP.PowerShell
$ErrorActionPreference = "Stop"

# ====== CONFIG ======
$SiteUrl     = $env:SITE_URL
$TenantId    = $env:TENANT_ID
$ClientId    = $env:CLIENT_ID
$PfxBase64   = $env:CERT_PFX_BASE64
$PfxPassword = $env:CERT_PASSWORD

if (-not $SiteUrl -or -not $TenantId -or -not $ClientId -or -not $PfxBase64 -or -not $PfxPassword) {
  throw "Missing required environment variables. Check GitHub Secrets."
}

# ====== AUTH ======
Write-Host "Connecting to SharePoint site: $SiteUrl"

$pfxBytes = [Convert]::FromBase64String($PfxBase64)
$tempPfx  = Join-Path $env:RUNNER_TEMP "pnp-cert.pfx"
[IO.File]::WriteAllBytes($tempPfx, $pfxBytes)

Connect-PnPOnline -Url $SiteUrl -Tenant $TenantId -ClientId $ClientId -CertificatePath $tempPfx -CertificatePassword (ConvertTo-SecureString $PfxPassword -AsPlainText -Force)

Write-Host "Connected."

# ====== HELPERS ======
function Ensure-ChoiceField {
  param(
    [string]$ListTitle,
    [string]$InternalName,
    [string]$DisplayName,
    [string[]]$Choices,
    [string]$Description = ""
  )

  $list = Get-PnPList -Identity $ListTitle -ErrorAction Stop
  $field = Get-PnPField -List $list -Identity $InternalName -ErrorAction SilentlyContinue

  if (-not $field) {
    Write-Host "Creating field '$DisplayName' ($InternalName) on list '$ListTitle'..."
    Add-PnPField -List $list -InternalName $InternalName -DisplayName $DisplayName -Type Choice -AddToDefaultView:$false -Choices $Choices
    if ($Description) {
      Set-PnPField -List $list -Identity $InternalName -Values @{ Description = $Description }
    }
  } else {
    Write-Host "Field '$InternalName' already exists on '$ListTitle'."
    # Optional: update choices (careful: can remove existing values)
  }
}

function Ensure-View {
  param(
    [string]$ListTitle,
    [string]$ViewName,
    [string[]]$ViewFields,
    [string]$CamlQuery
  )

  $view = Get-PnPView -List $ListTitle -Identity $ViewName -ErrorAction SilentlyContinue
  if (-not $view) {
    Write-Host "Creating view '$ViewName' on '$ListTitle'..."
    Add-PnPView -List $ListTitle -Title $ViewName -Fields $ViewFields -Query $CamlQuery
  } else {
    Write-Host "View '$ViewName' already exists on '$ListTitle'."
  }
}

# ====== YOUR LIBRARIES ======
# Library names must match exactly what you see in SharePoint
$ArchitectureLibrary = "Architecture"
$SecurityLibrary     = "Security"
$OperationsLibrary   = "Operations"
$ProjectLibrary      = "Project"

# ====== COLUMNS (based on what you created) ======
# Architecture - Document Type
Ensure-ChoiceField `
  -ListTitle $ArchitectureLibrary `
  -InternalName "DocumentType" `
  -DisplayName "Document Type" `
  -Description "Classifies the type of architecture document." `
  -Choices @("Diagram","Design Document","Decision Records (ADR)","Reference","Other")

# Security - Document Type
Ensure-ChoiceField `
  -ListTitle $SecurityLibrary `
  -InternalName "DocumentType" `
  -DisplayName "Document Type" `
  -Choices @("Security Diagrams","Security Policies","Risk Assessment","Threat Model","Incident Procedures","Compliance Evidence","Other")

# Operations - Operation Type
Ensure-ChoiceField `
  -ListTitle $OperationsLibrary `
  -InternalName "OperationType" `
  -DisplayName "Operation Type" `
  -Choices @("Runbook","Standard Operating Procedure","Incident Handling","Maintenance","Monitoring & Alerts","Backup & Recovery","Change Management","Other")

# Project - Project Artifact Type
Ensure-ChoiceField `
  -ListTitle $ProjectLibrary `
  -InternalName "ProjectArtifactType" `
  -DisplayName "Project Artifact Type" `
  -Choices @("PRP","Research","Sprint Artifact","Presentation","Deliverable","Other")

# ====== VIEWS (filter by choice column) ======
# CAML filters: field internal name must match the internal name you used.
# IMPORTANT: Choice value must match EXACTLY, case-sensitive.

Ensure-View `
  -ListTitle $ArchitectureLibrary `
  -ViewName "Diagrams" `
  -ViewFields @("DocIcon","LinkFilename","Modified","Editor","DocumentType") `
  -CamlQuery "<Where><Eq><FieldRef Name='DocumentType'/><Value Type='Choice'>Diagram</Value></Eq></Where>"

Ensure-View `
  -ListTitle $ArchitectureLibrary `
  -ViewName "Design Documents" `
  -ViewFields @("DocIcon","LinkFilename","Modified","Editor","DocumentType") `
  -CamlQuery "<Where><Eq><FieldRef Name='DocumentType'/><Value Type='Choice'>Design Document</Value></Eq></Where>"

Ensure-View `
  -ListTitle $ArchitectureLibrary `
  -ViewName "Decision Records (ADR)" `
  -ViewFields @("DocIcon","LinkFilename","Modified","Editor","DocumentType") `
  -CamlQuery "<Where><Eq><FieldRef Name='DocumentType'/><Value Type='Choice'>Decision Records (ADR)</Value></Eq></Where>"

# Security examples
Ensure-View `
  -ListTitle $SecurityLibrary `
  -ViewName "Security Policies" `
  -ViewFields @("DocIcon","LinkFilename","Modified","Editor","DocumentType") `
  -CamlQuery "<Where><Eq><FieldRef Name='DocumentType'/><Value Type='Choice'>Security Policies</Value></Eq></Where>"

Ensure-View `
  -ListTitle $SecurityLibrary `
  -ViewName "Compliance Evidence" `
  -ViewFields @("DocIcon","LinkFilename","Modified","Editor","DocumentType") `
  -CamlQuery "<Where><Eq><FieldRef Name='DocumentType'/><Value Type='Choice'>Compliance Evidence</Value></Eq></Where>"

# Operations examples
Ensure-View `
  -ListTitle $OperationsLibrary `
  -ViewName "Runbooks" `
  -ViewFields @("DocIcon","LinkFilename","Modified","Editor","OperationType") `
  -CamlQuery "<Where><Eq><FieldRef Name='OperationType'/><Value Type='Choice'>Runbook</Value></Eq></Where>"

Ensure-View `
  -ListTitle $OperationsLibrary `
  -ViewName "Change Management" `
  -ViewFields @("DocIcon","LinkFilename","Modified","Editor","OperationType") `
  -CamlQuery "<Where><Eq><FieldRef Name='OperationType'/><Value Type='Choice'>Change Management</Value></Eq></Where>"

# Project examples
Ensure-View `
  -ListTitle $ProjectLibrary `
  -ViewName "PRP" `
  -ViewFields @("DocIcon","LinkFilename","Modified","Editor","ProjectArtifactType") `
  -CamlQuery "<Where><Eq><FieldRef Name='ProjectArtifactType'/><Value Type='Choice'>PRP</Value></Eq></Where>"

Ensure-View `
  -ListTitle $ProjectLibrary `
  -ViewName "Deliverable" `
  -ViewFields @("DocIcon","LinkFilename","Modified","Editor","ProjectArtifactType") `
  -CamlQuery "<Where><Eq><FieldRef Name='ProjectArtifactType'/><Value Type='Choice'>Deliverable</Value></Eq></Where>"

Write-Host "Provisioning completed successfully."
