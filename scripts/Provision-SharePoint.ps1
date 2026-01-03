#requires -Modules PnP.PowerShell

param(
  [Parameter(Mandatory=$true)]
  [string]$PfxPath
)

$ErrorActionPreference = "Stop"

# ====== CONFIG (from env) ======
$SiteUrl     = $env:SITE_URL
$TenantId    = $env:TENANT_ID
$ClientId    = $env:CLIENT_ID
$PfxPassword = $env:CERT_PASSWORD

if (-not $SiteUrl -or -not $TenantId -or -not $ClientId -or -not $PfxPassword) {
  throw "Missing required environment variables. Check GitHub Secrets (SITE_URL, TENANT_ID, CLIENT_ID, CERT_PASSWORD)."
}

if (!(Test-Path $PfxPath)) {
  throw "PFX file not found at path: $PfxPath"
}

Write-Host "== PnP Provisioning starting =="
Write-Host "Connecting to SharePoint site: $SiteUrl"

try {
  # ====== AUTH ======
  Connect-PnPOnline `
    -Url $SiteUrl `
    -Tenant $TenantId `
    -ClientId $ClientId `
    -CertificatePath $PfxPath `
    -CertificatePassword (ConvertTo-SecureString $PfxPassword -AsPlainText -Force)

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

    Write-Host "---- Ensure field [$InternalName] on [$ListTitle]"

    $list  = Get-PnPList -Identity $ListTitle -ErrorAction Stop
    $field = Get-PnPField -List $list -Identity $InternalName -ErrorAction SilentlyContinue

    if (-not $field) {
      Write-Host "Creating field '$DisplayName' ($InternalName)..."
      Add-PnPField -List $list -InternalName $InternalName -DisplayName $DisplayName -Type Choice -AddToDefaultView:$false -Choices $Choices | Out-Null

      if ($Description) {
        Set-PnPField -List $list -Identity $InternalName -Values @{ Description = $Description } | Out-Null
      }

      Write-Host "Field created."
      return
    }

    # Merge choices (no deletes)
    $existingChoices = @()
    if ($field.Choices) { $existingChoices = @($field.Choices) }

    $merged = @($existingChoices + $Choices | Select-Object -Unique)

    if ($merged.Count -ne $existingChoices.Count) {
      Write-Host "Updating choices (adding missing values only)..."
      # For choice fields, updating via SchemaXml is the most reliable
      $xml = [xml]$field.SchemaXml
      $choicesNode = $xml.Field.Choices
      $choicesNode.RemoveAll() | Out-Null

      foreach ($c in $merged) {
        $choice = $xml.CreateElement("CHOICE")
        $choice.InnerText = $c
        $choicesNode.AppendChild($choice) | Out-Null
      }

      Set-PnPField -List $list -Identity $InternalName -Values @{ SchemaXml = $xml.OuterXml } | Out-Null
      Write-Host "Choices updated."
    }
    else {
      Write-Host "Field exists and choices are already OK."
    }
  }

  function Ensure-View {
    param(
      [string]$ListTitle,
      [string]$ViewName,
      [string[]]$ViewFields,
      [string]$CamlWhere
    )

    # Always wrap as Query for consistency (used only for Add-PnPView)
    $query = "<Query>$CamlWhere</Query>"

    Write-Host "---- Ensure view [$ViewName] on [$ListTitle]"
    $view = Get-PnPView -List $ListTitle -Identity $ViewName -ErrorAction SilentlyContinue

    if (-not $view) {
      Write-Host "Creating view..."
      Add-PnPView -List $ListTitle -Title $ViewName -Fields $ViewFields -Query $query | Out-Null
      Write-Host "View created."
      return
    }

    # IMPORTANT: Set-PnPView does not support -Query in some PnP.PowerShell versions on GitHub runners
    Write-Host "View exists. Updating fields only (Query not changed to avoid -Query compatibility issues)..."
    Set-PnPView -List $ListTitle -Identity $ViewName -Fields $ViewFields | Out-Null
    Write-Host "View updated."
  }

  # ====== LIBRARIES ======
  $ArchitectureLibrary = "Architecture"
  $SecurityLibrary     = "Security"
  $OperationsLibrary   = "Operations"
  $ProjectLibrary      = "Project"

  # ====== COLUMNS ======
  Ensure-ChoiceField -ListTitle $ArchitectureLibrary -InternalName "DocumentType" -DisplayName "Document Type" `
    -Description "Classifies the type of architecture document." `
    -Choices @("Diagram","Design Document","Decision Records (ADR)","Reference","Other")

  Ensure-ChoiceField -ListTitle $SecurityLibrary -InternalName "DocumentType" -DisplayName "Document Type" `
    -Choices @("Security Diagrams","Security Policies","Risk Assessment","Threat Model","Incident Procedures","Compliance Evidence","Other")

  Ensure-ChoiceField -ListTitle $OperationsLibrary -InternalName "OperationType" -DisplayName "Operation Type" `
    -Choices @("Runbook","Standard Operating Procedure","Incident Handling","Maintenance","Monitoring & Alerts","Backup & Recovery","Change Management","Other")

  Ensure-ChoiceField -ListTitle $ProjectLibrary -InternalName "ProjectArtifactType" -DisplayName "Project Artifact Type" `
    -Choices @("PRP","Research","Sprint Artifact","Presentation","Deliverable","Other")

  # ====== VIEWS ======
  Ensure-View -ListTitle $ArchitectureLibrary -ViewName "Diagrams" `
    -ViewFields @("DocIcon","LinkFilename","Modified","Editor","DocumentType") `
    -CamlWhere "<Where><Eq><FieldRef Name='DocumentType'/><Value Type='Choice'>Diagram</Value></Eq></Where>"

  Ensure-View -ListTitle $ArchitectureLibrary -ViewName "Design Documents" `
    -ViewFields @("DocIcon","LinkFilename","Modified","Editor","DocumentType") `
    -CamlWhere "<Where><Eq><FieldRef Name='DocumentType'/><Value Type='Choice'>Design Document</Value></Eq></Where>"

  Ensure-View -ListTitle $ArchitectureLibrary -ViewName "Decision Records (ADR)" `
    -ViewFields @("DocIcon","LinkFilename","Modified","Editor","DocumentType") `
    -CamlWhere "<Where><Eq><FieldRef Name='DocumentType'/><Value Type='Choice'>Decision Records (ADR)</Value></Eq></Where>"

  Ensure-View -ListTitle $SecurityLibrary -ViewName "Security Policies" `
    -ViewFields @("DocIcon","LinkFilename","Modified","Editor","DocumentType") `
    -CamlWhere "<Where><Eq><FieldRef Name='DocumentType'/><Value Type='Choice'>Security Policies</Value></Eq></Where>"

  Ensure-View -ListTitle $SecurityLibrary -ViewName "Compliance Evidence" `
    -ViewFields @("DocIcon","LinkFilename","Modified","Editor","DocumentType") `
    -CamlWhere "<Where><Eq><FieldRef Name='DocumentType'/><Value Type='Choice'>Compliance Evidence</Value></Eq></Where>"

  Ensure-View -ListTitle $OperationsLibrary -ViewName "Runbooks" `
    -ViewFields @("DocIcon","LinkFilename","Modified","Editor","OperationType") `
    -CamlWhere "<Where><Eq><FieldRef Name='OperationType'/><Value Type='Choice'>Runbook</Value></Eq></Where>"

  Ensure-View -ListTitle $OperationsLibrary -ViewName "Change Management" `
    -ViewFields @("DocIcon","LinkFilename","Modified","Editor","OperationType") `
    -CamlWhere "<Where><Eq><FieldRef Name='OperationType'/><Value Type='Choice'>Change Management</Value></Eq></Where>"

  Ensure-View -ListTitle $ProjectLibrary -ViewName "PRP" `
    -ViewFields @("DocIcon","LinkFilename","Modified","Editor","ProjectArtifactType") `
    -CamlWhere "<Where><Eq><FieldRef Name='ProjectArtifactType'/><Value Type='Choice'>PRP</Value></Eq></Where>"

  Ensure-View -ListTitle $ProjectLibrary -ViewName "Deliverable" `
    -ViewFields @("DocIcon","LinkFilename","Modified","Editor","ProjectArtifactType") `
    -CamlWhere "<Where><Eq><FieldRef Name='ProjectArtifactType'/><Value Type='Choice'>Deliverable</Value></Eq></Where>"

  Write-Host "== Provisioning completed successfully. =="

}
finally {
  try { Disconnect-PnPOnline | Out-Null } catch {}
}
