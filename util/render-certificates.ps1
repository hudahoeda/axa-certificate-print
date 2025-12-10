Param(
  [string]$BestCsv = "util/best-participant.csv",
  [string]$ParticipantCsv = "util/participant-certificate.csv",
  [string]$GeneratedCsv = "util/print-certificates.generated.csv"
)

$ErrorActionPreference = "Stop"

function Strip-Bom {
  param([string]$Line)
  return $Line -replace "^\uFEFF", ""
}

function Read-FlexibleCsv {
  param([string]$Path)

  if (-not (Test-Path $Path)) {
    Write-Warning "CSV not found, skipping: $Path"
    return @()
  }

  $lines = Get-Content -LiteralPath $Path | ForEach-Object { Strip-Bom $_ } | ForEach-Object {
    if (-not $_) { return }
    if ($_ -match '^\s*#') {
      $candidate = ($_ -replace '^\s*#\s*', '')
      if ($candidate -and $candidate -match ',') { return $candidate }
      return
    }
    $_
  }

  if (-not $lines) { return @() }

  $csvText = $lines -join "`n"
  return ($csvText | ConvertFrom-Csv)
}

function Get-FieldValue {
  param(
    [Parameter(Mandatory = $true)][psobject]$Row,
    [Parameter(Mandatory = $true)][string[]]$Names
  )

  foreach ($name in $Names) {
    if ($Row.PSObject.Properties.Name -contains $name) {
      $value = $Row.$name
      if ($value) { return $value.ToString().Trim() }
    }
  }

  return $null
}

function Slugify {
  param([string]$Text)

  if (-not $Text) { return "" }
  $slug = $Text -replace "[^A-Za-z0-9]+", "-"
  $slug = $slug.Trim("-")
  if (-not $slug) { return "entry" }
  return $slug.ToLowerInvariant()
}

function Ensure-DirForPath {
  param([string]$Path)
  $dir = Split-Path -Path $Path -Parent
  if ($dir -and (-not (Test-Path $dir))) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
  }
}

function Render-TemplateFile {
  param(
    [Parameter(Mandatory = $true)][string]$TemplatePath,
    [Parameter(Mandatory = $true)][string]$DestinationPath,
    [Parameter(Mandatory = $true)][hashtable]$Replacements
  )

  $content = Get-Content -LiteralPath $TemplatePath -Raw
  foreach ($pair in $Replacements.GetEnumerator()) {
    $content = $content.Replace($pair.Key, $pair.Value)
  }

  Ensure-DirForPath $DestinationPath
  Set-Content -LiteralPath $DestinationPath -Value $content -Encoding UTF8
}

function Build-PdfPath {
  param(
    [Parameter(Mandatory = $true)][psobject]$Row,
    [Parameter(Mandatory = $true)][string]$BaseName
  )

  $pdf = Get-FieldValue $Row @("pdf", "pdf_output")
  if ($pdf) {
    # Use directory from CSV, but derive filename from BaseName to avoid collisions.
    $dir = Split-Path -Path $pdf -Parent
    if (-not $dir) { $dir = "output" }
    return (Join-Path $dir ($BaseName + ".pdf"))
  }

  return (Join-Path "output" ($BaseName + ".pdf"))
}

$bestRows = Read-FlexibleCsv $BestCsv
$participantRows = Read-FlexibleCsv $ParticipantCsv

if (-not $bestRows -and -not $participantRows) {
  Write-Error "No rows found to render. Checked $BestCsv and $ParticipantCsv"
}

$manifest = @()

if ($bestRows) {
  foreach ($row in $bestRows) {
    $templatePath = Get-FieldValue $row @("html", "html_path", "template", "template_path")
    $name = Get-FieldValue $row @("Nama", "Name")
    $batch = Get-FieldValue $row @("Batch", "BATCH")

    if (-not $templatePath) { Write-Warning "Missing template path for best-participant row, skipping."; continue }
    if (-not $name) { Write-Warning "Missing name for best-participant row $templatePath, skipping."; continue }
    if (-not (Test-Path $templatePath)) { Write-Warning "Template not found, skipping: $templatePath"; continue }

    $resolvedTemplate = (Resolve-Path -LiteralPath $templatePath).ProviderPath
    $templateDir = Split-Path -Path $resolvedTemplate -Parent

    $slugName = Slugify $name
    $slugBatch = Slugify $batch
    $baseName = "best-participant-$slugName"
    if ($slugBatch) { $baseName = "$baseName-$slugBatch" }

    $destHtml = Join-Path $templateDir ($baseName + ".html")

    $batchValue = if ($batch) { $batch } else { "" }

    Render-TemplateFile -TemplatePath $resolvedTemplate -DestinationPath $destHtml -Replacements @{
      "{{Placeholder for Student Name}}" = $name
      "{{BATCH Placeholder}}" = $batchValue
    }

    $pdfPath = Build-PdfPath -Row $row -BaseName $baseName

    $manifest += [pscustomobject]@{
      html = $destHtml
      pdf  = $pdfPath
    }
  }
}

if ($participantRows) {
  foreach ($row in $participantRows) {
    $templatePath = Get-FieldValue $row @("html", "html_path", "template", "template_path")
    $name = Get-FieldValue $row @("Nama", "Name")

    if (-not $templatePath) { Write-Warning "Missing template path for participant row, skipping."; continue }
    if (-not $name) { Write-Warning "Missing name for participant row $templatePath, skipping."; continue }
    if (-not (Test-Path $templatePath)) { Write-Warning "Template not found, skipping: $templatePath"; continue }

    $resolvedTemplate = (Resolve-Path -LiteralPath $templatePath).ProviderPath
    $templateDir = Split-Path -Path $resolvedTemplate -Parent

    $slugName = Slugify $name
    $baseName = "participant-$slugName"
    $destHtml = Join-Path $templateDir ($baseName + ".html")

    Render-TemplateFile -TemplatePath $resolvedTemplate -DestinationPath $destHtml -Replacements @{
      "Placeholder for Student Name" = $name
    }

    $pdfPath = Build-PdfPath -Row $row -BaseName $baseName

    $manifest += [pscustomobject]@{
      html = $destHtml
      pdf  = $pdfPath
    }
  }
}

if (-not $manifest -or $manifest.Count -eq 0) {
  Write-Error "No output generated from the provided CSV files."
}

Ensure-DirForPath $GeneratedCsv
$manifest | Export-Csv -Path $GeneratedCsv -NoTypeInformation -Encoding UTF8

Write-Host "Generated $($manifest.Count) HTML file(s) with replaced placeholders."
Write-Host "Print manifest written to $GeneratedCsv"
