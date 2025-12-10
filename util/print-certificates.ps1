Param(
  [string[]]$CsvPaths = @(
    "util/participant-certificate.csv",
    "util/best-participant.csv"
  ),
  [string]$PageRange = "1" # Limit to first page by default; set to "" to disable.
)

$ErrorActionPreference = "Stop"

function Get-BrowserPath {
  # Prefer Microsoft Edge, then fall back to other Chromium-based browsers
  if ($env:BROWSER) {
    $cmd = Get-Command $env:BROWSER -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
  }

  foreach ($edgePath in @(
    "$Env:ProgramFiles\\Microsoft\\Edge\\Application\\msedge.exe",
    "$Env:ProgramFiles (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    "$Env:LocalAppData\\Microsoft\\Edge\\Application\\msedge.exe"
  )) {
    if ($edgePath -and (Test-Path $edgePath)) { return $edgePath }
  }

  foreach ($candidate in @("msedge", "msedge.exe", "chrome", "chrome.exe", "google-chrome", "chromium", "chromium.exe")) {
    $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
  }

  return $null
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

function StripBom {
  param([string]$Line)
  # Remove any leading UTF-8 BOM
  return $Line -replace "^\uFEFF", ""
}

$browser = Get-BrowserPath
if (-not $browser) {
  Write-Error "No headless Microsoft Edge/Chromium found. Install Edge or set BROWSER=<path>."
}

$allRows = @()
$missingCsv = @()

foreach ($CsvPath in $CsvPaths) {
  if (-not (Test-Path $CsvPath)) {
    $missingCsv += $CsvPath
    continue
  }

  # Read CSV while allowing comment headers (# html_path,...) and ignoring empty lines.
  $lines = Get-Content $CsvPath | ForEach-Object { StripBom $_ } | ForEach-Object {
    if (-not $_) { return }
    if ($_ -match '^\s*#') {
      $maybeHeader = ($_ -replace '^\s*#\s*', '')
      if ($maybeHeader -and $maybeHeader -match ',') { return $maybeHeader }
      return
    }
    $_
  }

  if (-not $lines) {
    Write-Warning "CSV is empty after filtering comments: $CsvPath"
    continue
  }

  $csvText = $lines -join "`n"
  $allRows += ($csvText | ConvertFrom-Csv)
}

if ($missingCsv.Count -eq $CsvPaths.Count) {
  Write-Error ("CSV file(s) not found: " + ($missingCsv -join ", "))
}

if (-not $allRows -or $allRows.Count -eq 0) {
  Write-Error ("No printable rows found in: " + ($CsvPaths -join ", "))
}

foreach ($row in $allRows) {
  $htmlPath = Get-FieldValue $row @("html", "html_path")
  $pdfPath = Get-FieldValue $row @("pdf", "pdf_output")

  if (-not $htmlPath) { continue }
  if (-not $pdfPath) {
    Write-Warning "PDF output path missing, skipping row linked to $htmlPath"
    continue
  }

  $absHtml = Resolve-Path $htmlPath -ErrorAction SilentlyContinue
  if (-not $absHtml) {
    Write-Warning "HTML not found, skipping: $htmlPath"
    continue
  }

  $absHtml = $absHtml.ProviderPath

  # Build absolute PDF path and ensure directory exists
  if ([System.IO.Path]::IsPathRooted($pdfPath)) {
    $absPdf = $pdfPath
  } else {
    $absPdf = Join-Path (Get-Location) $pdfPath
  }

  $pdfDir = Split-Path $absPdf
  if (-not (Test-Path $pdfDir)) {
    New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
  }

  $fileUri = [Uri]::new($absHtml).AbsoluteUri
  $args = @(
    "--headless",
    "--disable-gpu",
    "--no-sandbox",
    "--print-to-pdf=$absPdf",
    $fileUri
  )

  if ($PageRange -and $PageRange.Trim()) {
    $args += "--print-to-pdf-page-range=$PageRange"
  }

  Write-Host "Printing $absHtml -> $absPdf"
  & $browser @args
}
