Param(
  [string]$Runtime = "win-x64",
  [switch]$SelfContained
)

$ErrorActionPreference = "Stop"

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$project = Join-Path $root "src/CertificatePrinter/CertificatePrinter.csproj"
$outputDir = Join-Path $root "dist"

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
  Write-Error "dotnet SDK not found. Install .NET 6+ SDK to build the executable."
}

$selfContainedValue = if ($SelfContained.IsPresent) { "true" } else { "false" }

$publishArgs = @(
  "publish", $project,
  "-c", "Release",
  "-r", $Runtime,
  "--self-contained", $selfContainedValue,
  "-o", $outputDir,
  "/p:PublishSingleFile=true",
  "/p:EnableCompressionInSingleFile=true"
)

Write-Host "Building CertificatePrinter for $Runtime (self-contained: $($SelfContained.IsPresent))..."
dotnet @publishArgs

$exePath = Join-Path $outputDir "CertificatePrinter.exe"
if (Test-Path $exePath) {
  Write-Host "Executable ready: $exePath"
} else {
  Write-Warning "Publish finished but executable not found in $outputDir. Check dotnet output above."
}
