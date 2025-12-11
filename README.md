# AXA Certificate Printer

Simple Windows app (GUI + optional CLI) that renders the AXA certificate templates and prints them to PDF using Edge/Chromium. It replaces the placeholders inside the HTML templates with the values from the CSV files and prints them via the browser's `--print-to-pdf` option.

## Build the executable

Requires the .NET 6+ SDK on Windows.

```
pwsh util/build-exe.ps1
# or: dotnet publish src/CertificatePrinter -c Release -r win-x64 --self-contained false /p:PublishSingleFile=true -o dist
```

The executable is emitted to `dist/CertificatePrinter.exe`.

## Run it (GUI)

1) Ensure Microsoft Edge or Chrome/Chromium is installed (or set `BROWSER` to its path).
2) Double-click `dist/CertificatePrinter.exe`.
3) Pick the CSV files (best and/or participant), optionally choose an output folder (overrides CSV `pdf_output` paths), adjust page range/manifest path if desired, then click “Generate & Print”.
4) The app creates filled-in HTML copies next to the templates, writes the manifest to `util/print-certificates.generated.csv` (or your chosen path), and prints PDFs to the CSV-provided paths or the override folder.

## Run it (CLI, optional)

```
CertificatePrinter.exe --best util/best-participant.csv --participant util/participant-certificate.csv --manifest util/print-certificates.generated.csv --page-range 1 --output output
```
