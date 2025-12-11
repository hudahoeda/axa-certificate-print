using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;

namespace CertificatePrinter;

internal static class CertificateEngine
{
    public static CertificateRunResult Run(CertificateEngineOptions options, IProgress<string>? progress = null)
    {
        var logger = progress ?? new Progress<string>(_ => { });

        var bestRows = ReadFlexibleCsv(options.BestCsv, "best participant", logger);
        var participantRows = ReadFlexibleCsv(options.ParticipantCsv, "participant", logger);

        if (bestRows.Count == 0 && participantRows.Count == 0)
        {
            throw new InvalidOperationException("No certificates to generate. Provide at least one CSV with rows to process.");
        }

        var manifestEntries = new List<ManifestEntry>();

        if (bestRows.Count > 0)
        {
            logger.Report($"Rendering {bestRows.Count} best-participant certificate(s)...");
            manifestEntries.AddRange(RenderBestParticipant(bestRows, options.OutputDirectoryOverride, logger));
        }

        if (participantRows.Count > 0)
        {
            logger.Report($"Rendering {participantRows.Count} participant certificate(s)...");
            manifestEntries.AddRange(RenderParticipant(participantRows, options.OutputDirectoryOverride, logger));
        }

        manifestEntries = EnsureUniquePdfPaths(manifestEntries);

        WriteManifest(manifestEntries, options.ManifestPath);
        logger.Report($"Manifest written to {Path.GetFullPath(options.ManifestPath)}");

        var browser = FindBrowser();
        if (browser == null)
        {
            throw new InvalidOperationException("No Chromium/Edge browser found. Install Microsoft Edge/Chrome or set BROWSER to the browser path.");
        }

        logger.Report($"Using browser: {browser}");
        PrintCertificates(browser, manifestEntries, options.PageRange, logger);

        return new CertificateRunResult(manifestEntries.Count, manifestEntries);
    }

    private static List<Dictionary<string, string>> ReadFlexibleCsv(string? path, string label, IProgress<string> logger)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            logger.Report($"[{label}] No path provided, skipping.");
            return new List<Dictionary<string, string>>();
        }

        if (!File.Exists(path))
        {
            logger.Report($"[{label}] CSV not found, skipping: {path}");
            return new List<Dictionary<string, string>>();
        }

        var filteredLines = new List<string>();
        foreach (var rawLine in File.ReadLines(path, Encoding.UTF8))
        {
            var line = StripBom(rawLine);
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            if (IsComment(line))
            {
                var candidate = RemoveCommentMarker(line);
                if (!string.IsNullOrWhiteSpace(candidate) && candidate.Contains(','))
                {
                    filteredLines.Add(candidate);
                }

                continue;
            }

            filteredLines.Add(line);
        }

        if (filteredLines.Count == 0)
        {
            logger.Report($"[{label}] CSV is empty after filtering comments: {path}");
            return new List<Dictionary<string, string>>();
        }

        var csvText = string.Join(Environment.NewLine, filteredLines);
        using var reader = new StringReader(csvText);
        using var parser = new TextFieldParser(reader)
        {
            Delimiters = new[] { "," },
            HasFieldsEnclosedInQuotes = true,
            TrimWhiteSpace = true
        };

        var rows = new List<Dictionary<string, string>>();
        if (parser.EndOfData)
        {
            return rows;
        }

        var headers = parser.ReadFields()?.Select(h => h?.Trim() ?? string.Empty).ToArray() ?? Array.Empty<string>();
        while (!parser.EndOfData)
        {
            var fields = parser.ReadFields();
            if (fields == null)
            {
                continue;
            }

            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < headers.Length && i < fields.Length; i++)
            {
                var header = headers[i];
                if (string.IsNullOrWhiteSpace(header))
                {
                    continue;
                }

                row[header] = fields[i]?.Trim() ?? string.Empty;
            }

            rows.Add(row);
        }

        return rows;
    }

    private static IEnumerable<ManifestEntry> RenderBestParticipant(IEnumerable<Dictionary<string, string>> rows, string? outputOverride, IProgress<string> logger)
    {
        var manifest = new List<ManifestEntry>();

        foreach (var row in rows)
        {
            var templatePath = GetField(row, "html", "html_path", "template", "template_path");
            var name = GetField(row, "Nama", "Name");
            var batch = GetField(row, "Batch", "BATCH") ?? string.Empty;

            if (string.IsNullOrWhiteSpace(templatePath))
            {
                logger.Report("Missing template path for best-participant row, skipping.");
                continue;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                logger.Report($"Missing name for best-participant row {templatePath}, skipping.");
                continue;
            }

            if (!File.Exists(templatePath))
            {
                logger.Report($"Template not found, skipping: {templatePath}");
                continue;
            }

            var resolvedTemplate = Path.GetFullPath(templatePath);
            var templateDir = Path.GetDirectoryName(resolvedTemplate) ?? Environment.CurrentDirectory;

            var slugName = Slugify(name);
            var slugBatch = Slugify(batch);
            var baseName = string.IsNullOrEmpty(slugBatch)
                ? $"best-participant-{slugName}"
                : $"best-participant-{slugName}-{slugBatch}";

            var destHtml = Path.Combine(templateDir, $"{baseName}.html");
            var batchValue = batch ?? string.Empty;

            RenderTemplate(resolvedTemplate, destHtml, new Dictionary<string, string>
            {
                ["{{Placeholder for Student Name}}"] = name,
                ["{{BATCH Placeholder}}"] = batchValue
            });

            var pdfPath = BuildPdfPath(row, baseName, outputOverride);
            manifest.Add(new ManifestEntry(Path.GetFullPath(destHtml), pdfPath));
        }

        return manifest;
    }

    private static IEnumerable<ManifestEntry> RenderParticipant(IEnumerable<Dictionary<string, string>> rows, string? outputOverride, IProgress<string> logger)
    {
        var manifest = new List<ManifestEntry>();

        foreach (var row in rows)
        {
            var templatePath = GetField(row, "html", "html_path", "template", "template_path");
            var name = GetField(row, "Nama", "Name");

            if (string.IsNullOrWhiteSpace(templatePath))
            {
                logger.Report("Missing template path for participant row, skipping.");
                continue;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                logger.Report($"Missing name for participant row {templatePath}, skipping.");
                continue;
            }

            if (!File.Exists(templatePath))
            {
                logger.Report($"Template not found, skipping: {templatePath}");
                continue;
            }

            var resolvedTemplate = Path.GetFullPath(templatePath);
            var templateDir = Path.GetDirectoryName(resolvedTemplate) ?? Environment.CurrentDirectory;
            var baseName = $"participant-{Slugify(name)}";
            var destHtml = Path.Combine(templateDir, $"{baseName}.html");

            RenderTemplate(resolvedTemplate, destHtml, new Dictionary<string, string>
            {
                ["Placeholder for Student Name"] = name
            });

            var pdfPath = BuildPdfPath(row, baseName, outputOverride);
            manifest.Add(new ManifestEntry(Path.GetFullPath(destHtml), pdfPath));
        }

        return manifest;
    }

    private static void RenderTemplate(string templatePath, string destinationPath, Dictionary<string, string> replacements)
    {
        var content = File.ReadAllText(templatePath, Encoding.UTF8);
        foreach (var pair in replacements)
        {
            content = content.Replace(pair.Key, pair.Value, StringComparison.Ordinal);
        }

        EnsureDirForFile(destinationPath);
        File.WriteAllText(destinationPath, content, Encoding.UTF8);
    }

    private static string BuildPdfPath(IReadOnlyDictionary<string, string> row, string baseName, string? outputOverride)
    {
        if (!string.IsNullOrWhiteSpace(outputOverride))
        {
            return Path.GetFullPath(Path.Combine(outputOverride, $"{baseName}.pdf"));
        }

        var pdfField = GetField(row, "pdf", "pdf_output");
        if (!string.IsNullOrWhiteSpace(pdfField))
        {
            var dir = Path.GetDirectoryName(pdfField);
            dir = string.IsNullOrWhiteSpace(dir) ? "output" : dir;
            return Path.GetFullPath(Path.Combine(dir, $"{baseName}.pdf"));
        }

        return Path.GetFullPath(Path.Combine("output", $"{baseName}.pdf"));
    }

    private static void EnsureDirForFile(string path)
    {
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }
    }

    private static void WriteManifest(IEnumerable<ManifestEntry> manifest, string path)
    {
        EnsureDirForFile(path);
        var sb = new StringBuilder();
        sb.AppendLine("html,pdf");
        foreach (var entry in manifest)
        {
            sb.AppendLine($"\"{entry.HtmlPath}\",\"{entry.PdfPath}\"");
        }

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
    }

    private static string? FindBrowser()
    {
        if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("BROWSER")))
        {
            var candidate = Environment.GetEnvironmentVariable("BROWSER")!;
            var pathFromEnv = ResolveExecutable(candidate);
            if (pathFromEnv != null)
            {
                return pathFromEnv;
            }
        }

        var edgeCandidates = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft", "Edge", "Application", "msedge.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft", "Edge", "Application", "msedge.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Edge", "Application", "msedge.exe")
        };

        foreach (var edgePath in edgeCandidates)
        {
            if (!string.IsNullOrWhiteSpace(edgePath) && File.Exists(edgePath))
            {
                return edgePath;
            }
        }

        var otherCandidates = new[] { "msedge", "msedge.exe", "chrome", "chrome.exe", "google-chrome", "chromium", "chromium.exe" };
        foreach (var candidate in otherCandidates)
        {
            var resolved = ResolveExecutable(candidate);
            if (resolved != null)
            {
                return resolved;
            }
        }

        return null;
    }

    private static string? ResolveExecutable(string candidate)
    {
        if (File.Exists(candidate))
        {
            return candidate;
        }

        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var dir in pathEnv.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            var full = Path.Combine(dir, candidate);
            if (File.Exists(full))
            {
                return full;
            }
        }

        return null;
    }

    private static void PrintCertificates(string browserPath, IEnumerable<ManifestEntry> manifest, string pageRange, IProgress<string> logger)
    {
        var list = manifest.ToList();
        for (var i = 0; i < list.Count; i++)
        {
            var entry = list[i];
            if (!File.Exists(entry.HtmlPath))
            {
                logger.Report($"[{i + 1}/{list.Count}] HTML not found, skipping: {entry.HtmlPath}");
                continue;
            }

            EnsureDirForFile(entry.PdfPath);
            var fileUri = new Uri(entry.HtmlPath).AbsoluteUri;

            var psi = new ProcessStartInfo
            {
                FileName = browserPath,
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true
            };

            psi.ArgumentList.Add("--headless");
            psi.ArgumentList.Add("--disable-gpu");
            psi.ArgumentList.Add("--no-sandbox");
            psi.ArgumentList.Add($"--print-to-pdf={entry.PdfPath}");

            if (!string.IsNullOrWhiteSpace(pageRange))
            {
                psi.ArgumentList.Add($"--print-to-pdf-page-range={pageRange}");
            }

            psi.ArgumentList.Add(fileUri);

            logger.Report($"[{i + 1}/{list.Count}] Printing {entry.HtmlPath} -> {entry.PdfPath}");
            using var process = Process.Start(psi);
            if (process == null)
            {
                logger.Report("Failed to start browser process.");
                continue;
            }

            var output = process.StandardOutput.ReadToEnd();
            var error = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException($"Browser exited with code {process.ExitCode} for {entry.HtmlPath}. Output:\n{output}\n{error}");
            }
        }
    }

    private static List<ManifestEntry> EnsureUniquePdfPaths(IEnumerable<ManifestEntry> entries)
    {
        var result = new List<ManifestEntry>();
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in entries)
        {
            var pdfPath = entry.PdfPath;
            if (counts.TryGetValue(pdfPath, out var count))
            {
                count += 1;
                counts[pdfPath] = count;

                var dir = Path.GetDirectoryName(pdfPath) ?? "";
                var fileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfPath);
                var ext = Path.GetExtension(pdfPath);
                var newFileName = $"{fileNameWithoutExt}-{count}{ext}";
                var newPath = Path.GetFullPath(Path.Combine(dir, newFileName));
                result.Add(entry with { PdfPath = newPath });
            }
            else
            {
                counts[pdfPath] = 0;
                result.Add(entry);
            }
        }

        return result;
    }

    private static string? GetField(IReadOnlyDictionary<string, string> row, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (row.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return null;
    }

    private static string Slugify(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return "entry";
        }

        var slug = Regex.Replace(text, "[^A-Za-z0-9]+", "-");
        slug = slug.Trim('-');
        return string.IsNullOrWhiteSpace(slug) ? "entry" : slug.ToLowerInvariant();
    }

    private static string StripBom(string line)
    {
        if (string.IsNullOrEmpty(line))
        {
            return string.Empty;
        }

        return line.TrimStart('\uFEFF');
    }

    private static bool IsComment(string line)
    {
        return line.TrimStart().StartsWith("#");
    }

    private static string RemoveCommentMarker(string line)
    {
        var trimmed = line.TrimStart();
        if (trimmed.StartsWith("#"))
        {
            trimmed = trimmed.TrimStart('#').TrimStart();
        }

        return trimmed;
    }
}

internal record CertificateRunResult(int GeneratedCount, List<ManifestEntry> Manifest);

internal record ManifestEntry(string HtmlPath, string PdfPath);

internal record CertificateEngineOptions(
    string? BestCsv,
    string? ParticipantCsv,
    string ManifestPath,
    string PageRange,
    string? OutputDirectoryOverride
);
