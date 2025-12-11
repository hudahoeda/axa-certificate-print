using System;
using System.Linq;
using System.Windows.Forms;

namespace CertificatePrinter;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        // If any CLI-like argument is present, run in CLI mode for backward compatibility.
        if (args.Length > 0 && args.Any(a => a.StartsWith("-")))
        {
            return RunCli(args);
        }

        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
        return 0;
    }

    private static int RunCli(string[] args)
    {
        try
        {
            var options = ParseArgs(args);
            if (options.ShowHelp)
            {
                PrintHelp();
                return 0;
            }

            var bestCsv = options.BestCsv ?? ProgramDefaults.DefaultBestCsv;
            var participantCsv = options.ParticipantCsv ?? ProgramDefaults.DefaultParticipantCsv;
            var manifestPath = options.ManifestPath ?? ProgramDefaults.DefaultManifestPath;
            var pageRange = options.PageRange ?? ProgramDefaults.DefaultPageRange;
            var outputDir = options.OutputDir;

            var engineOptions = new CertificateEngineOptions(bestCsv, participantCsv, manifestPath, pageRange, outputDir);
            var result = CertificateEngine.Run(engineOptions, new Progress<string>(Console.WriteLine));

            Console.WriteLine($"Generated and printed {result.GeneratedCount} certificate(s).");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    private static CliOptions ParseArgs(string[] args)
    {
        var result = new CliOptions();

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            switch (arg.ToLowerInvariant())
            {
                case "--best":
                case "-b":
                    result.BestCsv = ReadNextArg(args, ref i, "best CSV path");
                    break;
                case "--participant":
                case "-p":
                    result.ParticipantCsv = ReadNextArg(args, ref i, "participant CSV path");
                    break;
                case "--manifest":
                case "-m":
                    result.ManifestPath = ReadNextArg(args, ref i, "manifest output path");
                    break;
                case "--page-range":
                case "-r":
                    result.PageRange = ReadNextArg(args, ref i, "page range");
                    break;
                case "--output":
                case "-o":
                    result.OutputDir = ReadNextArg(args, ref i, "output directory (overrides CSV pdf paths)");
                    break;
                case "--help":
                case "-h":
                case "/?":
                    result.ShowHelp = true;
                    break;
                default:
                    Console.Error.WriteLine($"Unknown argument: {arg}");
                    result.ShowHelp = true;
                    break;
            }
        }

        return result;
    }

    private static string ReadNextArg(string[] args, ref int index, string description)
    {
        if (index + 1 >= args.Length)
        {
            throw new ArgumentException($"Missing value after {args[index]} ({description}).");
        }

        index += 1;
        return args[index];
    }

    private static void PrintHelp()
    {
        Console.WriteLine("AXA Certificate Printer");
        Console.WriteLine();
        Console.WriteLine("Usage (CLI mode):");
        Console.WriteLine("  CertificatePrinter.exe --best <path> --participant <path> --manifest <path> --page-range <range> --output <folder>");
        Console.WriteLine();
        Console.WriteLine("Defaults when omitted:");
        Console.WriteLine($"  Best participant CSV : {ProgramDefaults.DefaultBestCsv}");
        Console.WriteLine($"  Participant CSV      : {ProgramDefaults.DefaultParticipantCsv}");
        Console.WriteLine($"  Manifest output      : {ProgramDefaults.DefaultManifestPath}");
        Console.WriteLine($"  Page range           : {ProgramDefaults.DefaultPageRange} (empty string prints all pages)");
        Console.WriteLine($"  Output folder        : use pdf/pdf_output from CSV unless overridden with --output");
        Console.WriteLine();
        Console.WriteLine("Run without arguments to open the GUI.");
    }

    private class CliOptions
    {
        public string? BestCsv { get; set; }
        public string? ParticipantCsv { get; set; }
        public string? ManifestPath { get; set; }
        public string? PageRange { get; set; }
        public string? OutputDir { get; set; }
        public bool ShowHelp { get; set; }
    }
}
