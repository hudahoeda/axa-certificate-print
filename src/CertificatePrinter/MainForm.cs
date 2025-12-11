using System.Threading.Tasks;
using System.Windows.Forms;

namespace CertificatePrinter;

public class MainForm : Form
{
    private readonly TextBox _bestCsv;
    private readonly TextBox _participantCsv;
    private readonly TextBox _outputDir;
    private readonly TextBox _manifestPath;
    private readonly TextBox _pageRange;
    private readonly TextBox _log;
    private readonly Button _runButton;

    public MainForm()
    {
        Text = "Certificate Printer";
        Width = 760;
        Height = 540;
        MinimumSize = new System.Drawing.Size(720, 480);

        var padding = 12;
        var labelWidth = 150;
        var inputWidth = 400;
        var buttonWidth = 80;
        var rowHeight = 28;
        var top = padding;

        Controls.Add(CreateLabel("Best participant CSV", padding, top, labelWidth));
        _bestCsv = CreateTextBox(ProgramDefaults.DefaultBestCsv, padding + labelWidth, top, inputWidth);
        Controls.Add(_bestCsv);
        var browseBest = CreateButton("Browse", padding + labelWidth + inputWidth + 8, top, buttonWidth, (_, _) => BrowseCsv(_bestCsv));
        Controls.Add(browseBest);
        top += rowHeight + 6;

        Controls.Add(CreateLabel("Participant CSV", padding, top, labelWidth));
        _participantCsv = CreateTextBox(ProgramDefaults.DefaultParticipantCsv, padding + labelWidth, top, inputWidth);
        Controls.Add(_participantCsv);
        var browseParticipant = CreateButton("Browse", padding + labelWidth + inputWidth + 8, top, buttonWidth, (_, _) => BrowseCsv(_participantCsv));
        Controls.Add(browseParticipant);
        top += rowHeight + 6;

        Controls.Add(CreateLabel("Output folder (optional)", padding, top, labelWidth));
        _outputDir = CreateTextBox(string.Empty, padding + labelWidth, top, inputWidth);
        Controls.Add(_outputDir);
        var browseOutput = CreateButton("Browse", padding + labelWidth + inputWidth + 8, top, buttonWidth, (_, _) => BrowseFolder(_outputDir));
        Controls.Add(browseOutput);
        top += rowHeight + 6;

        Controls.Add(CreateLabel("Manifest path", padding, top, labelWidth));
        _manifestPath = CreateTextBox(ProgramDefaults.DefaultManifestPath, padding + labelWidth, top, inputWidth);
        Controls.Add(_manifestPath);
        var browseManifest = CreateButton("Browse", padding + labelWidth + inputWidth + 8, top, buttonWidth, (_, _) => BrowseManifest(_manifestPath));
        Controls.Add(browseManifest);
        top += rowHeight + 6;

        Controls.Add(CreateLabel("Page range", padding, top, labelWidth));
        _pageRange = CreateTextBox(ProgramDefaults.DefaultPageRange, padding + labelWidth, top, 120);
        Controls.Add(_pageRange);
        top += rowHeight + 10;

        _runButton = CreateButton("Generate & Print", padding, top, 160, RunButton_Click);
        Controls.Add(_runButton);

        var clearLogButton = CreateButton("Clear log", padding + 170, top, 100, (_, _) => _log.Clear());
        Controls.Add(clearLogButton);
        top += rowHeight + 10;

        _log = new TextBox
        {
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            ReadOnly = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            Left = padding,
            Top = top,
            Width = ClientSize.Width - padding * 2,
            Height = ClientSize.Height - top - padding
        };
        Controls.Add(_log);

        Resize += (_, _) =>
        {
            _log.Width = ClientSize.Width - padding * 2;
            _log.Height = ClientSize.Height - _log.Top - padding;
        };
    }

    private static Label CreateLabel(string text, int left, int top, int width)
    {
        return new Label
        {
            Text = text,
            Left = left,
            Top = top + 6,
            Width = width,
            AutoSize = false
        };
    }

    private static TextBox CreateTextBox(string text, int left, int top, int width)
    {
        return new TextBox
        {
            Text = text,
            Left = left,
            Top = top,
            Width = width,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
    }

    private static Button CreateButton(string text, int left, int top, int width, EventHandler onClick)
    {
        var btn = new Button
        {
            Text = text,
            Left = left,
            Top = top,
            Width = width
        };
        btn.Click += onClick;
        return btn;
    }

    private void BrowseCsv(TextBox target)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            target.Text = dialog.FileName;
        }
    }

    private void BrowseFolder(TextBox target)
    {
        using var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog() == DialogResult.OK)
        {
            target.Text = dialog.SelectedPath;
        }
    }

    private void BrowseManifest(TextBox target)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = target.Text
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            target.Text = dialog.FileName;
        }
    }

    private async void RunButton_Click(object? sender, EventArgs e)
    {
        var bestCsv = _bestCsv.Text.Trim();
        var participantCsv = _participantCsv.Text.Trim();
        var manifestPath = _manifestPath.Text.Trim();
        var pageRange = _pageRange.Text.Trim();
        var outputDir = string.IsNullOrWhiteSpace(_outputDir.Text) ? null : _outputDir.Text.Trim();

        if (string.IsNullOrWhiteSpace(bestCsv) && string.IsNullOrWhiteSpace(participantCsv))
        {
            MessageBox.Show("Please provide at least one CSV path.", "Missing input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _runButton.Enabled = false;
        AppendLog("Starting...");

        try
        {
            var options = new CertificateEngineOptions(bestCsv, participantCsv, manifestPath, pageRange, outputDir);
            var progress = new Progress<string>(AppendLog);

            await Task.Run(() => CertificateEngine.Run(options, progress));

            AppendLog("Completed.");
            MessageBox.Show("Certificates generated and printed.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"Error: {ex.Message}");
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _runButton.Enabled = true;
        }
    }

    private void AppendLog(string message)
    {
        if (InvokeRequired)
        {
            Invoke(new Action<string>(AppendLog), message);
            return;
        }

        var timestamp = DateTime.Now.ToString("HH:mm:ss");
        _log.AppendText($"[{timestamp}] {message}{Environment.NewLine}");
    }
}

internal static class ProgramDefaults
{
    public const string DefaultBestCsv = "util/best-participant.csv";
    public const string DefaultParticipantCsv = "util/participant-certificate.csv";
    public const string DefaultManifestPath = "util/print-certificates.generated.csv";
    public const string DefaultPageRange = "1";
}
