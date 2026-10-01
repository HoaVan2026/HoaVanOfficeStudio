using System.Diagnostics;
using System.Text;

namespace HoaVan.OfficeStudio.Services;

public sealed class OfficeCliService
{
    private readonly string _enginePath;

    public OfficeCliService(string? explicitEnginePath = null)
    {
        _enginePath = ResolveEnginePath(explicitEnginePath);
    }

    public string EnginePath => _enginePath;

    public bool IsEngineAvailable => File.Exists(_enginePath) || _enginePath.Equals("officecli", StringComparison.OrdinalIgnoreCase) || _enginePath.Equals("officecli.exe", StringComparison.OrdinalIgnoreCase);

    public async Task<OfficeCliResult> RunAsync(IEnumerable<string> arguments, CancellationToken cancellationToken = default)
    {
        var psi = new ProcessStartInfo
        {
            FileName = _enginePath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = Environment.CurrentDirectory
        };

        foreach (var arg in arguments)
            psi.ArgumentList.Add(arg);

        using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };
        var stdout = new StringBuilder();
        var stderr = new StringBuilder();
        process.OutputDataReceived += (_, e) => { if (e.Data is not null) stdout.AppendLine(e.Data); };
        process.ErrorDataReceived += (_, e) => { if (e.Data is not null) stderr.AppendLine(e.Data); };

        try
        {
            if (!process.Start())
                return new OfficeCliResult(-1, string.Empty, "Không thể khởi chạy OfficeCLI.");
        }
        catch (Exception ex)
        {
            return new OfficeCliResult(-1, string.Empty, $"Không thể khởi chạy OfficeCLI: {ex.Message}");
        }

        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        try
        {
            await process.WaitForExitAsync(cancellationToken);
            return new OfficeCliResult(process.ExitCode, stdout.ToString(), stderr.ToString());
        }
        catch (OperationCanceledException)
        {
            try { if (!process.HasExited) process.Kill(entireProcessTree: true); } catch { }
            throw;
        }
    }

    public Task<OfficeCliResult> CreateAsync(string filePath, string locale = "vi-VN", bool force = false, CancellationToken ct = default)
    {
        var args = new List<string> { "create", filePath, "--locale", locale };
        if (force) args.Add("--force");
        return RunAsync(args, ct);
    }

    public Task<OfficeCliResult> OutlineAsync(string filePath, CancellationToken ct = default) =>
        RunAsync(new[] { "view", filePath, "outline" }, ct);

    public Task<OfficeCliResult> IssuesAsync(string filePath, CancellationToken ct = default) =>
        RunAsync(new[] { "view", filePath, "issues" }, ct);

    public Task<OfficeCliResult> PreviewHtmlAsync(string filePath, CancellationToken ct = default) =>
        RunAsync(new[] { "view", filePath, "html", "--browser" }, ct);

    public Task<OfficeCliResult> ExportPdfAsync(string filePath, string outputPath, CancellationToken ct = default) =>
        RunAsync(new[] { "view", filePath, "pdf", "--out", outputPath }, ct);

    public Task<OfficeCliResult> CloseAsync(string filePath, CancellationToken ct = default) =>
        RunAsync(new[] { "close", filePath }, ct);

    public Task<OfficeCliResult> VersionAsync(CancellationToken ct = default) =>
        RunAsync(new[] { "--version" }, ct);

    private static string ResolveEnginePath(string? explicitPath)
    {
        if (!string.IsNullOrWhiteSpace(explicitPath)) return explicitPath;
        var envPath = Environment.GetEnvironmentVariable("OFFICECLI_PATH");
        if (!string.IsNullOrWhiteSpace(envPath)) return envPath;

        var baseDir = AppContext.BaseDirectory;
        var bundled = Path.Combine(baseDir, "tools", "officecli-win-x64.exe");
        if (File.Exists(bundled)) return bundled;

        var adjacent = Path.Combine(baseDir, "officecli-win-x64.exe");
        if (File.Exists(adjacent)) return adjacent;

        return "officecli.exe";
    }
}
