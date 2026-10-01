using System.IO;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using HoaVan.OfficeStudio.Models;
using HoaVan.OfficeStudio.Services;
using HoaVan.OfficeStudio.Utilities;

namespace HoaVan.OfficeStudio;

public partial class MainWindow : Window
{
    private readonly OfficeCliService _officeCli = new();
    private string? _currentFile;
    private CancellationTokenSource? _cts;

    public MainWindow()
    {
        InitializeComponent();
        Loaded += MainWindow_Loaded;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        AppendLog($"Engine path: {_officeCli.EnginePath}");
        var result = await _officeCli.VersionAsync();
        if (result.Success)
        {
            EngineStatusText.Text = $"Sẵn sàng · {result.StandardOutput.Trim()}";
            EngineStatusText.Foreground = System.Windows.Media.Brushes.LightGreen;
            AppendResult("OfficeCLI version", result);
        }
        else
        {
            EngineStatusText.Text = "Chưa tìm thấy OfficeCLI. Hãy đặt officecli-win-x64.exe vào thư mục tools của app hoặc đặt biến OFFICECLI_PATH.";
            EngineStatusText.Foreground = System.Windows.Media.Brushes.LightSalmon;
            AppendResult("Kiểm tra engine", result);
        }
    }

    private OfficeDocumentType SelectedType()
    {
        var item = (ComboBoxItem)DocumentTypeCombo.SelectedItem;
        return Enum.Parse<OfficeDocumentType>(item.Tag?.ToString() ?? "Word");
    }

    private async void CreateDocument_Click(object sender, RoutedEventArgs e)
    {
        var type = SelectedType();
        var cleanName = PathHelper.SanitizeFileName(NewDocumentNameText.Text);
        var dialog = new SaveFileDialog
        {
            FileName = cleanName + type.Extension(),
            Filter = type switch
            {
                OfficeDocumentType.Word => "Word Document (*.docx)|*.docx",
                OfficeDocumentType.PowerPoint => "PowerPoint Presentation (*.pptx)|*.pptx",
                OfficeDocumentType.Excel => "Excel Workbook (*.xlsx)|*.xlsx",
                _ => "All files (*.*)|*.*"
            },
            AddExtension = true,
            DefaultExt = type.Extension()
        };
        if (dialog.ShowDialog() != true) return;

        await RunBusyAsync("Tạo tài liệu", async ct =>
        {
            var result = await _officeCli.CreateAsync(dialog.FileName, "vi-VN", force: false, ct);
            AppendResult("create", result);
            if (result.Success) SetCurrentFile(dialog.FileName);
        });
    }

    private void SelectFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Office files (*.docx;*.pptx;*.xlsx)|*.docx;*.pptx;*.xlsx|All files (*.*)|*.*"
        };
        if (dialog.ShowDialog() == true) SetCurrentFile(dialog.FileName);
    }

    private async void Outline_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureCurrentFile()) return;
        await RunBusyAsync("Đọc outline", async ct => AppendResult("outline", await _officeCli.OutlineAsync(_currentFile!, ct)));
    }

    private async void Issues_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureCurrentFile()) return;
        await RunBusyAsync("Kiểm tra issues", async ct => AppendResult("issues", await _officeCli.IssuesAsync(_currentFile!, ct)));
    }

    private async void Preview_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureCurrentFile()) return;
        await RunBusyAsync("Mở preview HTML", async ct => AppendResult("preview", await _officeCli.PreviewHtmlAsync(_currentFile!, ct)));
    }

    private async void ExportPdf_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureCurrentFile()) return;
        var dialog = new SaveFileDialog
        {
            FileName = Path.GetFileNameWithoutExtension(_currentFile) + ".pdf",
            Filter = "PDF (*.pdf)|*.pdf",
            DefaultExt = ".pdf"
        };
        if (dialog.ShowDialog() != true) return;
        await RunBusyAsync("Xuất PDF", async ct => AppendResult("export pdf", await _officeCli.ExportPdfAsync(_currentFile!, dialog.FileName, ct)));
    }

    private void SaveCopy_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureCurrentFile()) return;
        var ext = Path.GetExtension(_currentFile);
        var dialog = new SaveFileDialog
        {
            FileName = Path.GetFileNameWithoutExtension(_currentFile) + "_copy" + ext,
            Filter = $"Office file (*{ext})|*{ext}",
            DefaultExt = ext
        };
        if (dialog.ShowDialog() != true) return;
        try
        {
            File.Copy(_currentFile!, dialog.FileName, overwrite: true);
            AppendLog($"[OK] Đã lưu bản sao: {dialog.FileName}");
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Không thể lưu bản sao: {ex.Message}");
        }
    }

    private async void AdvancedCommand_Click(object sender, RoutedEventArgs e)
    {
        var raw = AdvancedCommandText.Text.Trim();
        if (string.IsNullOrWhiteSpace(raw)) return;
        if (raw.Contains("{file}", StringComparison.OrdinalIgnoreCase))
        {
            if (!EnsureCurrentFile()) return;
            raw = raw.Replace("{file}", _currentFile!, StringComparison.OrdinalIgnoreCase);
        }

        var args = SplitCommandLine(raw);
        if (args.Count == 0) return;
        await RunBusyAsync("Advanced command", async ct => AppendResult(raw, await _officeCli.RunAsync(args, ct)));
    }

    private void ClearLog_Click(object sender, RoutedEventArgs e) => LogText.Clear();

    private bool EnsureCurrentFile()
    {
        if (!string.IsNullOrWhiteSpace(_currentFile) && File.Exists(_currentFile)) return true;
        MessageBox.Show("Hãy chọn một file .docx, .pptx hoặc .xlsx trước.", "Hoa Van Office Studio", MessageBoxButton.OK, MessageBoxImage.Information);
        return false;
    }

    private void SetCurrentFile(string path)
    {
        _currentFile = Path.GetFullPath(path);
        CurrentFileText.Text = _currentFile;
        AppendLog($"[FILE] {_currentFile}");
    }

    private async Task RunBusyAsync(string label, Func<CancellationToken, Task> action)
    {
        _cts?.Cancel();
        _cts = new CancellationTokenSource();
        IsEnabled = false;
        AppendLog($"\n>>> {label}");
        try
        {
            await action(_cts.Token);
        }
        catch (OperationCanceledException)
        {
            AppendLog("[CANCELLED] Tác vụ đã bị hủy.");
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] {ex.Message}");
        }
        finally
        {
            IsEnabled = true;
        }
    }

    private void AppendResult(string label, OfficeCliResult result)
    {
        AppendLog($"[{(result.Success ? "OK" : "FAIL")}] {label} · exit={result.ExitCode}");
        if (!string.IsNullOrWhiteSpace(result.StandardOutput)) AppendLog(result.StandardOutput.TrimEnd());
        if (!string.IsNullOrWhiteSpace(result.StandardError)) AppendLog("stderr: " + result.StandardError.TrimEnd());
    }

    private void AppendLog(string text)
    {
        Dispatcher.Invoke(() =>
        {
            LogText.AppendText(text + Environment.NewLine);
            LogText.ScrollToEnd();
        });
    }

    private static List<string> SplitCommandLine(string commandLine)
    {
        var args = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;
        char quote = '\0';

        for (var i = 0; i < commandLine.Length; i++)
        {
            var ch = commandLine[i];
            if ((ch == '"' || ch == '\'') && (!inQuotes || quote == ch))
            {
                if (inQuotes) { inQuotes = false; quote = '\0'; }
                else { inQuotes = true; quote = ch; }
                continue;
            }

            if (char.IsWhiteSpace(ch) && !inQuotes)
            {
                if (current.Length > 0) { args.Add(current.ToString()); current.Clear(); }
                continue;
            }
            current.Append(ch);
        }

        if (current.Length > 0) args.Add(current.ToString());
        return args;
    }
}
