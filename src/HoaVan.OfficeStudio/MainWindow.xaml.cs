using System.Diagnostics;
using System.Globalization;
using System.IO;
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
    private readonly string _outputFolder;

    public MainWindow()
    {
        InitializeComponent();
        _outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "HoaVanOfficeStudio");
        Directory.CreateDirectory(_outputFolder);
        Loaded += MainWindow_Loaded;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        AppendLog($"Output folder: {_outputFolder}");
        AppendLog($"Engine path: {_officeCli.EnginePath}");
        var result = await _officeCli.VersionAsync();
        if (result.Success)
        {
            EngineStatusText.Text = $"OfficeCLI Engine: READY · {result.StandardOutput.Trim()}";
            EngineStatusText.Foreground = System.Windows.Media.Brushes.LightGreen;
            AppendResult("OfficeCLI version", result);
        }
        else
        {
            EngineStatusText.Text = "OfficeCLI Engine: NOT FOUND. Bản Portable cần có tools\\officecli-win-x64.exe.";
            EngineStatusText.Foreground = System.Windows.Media.Brushes.Salmon;
            AppendResult("Kiểm tra engine", result);
        }
    }

    private void AiOutputTypeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (AiFormatHintText is null || AiFileNameText is null) return;
        var type = GetSelectedAiType();
        AiFormatHintText.Text = type switch
        {
            OfficeDocumentType.Word => "Word: dùng # cho tiêu đề lớn, ## cho đề mục. Nội dung thường sẽ thành đoạn văn.",
            OfficeDocumentType.PowerPoint => "PowerPoint: mỗi dòng bắt đầu bằng # là một slide mới; các dòng - bên dưới là nội dung slide.",
            OfficeDocumentType.Excel => "Excel: paste bảng phân cách bằng TAB (khuyên dùng) hoặc dấu phẩy. Mỗi dòng là một hàng.",
            _ => string.Empty
        };
        AiFileNameText.Text = type switch
        {
            OfficeDocumentType.Word => "Hoc_lieu_Word",
            OfficeDocumentType.PowerPoint => "Bai_giang_AI",
            OfficeDocumentType.Excel => "Bang_du_lieu_AI",
            _ => "Hoc_lieu_AI"
        };
    }

    private void QuickTemplateCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (AiContentText is null || QuickTemplateCombo?.SelectedItem is not ComboBoxItem item) return;
        var tag = item.Tag?.ToString();
        if (tag == "LessonPlan")
        {
            AiOutputTypeCombo.SelectedIndex = 0;
            AiContentText.Text = LessonPlanTemplate();
        }
        else if (tag == "TeachingSlides")
        {
            AiOutputTypeCombo.SelectedIndex = 1;
            AiContentText.Text = SlideTemplate();
        }
        else if (tag == "DataTable")
        {
            AiOutputTypeCombo.SelectedIndex = 2;
            AiContentText.Text = ExcelTemplate();
        }
    }

    private void LoadExample_Click(object sender, RoutedEventArgs e)
    {
        AiContentText.Text = GetSelectedAiType() switch
        {
            OfficeDocumentType.Word => LessonPlanTemplate(),
            OfficeDocumentType.PowerPoint => SlideTemplate(),
            OfficeDocumentType.Excel => ExcelTemplate(),
            _ => string.Empty
        };
    }

    private void ClearAiContent_Click(object sender, RoutedEventArgs e) => AiContentText.Clear();

    private async void CreateFromAiContent_Click(object sender, RoutedEventArgs e)
    {
        var content = AiContentText.Text.Trim();
        if (string.IsNullOrWhiteSpace(content))
        {
            MessageBox.Show("Hãy dán nội dung cần tạo trước.", "Hoa Van Office Studio", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var type = GetSelectedAiType();
        var path = BuildOutputPath(AiFileNameText.Text, type);
        await RunBusyAsync("AI Create", async ct =>
        {
            var create = await _officeCli.CreateAsync(path, force: true, ct: ct);
            AppendResult("create", create);
            if (!create.Success) return;

            switch (type)
            {
                case OfficeDocumentType.Word:
                    await BuildWordFromTextAsync(path, content, ct);
                    break;
                case OfficeDocumentType.PowerPoint:
                    await BuildPowerPointFromTextAsync(path, content, ct);
                    break;
                case OfficeDocumentType.Excel:
                    await BuildExcelFromTextAsync(path, content, ct);
                    break;
            }

            var close = await _officeCli.CloseAsync(path, ct);
            AppendResult("save/close", close);
            SetCurrentFile(path);
            OpenFile(path);
        });
    }

    private async Task BuildWordFromTextAsync(string path, string content, CancellationToken ct)
    {
        var lines = NormalizeLines(content);
        foreach (var raw in lines)
        {
            var line = raw.Trim();
            if (string.IsNullOrWhiteSpace(line)) continue;

            string? style = null;
            string text = line;
            if (line.StartsWith("### ")) { style = "Heading3"; text = line[4..].Trim(); }
            else if (line.StartsWith("## ")) { style = "Heading2"; text = line[3..].Trim(); }
            else if (line.StartsWith("# ")) { style = "Heading1"; text = line[2..].Trim(); }
            else if (line.StartsWith("- ")) { text = "• " + line[2..].Trim(); }

            var result = await _officeCli.AddWordParagraphAsync(path, text, style, ct);
            AppendResult($"Word: {TrimForLog(text)}", result);
            if (!result.Success) throw new InvalidOperationException("Không thể thêm nội dung Word. Xem Activity Log.");
        }
    }

    private async Task BuildPowerPointFromTextAsync(string path, string content, CancellationToken ct)
    {
        var slides = ParseSlides(content);
        if (slides.Count == 0) slides.Add(new SlideDraft("Bài giảng", new List<string> { content }));

        for (var s = 0; s < slides.Count; s++)
        {
            var slide = slides[s];
            var addSlide = await _officeCli.AddSlideAsync(path, slide.Title, ct);
            AppendResult($"Slide {s + 1}: {slide.Title}", addSlide);
            if (!addSlide.Success) throw new InvalidOperationException("Không thể tạo slide. Xem Activity Log.");

            var maxItems = Math.Min(slide.Items.Count, 7);
            for (var i = 0; i < maxItems; i++)
            {
                var text = slide.Items[i].TrimStart('-', '•', ' ').Trim();
                if (string.IsNullOrWhiteSpace(text)) continue;
                var y = (3.4 + i * 1.35).ToString("0.00", CultureInfo.InvariantCulture) + "cm";
                var addText = await _officeCli.AddSlideTextAsync(path, s + 1, "• " + text, y, ct);
                AppendResult($"  Nội dung {i + 1}", addText);
                if (!addText.Success) throw new InvalidOperationException("Không thể thêm nội dung slide. Xem Activity Log.");
            }
        }
    }

    private async Task BuildExcelFromTextAsync(string path, string content, CancellationToken ct)
    {
        var rows = NormalizeLines(content).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
        for (var r = 0; r < rows.Count; r++)
        {
            var cells = rows[r].Contains('\t') ? rows[r].Split('\t') : rows[r].Split(',');
            for (var c = 0; c < cells.Length; c++)
            {
                var value = cells[c].Trim();
                if (string.IsNullOrWhiteSpace(value)) continue;
                var cellRef = ColumnName(c + 1) + (r + 1);
                var result = await _officeCli.AddExcelCellAsync(path, cellRef, value, ct);
                AppendResult($"Excel {cellRef}", result);
                if (!result.Success) throw new InvalidOperationException("Không thể ghi dữ liệu Excel. Xem Activity Log.");
            }
        }
    }

    private async void CreateWord_Click(object sender, RoutedEventArgs e) => await CreateBlankAsync(OfficeDocumentType.Word, WordNameText.Text);
    private async void CreatePowerPoint_Click(object sender, RoutedEventArgs e) => await CreateBlankAsync(OfficeDocumentType.PowerPoint, PowerPointNameText.Text);
    private async void CreateExcel_Click(object sender, RoutedEventArgs e) => await CreateBlankAsync(OfficeDocumentType.Excel, ExcelNameText.Text);

    private async Task CreateBlankAsync(OfficeDocumentType type, string name)
    {
        var path = BuildOutputPath(name, type);
        await RunBusyAsync($"Tạo {type}", async ct =>
        {
            var result = await _officeCli.CreateAsync(path, force: true, ct: ct);
            AppendResult("create", result);
            if (result.Success)
            {
                await _officeCli.CloseAsync(path, ct);
                SetCurrentFile(path);
                OpenFile(path);
            }
        });
    }

    private void SelectFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Office files (*.docx;*.pptx;*.xlsx)|*.docx;*.pptx;*.xlsx|All files (*.*)|*.*" };
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
        await RunBusyAsync("Preview HTML", async ct => AppendResult("preview", await _officeCli.PreviewHtmlAsync(_currentFile!, ct)));
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

    private void OpenCurrentFile_Click(object sender, RoutedEventArgs e)
    {
        if (EnsureCurrentFile()) OpenFile(_currentFile!);
    }

    private void OpenOutputFolder_Click(object sender, RoutedEventArgs e) => OpenFile(_outputFolder);

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

    private OfficeDocumentType GetSelectedAiType()
    {
        if (AiOutputTypeCombo.SelectedItem is ComboBoxItem item && Enum.TryParse<OfficeDocumentType>(item.Tag?.ToString(), out var type)) return type;
        return OfficeDocumentType.Word;
    }

    private string BuildOutputPath(string name, OfficeDocumentType type)
    {
        var clean = PathHelper.SanitizeFileName(name);
        return Path.Combine(_outputFolder, clean + type.Extension());
    }

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
        try { await action(_cts.Token); }
        catch (OperationCanceledException) { AppendLog("[CANCELLED] Tác vụ đã bị hủy."); }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] {ex.Message}");
            MessageBox.Show(ex.Message, "Hoa Van Office Studio", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally { IsEnabled = true; }
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

    private static void OpenFile(string path)
    {
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
        catch { }
    }

    private static string[] NormalizeLines(string text) => text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

    private static List<SlideDraft> ParseSlides(string content)
    {
        var slides = new List<SlideDraft>();
        string? title = null;
        var items = new List<string>();

        void Flush()
        {
            if (title is null && items.Count == 0) return;
            slides.Add(new SlideDraft(title ?? "Nội dung", new List<string>(items)));
            title = null;
            items.Clear();
        }

        foreach (var raw in NormalizeLines(content))
        {
            var line = raw.Trim();
            if (line == "---") { Flush(); continue; }
            if (line.StartsWith("# "))
            {
                Flush();
                title = line[2..].Trim();
                continue;
            }
            if (!string.IsNullOrWhiteSpace(line)) items.Add(line);
        }
        Flush();
        return slides;
    }

    private static string ColumnName(int columnNumber)
    {
        var dividend = columnNumber;
        var name = string.Empty;
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            name = Convert.ToChar('A' + modulo) + name;
            dividend = (dividend - modulo) / 26;
        }
        return name;
    }

    private static string TrimForLog(string text) => text.Length <= 44 ? text : text[..44] + "...";

    private static List<string> SplitCommandLine(string commandLine)
    {
        var args = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;
        char quote = '\0';
        foreach (var ch in commandLine)
        {
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

    private static string LessonPlanTemplate() => """
# KẾ HOẠCH BÀI DẠY
## I. Mục tiêu
Kiến thức: ...
Năng lực: ...
Năng lực số / AI: ...
Phẩm chất: ...

## II. Thiết bị và học liệu
- Giáo viên: ...
- Học sinh: ...

## III. Tiến trình dạy học
### Hoạt động 1. Mở đầu
Mục tiêu: ...
Nội dung: ...
Sản phẩm: ...
Tổ chức thực hiện: ...

### Hoạt động 2. Hình thành kiến thức
Mục tiêu: ...
Nội dung: ...
Sản phẩm: ...
Tổ chức thực hiện: ...

### Hoạt động 3. Luyện tập
...

### Hoạt động 4. Vận dụng
...
""";

    private static string SlideTemplate() => """
# BÀI HỌC HÔM NAY
- Tên bài học
- Môn / lớp
- Mục tiêu chính

# KHỞI ĐỘNG
- Câu hỏi dẫn nhập
- Tình huống thực tế

# KIẾN THỨC TRỌNG TÂM
- Ý chính 1
- Ý chính 2
- Ví dụ minh họa

# LUYỆN TẬP
- Nhiệm vụ cá nhân
- Nhiệm vụ nhóm

# CỦNG CỐ
- 3 câu hỏi nhanh
- Exit ticket
""";

    private static string ExcelTemplate() => """
Họ tên\tĐiểm thường xuyên\tGiữa kỳ\tCuối kỳ\tGhi chú
Nguyễn Văn A\t8\t8.5\t9\t
Trần Thị B\t7.5\t8\t8.5\t
Lê Văn C\t9\t9\t9.5\t
""";

    private sealed record SlideDraft(string Title, List<string> Items);
}
