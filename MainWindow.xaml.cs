using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using FileConverter.Models;
using FileConverter.Services;
using Microsoft.Win32;

namespace FileConverter;

public partial class MainWindow : Window
{
    private const string FileDialogFilter =
        "Supported files|*.pdf;*.docx;*.doc;*.pptx;*.ppt;*.xlsx;*.xls;" +
        "*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.tif;*.tiff;*.gif|All files|*.*";

    private readonly ObservableCollection<ConversionJob> _jobs = new();
    private bool _busy;

    public MainWindow()
    {
        InitializeComponent();
        _jobs.CollectionChanged += (_, _) =>
        {
            EmptyHint.Visibility = _jobs.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            UpdateActionButtons();
        };
        JobsList.ItemsSource = _jobs;
        UpdateActionButtons();
    }

    // ---------- adding files ----------

    private void OnDropZoneClick(object sender, System.Windows.Input.MouseButtonEventArgs e) => BrowseForFiles();

    private void BrowseForFiles()
    {
        if (_busy) return;
        var dlg = new OpenFileDialog { Filter = FileDialogFilter, Multiselect = true };
        if (dlg.ShowDialog(this) != true) return;
        AddFiles(dlg.FileNames);
    }

    private void OnDropZoneDragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
        DropZone.Background = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(0x24, 0x27, 0x33));
        DropOutline.Stroke = (System.Windows.Media.Brush)FindResource("AccentBrush");
    }

    private void OnDropZoneDragLeave(object sender, DragEventArgs e) => ResetDropZoneVisual();

    private void OnDropZoneDrop(object sender, DragEventArgs e)
    {
        ResetDropZoneVisual();
        if (_busy) return;
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
            AddFiles(paths);
    }

    private void ResetDropZoneVisual()
    {
        DropZone.Background = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(0x1D, 0x1F, 0x26));
        DropOutline.Stroke = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(0x3A, 0x3D, 0x46));
    }

    private void AddFiles(IEnumerable<string> paths)
    {
        var skipped = 0;
        var added = 0;
        foreach (var path in paths)
        {
            if (!File.Exists(path)) continue;
            if (!ConversionCatalog.IsSupported(path)) { skipped++; continue; }
            if (_jobs.Any(j => j.SourcePath.Equals(path, StringComparison.OrdinalIgnoreCase))) continue;
            var targets = ConversionCatalog.TargetsFor(Path.GetExtension(path));
            _jobs.Add(new ConversionJob(path, targets[0]));
            added++;
        }

        StatusText.Text = skipped > 0
            ? $"Added {added} file(s) — skipped {skipped} unsupported file(s)"
            : added > 0 ? $"Added {added} file(s)" : StatusText.Text;
    }

    // ---------- convert all ----------

    private void OnConvertAll(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        var pending = _jobs.Where(j => j.Status == JobStatus.Pending).ToList();
        if (pending.Count == 0)
        {
            StatusText.Text = "Nothing to convert — add files first";
            return;
        }

        var total = pending.Count;
        var outputFolder = SameFolderCheck.IsChecked == true ? null : OutputFolderBox.Text.Trim();
        if (outputFolder == "") outputFolder = null;

        SetBusy(true);
        Progress.Maximum = total;
        Progress.Value = 0;
        var finished = 0;

        var thread = new Thread(() =>
        {
            ConversionEngine.Run(pending, outputFolder, job => Dispatcher.BeginInvoke(() =>
            {
                finished++;
                Progress.Value = finished;
                if (job.Status == JobStatus.Running)
                    StatusText.Text = $"Converting {finished}/{total} — {job.FileName}";
            }));

            Dispatcher.BeginInvoke(() =>
            {
                SetBusy(false);
                var done = _jobs.Count(j => j.Status == JobStatus.Done);
                var failed = _jobs.Count(j => j.Status == JobStatus.Failed);
                StatusText.Text = failed > 0
                    ? $"Finished — {done} converted, {failed} failed (see list)"
                    : $"Finished — {done} file(s) converted";
                UpdateActionButtons();
            });
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
    }

    // ---------- pdf tools ----------

    private void OnMerge(object sender, RoutedEventArgs e)
    {
        var pdfs = _jobs.Where(j => j.IsPdf() && j.Status == JobStatus.Pending)
                        .Select(j => j.SourcePath).ToList();
        if (pdfs.Count < 2) return;

        var initialFolder = PickDefaultFolder(pdfs[0]);
        var dlg = new SaveFileDialog
        {
            Filter = "PDF|*.pdf",
            FileName = "merged.pdf",
            InitialDirectory = initialFolder,
        };
        if (dlg.ShowDialog(this) != true) return;

        RunTool($"Merging {pdfs.Count} PDFs…", () => PdfTools.Merge(pdfs, dlg.FileName),
            $"Merged → {dlg.FileName}");
    }

    private void OnSplit(object sender, RoutedEventArgs e)
    {
        var job = SelectedJobs().FirstOrDefault();
        if (job is null || !job.IsPdf()) return;

        int pageCount;
        try { pageCount = PdfTools.ValidatePageCount(job.SourcePath); }
        catch (Exception ex) { StatusText.Text = $"Can't read PDF: {ex.Message}"; return; }

        var range = AskPageRange(pageCount);
        if (range is null) return;

        var dlg = new SaveFileDialog
        {
            Filter = "PDF|*.pdf",
            FileName = Path.GetFileNameWithoutExtension(job.FileName) + "-pages.pdf",
            InitialDirectory = PickDefaultFolder(job.SourcePath),
        };
        if (dlg.ShowDialog(this) != true) return;

        RunTool("Splitting PDF…", () => PdfTools.ExtractPages(job.SourcePath, dlg.FileName, range),
            $"Split → {dlg.FileName}");
    }

    private void OnOptimize(object sender, RoutedEventArgs e)
    {
        var pdfs = SelectedJobs().Where(j => j.IsPdf()).Select(j => j.SourcePath).ToList();
        if (pdfs.Count == 0) return;

        RunTool($"Optimizing {pdfs.Count} PDF(s)…", () =>
        {
            foreach (var pdf in pdfs)
            {
                var folder = Path.GetDirectoryName(pdf)!;
                var name = Path.GetFileNameWithoutExtension(pdf);
                var dest = ConversionEngine.UniquePath(Path.Combine(folder, name + "-optimized.pdf"));
                PdfTools.Optimize(pdf, dest);
            }
        }, "Optimized (files saved next to the originals)");
    }

    private void RunTool(string runningText, Action work, string doneText)
    {
        SetBusy(true);
        StatusText.Text = runningText;
        Task.Run(() =>
        {
            string? error = null;
            try { work(); }
            catch (Exception ex) { error = ex.Message; }
            Dispatcher.BeginInvoke(() =>
            {
                SetBusy(false);
                StatusText.Text = error is null ? doneText : $"Failed: {error}";
            });
        });
    }

    // ---------- small helpers ----------

    private IEnumerable<ConversionJob> SelectedJobs() =>
        JobsList.SelectedItems.Cast<ConversionJob>();

    private string PickDefaultFolder(string fallbackFile) =>
        SameFolderCheck.IsChecked != true && Directory.Exists(OutputFolderBox.Text.Trim())
            ? OutputFolderBox.Text.Trim()
            : Path.GetDirectoryName(fallbackFile)!;

    private void OnSameFolderChanged(object sender, RoutedEventArgs e)
    {
        if (OutputFolderBox is null) return; // fires during InitializeComponent, before the TextBox exists
        OutputFolderBox.IsEnabled = SameFolderCheck.IsChecked != true;
    }

    private void OnBrowseFolder(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog();
        if (dlg.ShowDialog(this) == true)
        {
            OutputFolderBox.Text = dlg.FolderName;
            SameFolderCheck.IsChecked = false;
        }
    }

    private void OnClear(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        _jobs.Clear();
        Progress.Value = 0;
        StatusText.Text = "Add files to get started";
    }

    private void OnRemoveRow(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        if ((sender as Button)?.Tag is ConversionJob job)
            _jobs.Remove(job);
    }

    private void OnOpenFolder(object sender, RoutedEventArgs e)
    {
        var folder = _jobs.FirstOrDefault(j => j.Status == JobStatus.Done && j.OutputPath is not null)?.OutputPath
                     is string done ? Path.GetDirectoryName(done)
                     : SameFolderCheck.IsChecked != true && Directory.Exists(OutputFolderBox.Text.Trim())
                         ? OutputFolderBox.Text.Trim()
                         : _jobs.FirstOrDefault()?.SourcePath is string src ? Path.GetDirectoryName(src) : null;

        if (folder is null || !Directory.Exists(folder))
        {
            StatusText.Text = "No output folder yet — convert something first";
            return;
        }
        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{folder}\"") { UseShellExecute = true });
    }

    private void OnSelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateActionButtons();

    private void UpdateActionButtons()
    {
        var selected = SelectedJobs().ToList();
        MergeBtn.IsEnabled = !_busy && _jobs.Count(j => j.IsPdf() && j.Status == JobStatus.Pending) >= 2;
        SplitBtn.IsEnabled = !_busy && selected.Count == 1 && selected[0].IsPdf();
        OptimizeBtn.IsEnabled = !_busy && selected.Any(j => j.IsPdf());
        ConvertAllBtn.IsEnabled = !_busy && _jobs.Any(j => j.Status == JobStatus.Pending);
        OpenFolderBtn.IsEnabled = !_busy;
    }

    private void SetBusy(bool busy)
    {
        _busy = busy;
        DropZone.IsEnabled = !busy;
        ClearBtn.IsEnabled = !busy && _jobs.Count > 0;
        UpdateActionButtons();
    }

    // ---------- split dialog ----------

    private string? AskPageRange(int pageCount)
    {
        var textBox = new TextBox
        {
            FontSize = 13,
            Padding = new Thickness(8, 6, 8, 6),
            Background = System.Windows.Media.Brushes.White,
        };

        var okBtn = new Button { Content = "Split", Width = 90, Padding = new Thickness(0, 6, 0, 6) };
        var cancelBtn = new Button { Content = "Cancel", Width = 80, Padding = new Thickness(0, 6, 0, 6) };

        var dialog = new Window
        {
            Title = "Split PDF",
            SizeToContent = SizeToContent.WidthAndHeight,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            ResizeMode = ResizeMode.NoResize,
            Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0x22, 0x24, 0x2A)),
            Content = new StackPanel { Margin = new Thickness(18) },
        };

        var panel = (StackPanel)dialog.Content;
        panel.Children.Add(new TextBlock
        {
            Text = $"Extract pages (document has {pageCount} page{(pageCount == 1 ? "" : "s")}):",
            Foreground = System.Windows.Media.Brushes.White,
            FontSize = 13,
            Margin = new Thickness(0, 0, 0, 8),
        });
        panel.Children.Add(textBox);
        panel.Children.Add(new TextBlock
        {
            Text = "Example: 1-3, 5",
            Foreground = System.Windows.Media.Brushes.Gray,
            FontSize = 11,
            Margin = new Thickness(0, 6, 0, 12),
        });
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttons.Children.Add(cancelBtn);
        buttons.Children.Add(okBtn);
        panel.Children.Add(buttons);

        string? result = null;
        okBtn.Click += (_, _) => { result = textBox.Text; dialog.Close(); };
        cancelBtn.Click += (_, _) => dialog.Close();

        textBox.Focus();
        dialog.ShowDialog();

        return string.IsNullOrWhiteSpace(result) ? null : result;
    }
}

file static class JobExtensions
{
    public static bool IsPdf(this ConversionJob job) =>
        job.Extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase);
}
