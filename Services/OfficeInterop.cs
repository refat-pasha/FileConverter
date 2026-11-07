using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace FileConverter.Services;

/// <summary>
/// Office conversions via late-bound COM automation (no interop assemblies).
/// Requires a desktop install of PowerPoint / Word / Excel; must run on an STA thread.
///
/// Every operation runs under a watchdog: if an Office app stops responding (a broken
/// Office install can block saves forever), the process is killed after a timeout so the
/// queue can continue with a fresh instance. Override via FILECONVERTER_OFFICE_TIMEOUT_MS.
/// </summary>
public static class OfficeInterop
{
    // MsoTriState
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // SaveAs formats
    private const int PpSaveAsPDF = 32;
    private const int WdFormatPDF = 17;
    private const int WdFormatXMLDocument = 16;
    private const int XlTypePDF = 0;

    /// <summary>Set by the watchdog when the last operation was aborted by timeout.</summary>
    private static volatile bool _lastOperationTimedOut;

    private sealed class OfficeApp
    {
        public dynamic App = null!;
        public int Pid;
        public bool Dead;
        public required string DisplayName;
        public required string ProcessName;
    }

    private static OfficeApp? _powerPoint;
    private static OfficeApp? _word;
    private static OfficeApp? _excel;

    private static TimeSpan Timeout =>
        int.TryParse(Environment.GetEnvironmentVariable("FILECONVERTER_OFFICE_TIMEOUT_MS"), out var ms) && ms > 0
            ? TimeSpan.FromMilliseconds(ms)
            : TimeSpan.FromSeconds(120);

    public static void OfficeToPdf(string source, string dest)
    {
        switch (Path.GetExtension(source).ToLowerInvariant())
        {
            case ".ppt":
            case ".pptx":
                PowerPointToPdf(source, dest);
                break;
            case ".doc":
            case ".docx":
                WordToPdf(source, dest);
                break;
            case ".xls":
            case ".xlsx":
                ExcelToPdf(source, dest);
                break;
            default:
                throw new NotSupportedException($"No Office→PDF converter for '{source}'");
        }
    }

    private static void PowerPointToPdf(string source, string dest)
    {
        var holder = EnsureApp(ref _powerPoint, "PowerPoint.Application", "PowerPoint", "POWERPNT");
        RunWithWatchdog(holder, () =>
        {
            dynamic? pres = null;
            try
            {
                // WithWindow:=msoFalse keeps the presentation invisible.
                pres = holder.App.Presentations.Open(source, ReadOnly: MsoTrue, Untitled: MsoFalse, WithWindow: MsoFalse);
                pres.SaveAs(dest, PpSaveAsPDF);
            }
            finally
            {
                TryRelease(() => pres!.Close());
                Release(pres);
            }
        });
    }

    private static void WordToPdf(string source, string dest)
    {
        var holder = EnsureWord();
        RunWithWatchdog(holder, () =>
        {
            dynamic? doc = null;
            try
            {
                doc = holder.App.Documents.Open(source, ConfirmConversions: false, ReadOnly: true, AddToRecentFiles: false, Visible: false);
                doc.SaveAs2(dest, WdFormatPDF);
            }
            finally
            {
                TryRelease(() => doc!.Close(false));
                Release(doc);
            }
        });
    }

    public static void PdfToDocx(string source, string dest)
    {
        var holder = EnsureWord();
        RunWithWatchdog(holder, () =>
        {
            dynamic? doc = null;
            try
            {
                // Word opens PDFs natively and converts them to an editable document.
                doc = holder.App.Documents.Open(source, ConfirmConversions: false, ReadOnly: true, AddToRecentFiles: false, Visible: false);
                doc.SaveAs2(dest, WdFormatXMLDocument);
            }
            finally
            {
                TryRelease(() => doc!.Close(false));
                Release(doc);
            }
        });
    }

    private static void ExcelToPdf(string source, string dest)
    {
        var holder = EnsureApp(ref _excel, "Excel.Application", "Excel", "EXCEL");
        RunWithWatchdog(holder, () =>
        {
            holder.App.Visible = false;
            holder.App.DisplayAlerts = false;
            dynamic? wb = null;
            try
            {
                wb = holder.App.Workbooks.Open(source, UpdateLinks: 0, ReadOnly: true);
                wb.ExportAsFixedFormat(XlTypePDF, dest);
            }
            finally
            {
                TryRelease(() => wb!.Close(false));
                Release(wb);
            }
        });
    }

    private static OfficeApp EnsureWord()
    {
        var holder = EnsureApp(ref _word, "Word.Application", "Word", "WINWORD");
        try
        {
            holder.App.Visible = false;
            holder.App.DisplayAlerts = 0; // wdAlertsNone — suppresses the PDF-conversion prompt
        }
        catch (COMException)
        {
            // App might have died; next operation will fail and be recreated after watchdog/quit.
        }
        return holder;
    }

    private static OfficeApp EnsureApp(ref OfficeApp? slot, string progId, string displayName, string processName)
    {
        if (slot is { Dead: false }) return slot;

        var type = Type.GetTypeFromProgID(progId)
            ?? throw new InvalidOperationException($"{displayName} is not installed (COM ProgID '{progId}' not found).");

        // Snapshot PIDs so we can identify the process we are about to spawn (for the watchdog).
        var before = Process.GetProcessesByName(processName).Select(p => p.Id).ToHashSet();
        var app = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException($"Could not start {displayName}.");
        var pid = Process.GetProcessesByName(processName).FirstOrDefault(p => !before.Contains(p.Id))?.Id ?? 0;

        slot = new OfficeApp { App = app, Pid = pid, DisplayName = displayName, ProcessName = processName };
        return slot;
    }

    private static void RunWithWatchdog(OfficeApp holder, Action action)
    {
        _lastOperationTimedOut = false;
        using var timer = new System.Threading.Timer(_ =>
        {
            // Set the flags BEFORE killing so the worker's pending COMException is
            // recognised as a timeout even if it surfaces while Kill() is still in flight.
            holder.Dead = true;
            _lastOperationTimedOut = true;
            KillOurProcess(holder);
        }, null, Timeout, System.Threading.Timeout.InfiniteTimeSpan);

        try
        {
            action();
        }
        catch (COMException) when (_lastOperationTimedOut)
        {
            throw new TimeoutException(
                $"{holder.DisplayName} stopped responding after {Timeout.TotalSeconds:0}s and was closed. " +
                "If this keeps happening, repair Office: Windows Settings → Apps → Microsoft Office → Modify → Quick Repair " +
                "(and make sure Office is signed in / activated).");
        }
    }

    /// <summary>
    /// Kills the Office process belonging to our automation instance. If we could not
    /// resolve its PID (e.g. COM attached to an existing process), falls back to killing
    /// window-less processes of that Office app — user-visible documents are spared.
    /// </summary>
    private static void KillOurProcess(OfficeApp holder)
    {
        try
        {
            if (holder.Pid != 0)
            {
                var p = Process.GetProcessById(holder.Pid);
                if (!p.HasExited) p.Kill();
                return;
            }
            foreach (var p in Process.GetProcessesByName(holder.ProcessName))
            {
                if (string.IsNullOrEmpty(p.MainWindowTitle))
                    try { p.Kill(); } catch { /* racing with exit */ }
            }
        }
        catch { /* process already gone */ }
    }

    /// <summary>Quits and releases any cached Office applications.</summary>
    public static void ShutdownApps()
    {
        foreach (var slot in new[] { _powerPoint, _word, _excel })
        {
            if (slot is null) continue;
            if (!slot.Dead) TryRelease(() => slot.App!.Quit());
            Release(slot.App);
            slot.Dead = true;
            KillOurProcess(slot); // Quit() can itself hang on a broken install
        }
        _powerPoint = null;
        _word = null;
        _excel = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private static void TryRelease(Action action)
    {
        try { action(); }
        catch { /* Office already closed or busy — best effort */ }
    }

    private static void Release(object? comObject)
    {
        if (comObject is null) return;
        try { Marshal.FinalReleaseComObject(comObject); }
        catch { /* already released */ }
    }
}
