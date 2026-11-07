using System.IO;
using FileConverter.Models;

namespace FileConverter.Services;

/// <summary>Runs conversion jobs sequentially on a dedicated STA thread (COM requires STA).</summary>
public static class ConversionEngine
{
    /// <summary>
    /// Converts the given pending jobs one by one. Returns when all are done.
    /// Callbacks are invoked from the worker thread — marshal to the UI thread yourself.
    /// </summary>
    public static void Run(IEnumerable<ConversionJob> jobs, string? outputFolder = null, Action<ConversionJob>? onJobChanged = null)
    {
        var pending = jobs.Where(j => j.Status == JobStatus.Pending).ToList();
        try
        {
            foreach (var job in pending)
            {
                job.Status = JobStatus.Running;
                onJobChanged?.Invoke(job);
                try
                {
                    Convert(job, outputFolder);
                    job.Status = JobStatus.Done;
                }
                catch (Exception ex)
                {
                    job.Error = FriendlyError(ex, job.SourcePath);
                    job.Status = JobStatus.Failed;
                }
                onJobChanged?.Invoke(job);
            }
        }
        finally
        {
            OfficeInterop.ShutdownApps();
        }
    }

    public static void Convert(ConversionJob job, string? outputFolder = null)
    {
        var source = job.SourcePath;
        var dest = UniquePath(BuildOutputPath(job, outputFolder));
        job.OutputPath = dest;

        var ext = Path.GetExtension(source).ToLowerInvariant();
        switch (ext)
        {
            case ".ppt":
            case ".pptx":
            case ".doc":
            case ".docx":
            case ".xls":
            case ".xlsx":
                if (job.TargetFormat != TargetFormat.PDF)
                    throw new NotSupportedException($"{ext} files can only convert to PDF.");
                OfficeInterop.OfficeToPdf(source, dest);
                break;

            case ".pdf":
                if (job.TargetFormat != TargetFormat.DOCX)
                    throw new NotSupportedException("PDF files can convert to DOCX (use PDF tools for merge/split).");
                OfficeInterop.PdfToDocx(source, dest);
                break;

            default:
                if (!ConversionCatalog.Matrix.ContainsKey(ext))
                    throw new NotSupportedException($"Unsupported file type '{ext}'.");
                ImageConverter.Convert(source, dest, job.TargetFormat);
                break;
        }
    }

    public static string BuildOutputPath(ConversionJob job, string? outputFolder = null)
    {
        var folder = string.IsNullOrWhiteSpace(outputFolder)
            ? Path.GetDirectoryName(job.SourcePath)!
            : outputFolder;
        folder = Path.GetFullPath(folder);
        Directory.CreateDirectory(folder);
        var name = Path.GetFileNameWithoutExtension(job.FileName);
        return Path.Combine(folder, name + ConversionCatalog.ExtensionFor(job.TargetFormat));
    }

    /// <summary>Returns a path that doesn't collide: "report.pdf" → "report (2).pdf".</summary>
    public static string UniquePath(string path)
    {
        if (!File.Exists(path)) return path;
        var folder = Path.GetDirectoryName(path)!;
        var name = Path.GetFileNameWithoutExtension(path);
        var ext = Path.GetExtension(path);
        for (var i = 2; ; i++)
        {
            var candidate = Path.Combine(folder, $"{name} ({i}){ext}");
            if (!File.Exists(candidate)) return candidate;
        }
    }

    private static string FriendlyError(Exception ex, string source)
    {
        if (ex is TimeoutException)
            return ex.Message;

        var ext = Path.GetExtension(source).ToLowerInvariant();
        var isOfficeSource = ext is ".ppt" or ".pptx" or ".doc" or ".docx" or ".xls" or ".xlsx" or ".pdf";
        var hint = isOfficeSource
            ? " — the file may be open in Office; close it and try again."
            : "";
        var message = ex.Message.Trim();
        return message.Length > 400 ? message[..400] + "…" : message + hint;
    }
}
