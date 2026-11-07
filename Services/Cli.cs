using System.IO;
using System.Runtime.InteropServices;
using FileConverter.Models;

namespace FileConverter.Services;

/// <summary>
/// Headless mode:
///   FileConverter &lt;input&gt; [--to pdf|docx|png|…] [--out folder]
///   FileConverter --merge out.pdf in1.pdf in2.pdf […]
///   FileConverter --split in.pdf --pages 1-3,5 [--out out.pdf]
///   FileConverter --optimize in.pdf [--out out.pdf]
/// </summary>
public static class Cli
{
    private const int AttachParentProcess = -1;

    [DllImport("kernel32.dll")]
    private static extern bool AttachConsole(int dwProcessId);

    public static bool TryRun(string[] args)
    {
        if (args.Length == 0) return false;

        AttachConsole(AttachParentProcess); // print into the calling terminal when there is one
        try
        {
            if (args[0] is "--help" or "-h" or "/?")
            {
                PrintUsage();
            }
            else
            {
                Run(args);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"error: {ex.Message}");
            Environment.ExitCode = 1;
        }
        finally
        {
            OfficeInterop.ShutdownApps();
        }
        return true;
    }

    private static void Run(string[] args)
    {
        switch (args[0])
        {
            case "--merge":
            {
                if (args.Length < 4)
                    throw new ArgumentException("--merge needs: <output.pdf> <input1.pdf> <input2.pdf> …");
                var mergeOut = Path.GetFullPath(args[1]);
                var inputs = args[2..].Select(Path.GetFullPath).ToList();
                EnsureParentFolder(mergeOut);
                Console.WriteLine($"Merging {inputs.Count} PDFs…");
                PdfTools.Merge(inputs, mergeOut);
                Console.WriteLine($"Done: {mergeOut}");
                return;
            }

            case "--split":
            {
                var input = args.Length > 1 ? Path.GetFullPath(args[1]) : throw new ArgumentException("--split needs an input PDF.");
                string? pages = null, @out = null;
                for (var i = 2; i < args.Length; i++)
                {
                    if (args[i] == "--pages" && i + 1 < args.Length) pages = args[++i];
                    else if (args[i] == "--out" && i + 1 < args.Length) @out = args[++i];
                    else throw new ArgumentException($"unexpected argument '{args[i]}'");
                }
                if (pages is null) throw new ArgumentException("--split needs --pages, e.g. --pages 1-3,5");
                var dest = @out is null ? SidecarPath(input, "-pages.pdf") : Path.GetFullPath(@out);
                EnsureParentFolder(dest);
                PdfTools.ExtractPages(input, dest, pages);
                Console.WriteLine($"Done: {dest}");
                return;
            }

            case "--optimize":
            {
                var input = args.Length > 1 ? Path.GetFullPath(args[1]) : throw new ArgumentException("--optimize needs an input PDF.");
                string? @out = null;
                for (var i = 2; i < args.Length; i++)
                {
                    if (args[i] == "--out" && i + 1 < args.Length) @out = args[++i];
                    else throw new ArgumentException($"unexpected argument '{args[i]}'");
                }
                var dest = @out is null ? SidecarPath(input, "-optimized.pdf") : Path.GetFullPath(@out);
                EnsureParentFolder(dest);
                var before = new FileInfo(input).Length;
                PdfTools.Optimize(input, dest);
                var after = new FileInfo(dest).Length;
                Console.WriteLine($"Done: {dest} ({before / 1024.0 / 1024:F2} MB → {after / 1024.0 / 1024:F2} MB)");
                return;
            }

            default:
                ConvertSingle(args);
                return;
        }
    }

    private static void ConvertSingle(string[] args)
    {
        var input = Path.GetFullPath(args[0]);
        TargetFormat? target = null;
        string? outFolder = null;

        for (var i = 1; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--to" when i + 1 < args.Length:
                    if (!ConversionCatalog.TryParseTarget(args[++i], out var parsed))
                        throw new ArgumentException($"unknown target format '{args[i]}' (pdf, docx, png, jpg, webp, bmp, tiff, gif)");
                    target = parsed;
                    break;
                case "--out" when i + 1 < args.Length:
                    outFolder = args[++i];
                    break;
                default:
                    throw new ArgumentException($"unexpected argument '{args[i]}'. See --help.");
            }
        }

        if (!File.Exists(input))
            throw new FileNotFoundException("input file not found", input);

        var targets = ConversionCatalog.TargetsFor(Path.GetExtension(input));
        if (targets.Length == 0)
            throw new NotSupportedException($"unsupported file type '{Path.GetExtension(input)}'");

        var chosen = target ?? targets[0];
        if (!targets.Contains(chosen))
            throw new NotSupportedException(
                $"'{Path.GetExtension(input)}' files can't convert to {chosen}. Available: {string.Join(", ", targets.Select(t => t.ToString().ToLowerInvariant()))}");

        if (outFolder is not null && !Directory.Exists(outFolder))
            Directory.CreateDirectory(outFolder);
        outFolder = outFolder is null ? null : Path.GetFullPath(outFolder);

        var job = new ConversionJob(input, chosen);
        Console.WriteLine($"Converting {job.FileName} → {chosen}…");
        ConversionEngine.Convert(job, outFolder);
        Console.WriteLine($"Done: {job.OutputPath}");
    }

    private static void EnsureParentFolder(string file)
    {
        var folder = Path.GetDirectoryName(file);
        if (!string.IsNullOrEmpty(folder) && !Directory.Exists(folder))
            Directory.CreateDirectory(folder);
    }

    private static string SidecarPath(string input, string suffix)
    {
        var folder = Path.GetDirectoryName(Path.GetFullPath(input))!;
        var name = Path.GetFileNameWithoutExtension(input);
        return ConversionEngine.UniquePath(Path.Combine(folder, name + suffix));
    }

    private static void PrintUsage()
    {
        Console.WriteLine("File Converter — usage:");
        Console.WriteLine("  FileConverter                                  open the GUI");
        Console.WriteLine("  FileConverter <input> [--to fmt] [--out dir]   convert one file");
        Console.WriteLine("      fmt: pdf, docx, png, jpg, webp, bmp, tiff, gif");
        Console.WriteLine("  FileConverter --merge out.pdf in1.pdf in2.pdf …");
        Console.WriteLine("  FileConverter --split in.pdf --pages 1-3,5 [--out out.pdf]");
        Console.WriteLine("  FileConverter --optimize in.pdf [--out out.pdf]");
        Console.WriteLine("Examples:");
        Console.WriteLine("  FileConverter slides.pptx");
        Console.WriteLine("  FileConverter photo.png --to webp --out D:\\out");
    }
}
