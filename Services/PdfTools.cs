using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace FileConverter.Services;

/// <summary>PDF utilities via PDFsharp — merge, extract page ranges, optimize.</summary>
public static class PdfTools
{
    public static void Merge(IReadOnlyList<string> sources, string dest)
    {
        using var output = new PdfDocument();
        foreach (var source in sources)
        {
            using var input = PdfReader.Open(source, PdfDocumentOpenMode.Import);
            for (var i = 0; i < input.PageCount; i++)
                output.AddPage(input.Pages[i]);
        }
        output.Save(dest);
    }

    /// <summary>Writes the given 1-based pages (e.g. "1-3, 5") to a new PDF.</summary>
    public static void ExtractPages(string source, string dest, string pageRange)
    {
        var pages = ParsePageRange(pageRange, ValidatePageCount(source));
        using var input = PdfReader.Open(source, PdfDocumentOpenMode.Import);
        using var output = new PdfDocument();
        foreach (var page in pages)
            output.AddPage(input.Pages[page - 1]);
        output.Save(dest);
    }

    /// <summary>
    /// Light optimization: rebuilds the document (drops orphaned objects)
    /// and re-serializes. Modest savings; heavy image downsampling needs Ghostscript.
    /// </summary>
    public static void Optimize(string source, string dest)
    {
        using var input = PdfReader.Open(source, PdfDocumentOpenMode.Import);
        using var output = new PdfDocument();
        for (var i = 0; i < input.PageCount; i++)
            output.AddPage(input.Pages[i]);
        output.Save(dest);
    }

    public static int ValidatePageCount(string source)
    {
        using var input = PdfReader.Open(source, PdfDocumentOpenMode.Import);
        return input.PageCount;
    }

    /// <summary>Parses "1-3, 5, 8-9" into sorted, de-duplicated 1-based page numbers.</summary>
    public static List<int> ParsePageRange(string text, int pageCount)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException("Page range is empty. Example: 1-3, 5");

        var pages = new SortedSet<int>();
        foreach (var part in text.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
        {
            var dash = part.IndexOf('-');
            if (dash < 0)
            {
                if (!int.TryParse(part, out var single))
                    throw new ArgumentException($"'{part}' is not a page number.");
                pages.Add(single);
            }
            else
            {
                if (!int.TryParse(part[..dash], out var from) || !int.TryParse(part[(dash + 1)..], out var to))
                    throw new ArgumentException($"'{part}' is not a valid range (use e.g. 2-5).");
                if (from > to) (from, to) = (to, from);
                for (var p = from; p <= to; p++) pages.Add(p);
            }
        }

        var invalid = pages.Where(p => p < 1 || p > pageCount).ToList();
        if (invalid.Count > 0)
            throw new ArgumentException($"Page(s) {string.Join(", ", invalid)} are outside 1–{pageCount}.");

        return pages.ToList();
    }
}
