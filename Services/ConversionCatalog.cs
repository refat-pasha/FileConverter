using System.IO;

namespace FileConverter.Services;

public enum TargetFormat { PDF, DOCX, PNG, JPG, WEBP, BMP, TIFF, GIF }

/// <summary>Which target formats are offered for each source extension.</summary>
public static class ConversionCatalog
{
    private static readonly TargetFormat[] OfficeToPdf = { TargetFormat.PDF };
    private static readonly TargetFormat[] PdfTargets = { TargetFormat.DOCX };
    private static readonly TargetFormat[] PngSrc = { TargetFormat.JPG, TargetFormat.WEBP, TargetFormat.BMP, TargetFormat.TIFF, TargetFormat.GIF, TargetFormat.PDF };
    private static readonly TargetFormat[] JpgSrc = { TargetFormat.PNG, TargetFormat.WEBP, TargetFormat.BMP, TargetFormat.TIFF, TargetFormat.GIF, TargetFormat.PDF };
    private static readonly TargetFormat[] WebpSrc = { TargetFormat.PNG, TargetFormat.JPG, TargetFormat.BMP, TargetFormat.TIFF, TargetFormat.GIF, TargetFormat.PDF };
    private static readonly TargetFormat[] BmpSrc = { TargetFormat.PNG, TargetFormat.JPG, TargetFormat.WEBP, TargetFormat.TIFF, TargetFormat.GIF, TargetFormat.PDF };
    private static readonly TargetFormat[] TiffSrc = { TargetFormat.PNG, TargetFormat.JPG, TargetFormat.WEBP, TargetFormat.BMP, TargetFormat.GIF, TargetFormat.PDF };
    private static readonly TargetFormat[] GifSrc = { TargetFormat.PNG, TargetFormat.JPG, TargetFormat.WEBP, TargetFormat.BMP, TargetFormat.TIFF, TargetFormat.PDF };

    public static readonly IReadOnlyDictionary<string, TargetFormat[]> Matrix =
        new Dictionary<string, TargetFormat[]>(StringComparer.OrdinalIgnoreCase)
        {
            [".ppt"] = OfficeToPdf,
            [".pptx"] = OfficeToPdf,
            [".doc"] = OfficeToPdf,
            [".docx"] = OfficeToPdf,
            [".xls"] = OfficeToPdf,
            [".xlsx"] = OfficeToPdf,
            [".pdf"] = PdfTargets,
            [".png"] = PngSrc,
            [".jpg"] = JpgSrc,
            [".jpeg"] = JpgSrc,
            [".webp"] = WebpSrc,
            [".bmp"] = BmpSrc,
            [".tif"] = TiffSrc,
            [".tiff"] = TiffSrc,
            [".gif"] = GifSrc,
        };

    public static bool IsSupported(string path) => Matrix.ContainsKey(Path.GetExtension(path));

    public static TargetFormat[] TargetsFor(string extension) =>
        Matrix.TryGetValue(extension, out var t) ? t : Array.Empty<TargetFormat>();

    public static string ExtensionFor(TargetFormat format) => format switch
    {
        TargetFormat.PDF => ".pdf",
        TargetFormat.DOCX => ".docx",
        TargetFormat.PNG => ".png",
        TargetFormat.JPG => ".jpg",
        TargetFormat.WEBP => ".webp",
        TargetFormat.BMP => ".bmp",
        TargetFormat.TIFF => ".tiff",
        TargetFormat.GIF => ".gif",
        _ => throw new ArgumentOutOfRangeException(nameof(format)),
    };

    public static bool TryParseTarget(string text, out TargetFormat format) =>
        Enum.TryParse(text.Trim().TrimStart('.', '-'), ignoreCase: true, out format)
        && Enum.IsDefined(typeof(TargetFormat), format);
}
