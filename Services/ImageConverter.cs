using System.IO;
using ImageMagick;
using PdfSharp.Pdf;

namespace FileConverter.Services;

/// <summary>
/// Image conversions via Magick.NET. Images→PDF is built with PDFsharp from
/// PNG frames so it never needs Ghostscript (Magick.NET delegates PDF
/// read/write to Ghostscript, which isn't assumed to be installed).
/// </summary>
public static class ImageConverter
{
    private const int LossyQuality = 90;

    public static void Convert(string source, string dest, TargetFormat target)
    {
        if (target == TargetFormat.PDF)
        {
            ImagesToPdf(new[] { source }, dest);
            return;
        }

        var format = MagickFormatFor(target);
        using var images = new MagickImageCollection(source);
        foreach (var image in images)
            image.Quality = LossyQuality;
        images.Write(dest, format);
    }

    /// <summary>Combines one or more image files into a single multi-page PDF.</summary>
    public static void ImagesToPdf(IReadOnlyList<string> sources, string dest)
    {
        using var pdf = new PdfDocument();
        var tempFiles = new List<string>();
        var images = new List<PdfSharp.Drawing.XImage>();

        try
        {
            foreach (var source in sources)
            {
                using var frames = new MagickImageCollection(source);
                foreach (var frame in frames)
                {
                    // Normalize to PNG at a known density so page size in points is stable.
                    if (frame.Density is null || (frame.Density.X <= 0 && frame.Density.Y <= 0))
                        frame.Density = new Density(96);

                    var tempFile = Path.Combine(Path.GetTempPath(), "fileconverter-" + Guid.NewGuid().ToString("N") + ".png");
                    frame.Write(tempFile);
                    tempFiles.Add(tempFile);

                    var image = PdfSharp.Drawing.XImage.FromFile(tempFile);
                    images.Add(image);

                    var page = pdf.AddPage();
                    page.Width = PdfSharp.Drawing.XUnit.FromPoint(image.PointWidth);
                    page.Height = PdfSharp.Drawing.XUnit.FromPoint(image.PointHeight);
                    using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page);
                    gfx.DrawImage(image, 0, 0, image.PointWidth, image.PointHeight);
                }
            }

            pdf.Save(dest);
        }
        finally
        {
            foreach (var image in images)
                image.Dispose();
            foreach (var tempFile in tempFiles)
                try { File.Delete(tempFile); } catch { /* best effort */ }
        }
    }

    private static MagickFormat MagickFormatFor(TargetFormat target) => target switch
    {
        TargetFormat.PNG => MagickFormat.Png,
        TargetFormat.JPG => MagickFormat.Jpeg,
        TargetFormat.WEBP => MagickFormat.WebP,
        TargetFormat.BMP => MagickFormat.Bmp,
        TargetFormat.TIFF => MagickFormat.Tiff,
        TargetFormat.GIF => MagickFormat.Gif,
        _ => throw new NotSupportedException($"{target} is not an image format"),
    };
}
