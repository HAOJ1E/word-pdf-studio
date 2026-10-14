using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Windows;

namespace WordPdfStudio.Services;

/// <summary>
/// Creates derived PDFs without changing input files. Page import preserves visual page content,
/// but is not a preservation workflow for document-level forms, bookmarks or digital signatures.
/// Password protection is applied only by Encrypt; other derived outputs are unencrypted.
/// </summary>
public static class PdfService
{
    // Prefer a normal Windows CJK font for compact subsets; the bundled font is the fallback on a clean PC.
    static PdfService() => GlobalFontSettings.FallbackFontResolver = BundledFontResolver.Instance;

    private static readonly HashSet<string> ImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"
    };

    public static int GetPageCount(string source, string inputPassword = "")
    {
        var path = ValidateInput(source, pdfOnly: true);
        using var input = OpenPdf(path, inputPassword);
        return input.PageCount;
    }

    public static void Merge(IReadOnlyList<string> inputs, string output, string inputPassword = "")
    {
        var paths = ValidateInputs(inputs, pdfOnly: true);
        if (paths.Count < 2)
            throw new ArgumentException("合并 PDF 至少需要选择 2 个文件。", nameof(inputs));
        var destination = ValidateOutput(output, paths);
        WithFileErrors(() =>
        {
            using var result = NewDocument("合并 PDF");
            foreach (var path in paths)
            {
                using var input = OpenPdf(path, inputPassword);
                CopyPages(input, result, Enumerable.Range(1, input.PageCount));
            }
            SaveAtomic(result, destination);
        });
    }

    public static void Extract(string source, string output, string range, string inputPassword = "")
    {
        if (string.IsNullOrWhiteSpace(range))
            throw new ArgumentException("请输入要提取的页码，例如 1-3,5,8；提取所有页面可填写“全部”。", nameof(range));
        var path = ValidateInput(source, pdfOnly: true);
        var destination = ValidateOutput(output, [path]);
        WithFileErrors(() =>
        {
            using var input = OpenPdf(path, inputPassword);
            var pages = ParsePageRange(range, input.PageCount);
            using var result = NewDocument("提取 PDF 页面");
            CopyPages(input, result, pages);
            SaveAtomic(result, destination);
        });
    }

    /// <summary>
    /// Writes every page to a distinct PDF. Existing files are never overwritten; on failure,
    /// files committed by this call are rolled back and this call's staging files are removed.
    /// </summary>
    public static IReadOnlyList<string> Split(string source, string outputDirectory, string inputPassword = "")
    {
        var path = ValidateInput(source, pdfOnly: true);
        var directory = NormalizePath(outputDirectory, "请选择拆分文件的保存文件夹。");
        if (File.Exists(directory))
            throw new ArgumentException("拆分保存位置必须是文件夹。", nameof(outputDirectory));

        return WithFileErrors(() =>
        {
            using var input = OpenPdf(path, inputPassword);
            var baseName = Path.GetFileNameWithoutExtension(path);
            if (baseName.Length > 100)
                baseName = baseName[..100];
            var digits = Math.Max(3, input.PageCount.ToString(CultureInfo.InvariantCulture).Length);
            var destinations = Enumerable.Range(1, input.PageCount)
                .Select(page => Path.Combine(directory,
                    $"{baseName}_第{page.ToString($"D{digits}", CultureInfo.InvariantCulture)}页.pdf"))
                .ToArray();
            foreach (var destination in destinations)
                ValidateOutput(destination, [path]);

            Directory.CreateDirectory(directory);
            var staged = new List<string>();
            var committed = new List<string>();
            try
            {
                for (var index = 0; index < input.PageCount; index++)
                {
                    using var result = NewDocument($"拆分 PDF：第 {index + 1} 页");
                    result.AddPage(input.Pages[index]);
                    staged.Add(SaveStaged(result, directory));
                }

                // Stage the complete batch before exposing any result file.
                for (var index = 0; index < destinations.Length; index++)
                {
                    File.Move(staged[index], destinations[index], overwrite: false);
                    committed.Add(destinations[index]);
                }
                return (IReadOnlyList<string>)Array.AsReadOnly(destinations);
            }
            catch
            {
                foreach (var file in committed)
                    DeleteOwnTemporaryFile(file);
                throw;
            }
            finally
            {
                foreach (var file in staged)
                    DeleteOwnTemporaryFile(file);
            }
        });
    }

    public static void Watermark(string source, string output, string text, string inputPassword = "")
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException("请输入水印文字。", nameof(text));
        var label = text.Trim().Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
        if (label.Length > 120)
            throw new ArgumentException("水印文字请控制在 120 个字符以内。", nameof(text));
        if (label.Any(char.IsControl))
            throw new ArgumentException("水印文字包含无法显示的控制字符，请删除后重试。", nameof(text));
        var path = ValidateInput(source, pdfOnly: true);
        var destination = ValidateOutput(output, [path]);
        WithFileErrors(() =>
        {
            using var input = OpenPdf(path, inputPassword);
            using var result = NewDocument("添加文字水印");
            CopyPages(input, result, Enumerable.Range(1, input.PageCount));
            const string family = "SimHei";
            var options = new XPdfFontOptions(PdfFontEncoding.Unicode);
            var brush = new XSolidBrush(XColor.FromArgb(42, 73, 92, 120));
            foreach (var page in result.Pages)
            {
                using var graphics = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
                var media = page.MediaBox;
                var crop = page.HasCropBox ? page.CropBox : media;
                var left = Math.Max(media.X1, crop.X1);
                var bottom = Math.Max(media.Y1, crop.Y1);
                var right = Math.Min(media.X2, crop.X2);
                var top = Math.Min(media.Y2, crop.Y2);
                var width = right - left;
                var height = top - bottom;
                if (!IsPositiveFinite(width) || !IsPositiveFinite(height))
                    throw new InvalidOperationException("PDF 包含无效的页面尺寸，无法添加水印。");

                // /Rotate is a viewer rotation; compensate to keep the visible diagonal consistent.
                var angle = -32d - NormalizeRotation(page.Rotate);
                var radians = angle * Math.PI / 180;
                var allowedWidth = 0.78 * Math.Min(
                    width / Math.Max(0.01, Math.Abs(Math.Cos(radians))),
                    height / Math.Max(0.01, Math.Abs(Math.Sin(radians))));
                var fontSize = Math.Min(44, Math.Min(width, height) * 0.10);
                var font = new XFont(family, fontSize, XFontStyleEx.Regular, options);
                var measured = graphics.MeasureString(label, font);
                if (measured.Width > allowedWidth)
                {
                    fontSize *= allowedWidth / measured.Width;
                    font = new XFont(family, fontSize, XFontStyleEx.Regular, options);
                    measured = graphics.MeasureString(label, font);
                }
                var state = graphics.Save();
                graphics.TranslateTransform((left + right) / 2, page.Height.Point - (bottom + top) / 2);
                graphics.RotateTransform(angle);
                graphics.DrawString(label, font, brush,
                    new XRect(-measured.Width / 2, -measured.Height / 2, measured.Width, measured.Height),
                    XStringFormats.Center);
                graphics.Restore(state);
            }
            SaveAtomic(result, destination);
        });
    }

    public static void Encrypt(string source, string output, string password, string inputPassword = "")
    {
        if (string.IsNullOrWhiteSpace(password))
            throw new ArgumentException("请设置非空的 PDF 打开密码。", nameof(password));
        if (Encoding.UTF8.GetByteCount(password) > 127)
            throw new ArgumentException("密码太长，UTF-8 编码后不能超过 127 字节（约 42 个汉字）。", nameof(password));
        var path = ValidateInput(source, pdfOnly: true);
        var destination = ValidateOutput(output, [path]);
        WithFileErrors(() =>
        {
            using var input = OpenPdf(path, inputPassword);
            using var result = NewDocument("加密 PDF");
            CopyPages(input, result, Enumerable.Range(1, input.PageCount));
            result.SecuritySettings.UserPassword = password;
            // Never reuse the user password as the owner credential.
            result.SecuritySettings.OwnerPassword = Convert.ToHexString(RandomNumberGenerator.GetBytes(32));
            result.SecurityHandler.SetEncryptionToV5();
            SaveAtomic(result, destination);
        });
    }

    public static void Rotate(string source, string output, int degrees, string inputPassword = "")
    {
        if (degrees % 90 != 0)
            throw new ArgumentException("旋转角度必须是 90 度的整数倍，例如 90、180 或 270。", nameof(degrees));
        var rotation = NormalizeRotation(degrees);
        if (rotation == 0)
            throw new ArgumentException("请选择有效的旋转角度：90、180 或 270 度。", nameof(degrees));
        var path = ValidateInput(source, pdfOnly: true);
        var destination = ValidateOutput(output, [path]);
        WithFileErrors(() =>
        {
            using var input = OpenPdf(path, inputPassword);
            using var result = NewDocument("旋转 PDF 页面");
            CopyPages(input, result, Enumerable.Range(1, input.PageCount));
            foreach (var page in result.Pages)
                page.Rotate = NormalizeRotation(NormalizeRotation(page.Rotate) + rotation);
            SaveAtomic(result, destination);
        });
    }

    /// <summary>Creates one A4 page per image, preserving aspect ratio and never enlarging an image.</summary>
    public static void ImagesToPdf(IReadOnlyList<string> inputs, string output)
    {
        var paths = ValidateInputs(inputs, pdfOnly: false);
        var destination = ValidateOutput(output, paths);
        WithFileErrors(() =>
        {
            using var result = NewDocument("图片转 PDF");
            foreach (var path in paths)
            {
                using var image = OpenImage(path);
                var originalWidth = IsPositiveFinite(image.PointWidth) ? image.PointWidth : image.PixelWidth * 0.75;
                var originalHeight = IsPositiveFinite(image.PointHeight) ? image.PointHeight : image.PixelHeight * 0.75;
                if (!IsPositiveFinite(originalWidth) || !IsPositiveFinite(originalHeight))
                    throw new InvalidOperationException($"图片“{Path.GetFileName(path)}”尺寸无效。");
                var page = result.AddPage();
                page.Size = PageSize.A4;
                page.Orientation = originalWidth > originalHeight ? PageOrientation.Landscape : PageOrientation.Portrait;
                const double margin = 28.35; // 10 mm.
                var scale = Math.Min(1, Math.Min(
                    (page.Width.Point - margin * 2) / originalWidth,
                    (page.Height.Point - margin * 2) / originalHeight));
                var width = originalWidth * scale;
                var height = originalHeight * scale;
                using var graphics = XGraphics.FromPdfPage(page);
                graphics.DrawImage(image, (page.Width.Point - width) / 2,
                    (page.Height.Point - height) / 2, width, height);
            }
            SaveAtomic(result, destination);
        });
    }

    /// <summary>Returns distinct 1-based page numbers in the requested order. Accepts all/全部/*.</summary>
    public static IReadOnlyList<int> ParsePageRange(string range, int pageCount)
    {
        if (pageCount < 1)
            throw new ArgumentOutOfRangeException(nameof(pageCount), "PDF 没有可提取的页面。");
        if (string.IsNullOrWhiteSpace(range))
            throw new ArgumentException("请输入页码，例如 1-3,5,8。", nameof(range));
        var expression = range.Trim();
        if (expression.Equals("all", StringComparison.OrdinalIgnoreCase) || expression is "全部" or "所有" or "*")
            return Array.AsReadOnly(Enumerable.Range(1, pageCount).ToArray());
        expression = expression.Replace('，', ',').Replace('、', ',').Replace('；', ',').Replace(';', ',')
            .Replace('－', '-').Replace('–', '-');
        var result = new List<int>();
        var seen = new HashSet<int>();
        foreach (var section in expression.Split(','))
        {
            var part = section.Trim();
            var endpoints = part.Split('-');
            if (endpoints.Length is < 1 or > 2 || part.Length == 0)
                throw new ArgumentException($"页码格式不正确：“{part}”。请使用 1-3,5,8 这样的格式。", nameof(range));
            var first = ParsePage(endpoints[0], pageCount);
            var last = endpoints.Length == 2 ? ParsePage(endpoints[1], pageCount) : first;
            if (last < first)
                throw new ArgumentException($"页码区间“{part}”顺序不正确，起始页不能大于结束页。", nameof(range));
            for (var page = first; ; page++)
            {
                if (seen.Add(page))
                    result.Add(page);
                if (page == last)
                    break; // Avoid overflow even if pageCount is Int32.MaxValue.
            }
        }
        return result.AsReadOnly();
    }

    private static int ParsePage(string value, int pageCount)
    {
        if (!int.TryParse(value.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out var page))
            throw new ArgumentException($"页码“{value.Trim()}”不是有效的正整数。", nameof(value));
        if (page < 1 || page > pageCount)
            throw new ArgumentException($"页码 {page} 超出范围，此 PDF 共 {pageCount} 页。", nameof(value));
        return page;
    }

    private static int NormalizeRotation(int degrees) => ((degrees % 360) + 360) % 360;
    private static bool IsPositiveFinite(double value) => double.IsFinite(value) && value > 0;

    private static PdfDocument NewDocument(string purpose)
    {
        var document = new PdfDocument();
        document.Info.Creator = "文页 PDF";
        document.Info.Title = purpose;
        document.Info.Subject = "由文档工坊生成的派生文件；原文件不变。表单交互、书签及数字签名不保证保留。";
        return document;
    }

    private static void CopyPages(PdfDocument input, PdfDocument result, IEnumerable<int> pages)
    {
        foreach (var page in pages)
            result.AddPage(input.Pages[page - 1]);
    }

    private static PdfDocument OpenPdf(string path, string password)
    {
        try
        {
            var document = PdfReader.Open(path, password ?? string.Empty, PdfDocumentOpenMode.Import);
            if (document.PageCount > 0)
                return document;
            document.Dispose();
            throw new InvalidOperationException($"PDF“{Path.GetFileName(path)}”没有可处理的页面。");
        }
        catch (PdfReaderException exception)
        {
            var message = exception.Message.ToLowerInvariant();
            if (message.Contains("password") || message.Contains("encrypted"))
                throw new InvalidOperationException($"无法打开“{Path.GetFileName(path)}”：文件已加密，请填写正确的原 PDF 密码后重试。", exception);
            throw new InvalidOperationException($"无法读取“{Path.GetFileName(path)}”：PDF 可能已损坏或使用了暂不支持的格式。", exception);
        }
        catch (UnauthorizedAccessException exception)
        {
            throw new InvalidOperationException($"没有权限读取“{Path.GetFileName(path)}”，请将文件复制到可访问的文件夹。", exception);
        }
        catch (IOException exception)
        {
            throw new InvalidOperationException($"无法读取“{Path.GetFileName(path)}”，请检查文件是否被占用或已经移动。", exception);
        }
    }

    private static XImage OpenImage(string path)
    {
        try
        {
            return XImage.FromFile(path);
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException or System.Runtime.InteropServices.ExternalException)
        {
            throw new InvalidOperationException($"无法读取图片“{Path.GetFileName(path)}”，图片可能已损坏；请尝试另存为 JPG 或 PNG。", exception);
        }
    }

    private sealed class BundledFontResolver : IFontResolver
    {
        public const string FamilyName = "WordPdf Noto CJK";
        private const string FaceName = "WordPdf.NotoSansCJKsc.Regular";
        public static readonly BundledFontResolver Instance = new();
        private static readonly Lazy<byte[]> FontBytes = new(LoadFont, true);

        public FontResolverInfo? ResolveTypeface(string familyName, bool bold, bool italic) =>
            new(FaceName, bold, italic);

        public byte[]? GetFont(string faceName) =>
            faceName.Equals(FaceName, StringComparison.Ordinal) ? FontBytes.Value : null;

        private static byte[] LoadFont()
        {
            var resource = Application.GetResourceStream(
                new Uri("pack://application:,,,/Assets/NotoSansCJKsc-Regular.otf", UriKind.Absolute))
                ?? throw new InvalidOperationException("程序缺少内置中文字体，请重新获取完整程序。");
            using (resource.Stream)
            using (var memory = new MemoryStream())
            {
                resource.Stream.CopyTo(memory);
                return memory.ToArray();
            }
        }
    }

    private static List<string> ValidateInputs(IReadOnlyList<string> inputs, bool pdfOnly)
    {
        if (inputs is null || inputs.Count == 0)
            throw new ArgumentException(pdfOnly ? "请先选择 PDF 文件。" : "请先选择图片。", nameof(inputs));
        // Order is meaningful; a file can intentionally be included more than once in a merge.
        return inputs.Select(path => ValidateInput(path, pdfOnly)).ToList();
    }

    private static string ValidateInput(string path, bool pdfOnly)
    {
        var fullPath = NormalizePath(path, pdfOnly ? "请选择 PDF 文件。" : "请选择图片文件。");
        if (!File.Exists(fullPath))
            throw new FileNotFoundException($"文件不存在：“{Path.GetFileName(fullPath)}”。", fullPath);
        var extension = Path.GetExtension(fullPath);
        if (pdfOnly ? !extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase) : !ImageExtensions.Contains(extension))
            throw new ArgumentException(pdfOnly ? "此操作仅支持 .pdf 文件。" : "支持 JPG、PNG、BMP、GIF 和 TIFF 图片。", nameof(path));
        return fullPath;
    }

    private static string ValidateOutput(string output, IReadOnlyList<string> inputs)
    {
        var destination = NormalizePath(output, "请选择 PDF 的保存位置。");
        if (!Path.GetExtension(destination).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("输出文件的扩展名必须是 .pdf。", nameof(output));
        if (inputs.Any(input => input.Equals(destination, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("输出位置不能与原文件相同，请使用新的文件名。", nameof(output));
        if (File.Exists(destination) || Directory.Exists(destination))
            throw new IOException($"保存位置已存在：“{Path.GetFileName(destination)}”。请使用其他文件名。原文件不会被覆盖。");
        return destination;
    }

    private static string NormalizePath(string path, string emptyMessage)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException(emptyMessage, nameof(path));
        try
        {
            return Path.GetFullPath(path);
        }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException or PathTooLongException)
        {
            throw new ArgumentException("文件路径无效或过长，请更换保存位置。", nameof(path), exception);
        }
    }

    private static string TemporaryPath(string directory) => Path.Combine(directory, $".wordpdf-{Guid.NewGuid():N}.tmp");

    private static void SaveAtomic(PdfDocument document, string destination)
    {
        var directory = Path.GetDirectoryName(destination)!;
        Directory.CreateDirectory(directory);
        var temporary = SaveStaged(document, directory);
        try
        {
            // Same-directory rename is atomic and will fail if a competing writer created the destination.
            File.Move(temporary, destination, overwrite: false);
        }
        finally
        {
            DeleteOwnTemporaryFile(temporary);
        }
    }

    private static string SaveStaged(PdfDocument document, string directory)
    {
        var temporary = TemporaryPath(directory);
        var created = false;
        try
        {
            using var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            created = true;
            document.Save(stream, closeStream: false);
            stream.Flush(flushToDisk: true);
            if (stream.Length < 8)
                throw new InvalidOperationException("未生成有效的 PDF，请重试。");
            return temporary;
        }
        catch
        {
            // A CreateNew collision does not make the existing file ours to delete.
            if (created)
                DeleteOwnTemporaryFile(temporary);
            throw;
        }
    }

    private static void DeleteOwnTemporaryFile(string path)
    {
        try { File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private static void WithFileErrors(Action action) => WithFileErrors(() => { action(); return true; });

    private static T WithFileErrors<T>(Func<T> action)
    {
        try { return action(); }
        catch (UnauthorizedAccessException exception)
        {
            throw new InvalidOperationException("没有权限保存文件，请选择您有写入权限的文件夹。", exception);
        }
        catch (IOException exception)
        {
            throw new InvalidOperationException("文件保存失败，请检查剩余磁盘空间、文件占用和是否存在同名文件。", exception);
        }
    }
}
