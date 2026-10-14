using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using WordPdfStudio.Services;

namespace WordPdfStudio;

/// <summary>Automation entry point. Reports contain results only, never the command line or credentials.</summary>
public static class HeadlessRunner
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public static async Task<int> RunAsync(string[] args)
    {
        var timer = Stopwatch.StartNew();
        string? reportPath = null;
        var command = "unknown";
        var output = "";
        IReadOnlyList<string> outputs = Array.Empty<string>();
        var secrets = new List<string>();
        var success = false;
        var error = "";
        try
        {
            var arguments = args.ToList();
            if (arguments.Count > 0 && arguments[^1].StartsWith("--report=", StringComparison.OrdinalIgnoreCase))
            {
                reportPath = ValidateReportPath(arguments[^1][9..]);
                arguments.RemoveAt(arguments.Count - 1);
            }
            if (arguments.Count == 0)
                throw new ArgumentException("缺少命令。支持 convert、merge、extract、split、watermark、encrypt、rotate、images 和 engine-check。");
            command = arguments[0].ToLowerInvariant();
            arguments.RemoveAt(0);

            // Optional input credentials are useful for processing an already encrypted PDF.
            var inputPassword = "";
            var passwordOption = arguments.FindLastIndex(value => value.StartsWith("--input-password=", StringComparison.OrdinalIgnoreCase));
            if (passwordOption >= 0)
            {
                inputPassword = arguments[passwordOption][17..];
                secrets.Add(inputPassword);
                arguments.RemoveAt(passwordOption);
            }

            switch (command)
            {
                case "convert":
                {
                    RequireAtLeast(arguments, 2, "convert INPUT OUTPUT [--quality=N] [--pdfa] [--range=TEXT]");
                    var options = new ConversionOptions();
                    foreach (var option in arguments.Skip(2))
                    {
                        if (option.Equals("--pdfa", StringComparison.OrdinalIgnoreCase))
                            options.PdfA = true;
                        else if (option.StartsWith("--quality=", StringComparison.OrdinalIgnoreCase) &&
                                 int.TryParse(option[10..], NumberStyles.None, CultureInfo.InvariantCulture, out var quality))
                            options.Quality = quality;
                        else if (option.StartsWith("--range=", StringComparison.OrdinalIgnoreCase))
                            options.PageRange = option[8..];
                        else
                            throw new ArgumentException("convert 参数无效。可用选项：--quality=1至100、--pdfa、--range=1-3,5。");
                    }
                    output = Path.GetFullPath(arguments[1]);
                    ProtectReport(reportPath, [arguments[0], output]);
                    await new ConversionService(new EngineManager()).ConvertAsync(arguments[0], output, options);
                    outputs = [output];
                    break;
                }
                case "merge":
                {
                    RequireAtLeast(arguments, 3, "merge OUTPUT INPUT INPUT...");
                    output = Path.GetFullPath(arguments[0]);
                    var inputs = arguments.Skip(1).ToArray();
                    ProtectReport(reportPath, [output, .. inputs]);
                    await Task.Run(() => PdfService.Merge(inputs, output, inputPassword));
                    outputs = [output];
                    break;
                }
                case "extract":
                    RequireExactly(arguments, 3, "extract INPUT OUTPUT RANGE");
                    output = Path.GetFullPath(arguments[1]);
                    ProtectReport(reportPath, [arguments[0], output]);
                    await Task.Run(() => PdfService.Extract(arguments[0], output, arguments[2], inputPassword));
                    outputs = [output];
                    break;
                case "split":
                    RequireExactly(arguments, 2, "split INPUT DIRECTORY");
                    output = Path.GetFullPath(arguments[1]);
                    ProtectReport(reportPath, [arguments[0]]);
                    outputs = await Task.Run(() => PdfService.Split(arguments[0], output, inputPassword));
                    break;
                case "watermark":
                    RequireExactly(arguments, 3, "watermark INPUT OUTPUT TEXT");
                    output = Path.GetFullPath(arguments[1]);
                    ProtectReport(reportPath, [arguments[0], output]);
                    await Task.Run(() => PdfService.Watermark(arguments[0], output, arguments[2], inputPassword));
                    outputs = [output];
                    break;
                case "encrypt":
                    RequireExactly(arguments, 3, "encrypt INPUT OUTPUT PASSWORD");
                    secrets.Add(arguments[2]);
                    output = Path.GetFullPath(arguments[1]);
                    ProtectReport(reportPath, [arguments[0], output]);
                    await Task.Run(() => PdfService.Encrypt(arguments[0], output, arguments[2], inputPassword));
                    outputs = [output];
                    break;
                case "rotate":
                    RequireExactly(arguments, 3, "rotate INPUT OUTPUT DEGREES");
                    if (!int.TryParse(arguments[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out var degrees))
                        throw new ArgumentException("旋转角度必须是整数，例如 90、180 或 270。");
                    output = Path.GetFullPath(arguments[1]);
                    ProtectReport(reportPath, [arguments[0], output]);
                    await Task.Run(() => PdfService.Rotate(arguments[0], output, degrees, inputPassword));
                    outputs = [output];
                    break;
                case "images":
                {
                    RequireAtLeast(arguments, 2, "images OUTPUT INPUT...");
                    output = Path.GetFullPath(arguments[0]);
                    var inputs = arguments.Skip(1).ToArray();
                    ProtectReport(reportPath, [output, .. inputs]);
                    await Task.Run(() => PdfService.ImagesToPdf(inputs, output));
                    outputs = [output];
                    break;
                }
                case "engine-check":
                    RequireExactly(arguments, 0, "engine-check");
                    output = await new EngineManager().EnsureReadyAsync();
                    outputs = [output];
                    break;
                default:
                    command = "unknown";
                    throw new ArgumentException("未知命令。支持 convert、merge、extract、split、watermark、encrypt、rotate、images 和 engine-check。");
            }
            success = true;
        }
        catch (Exception exception)
        {
            error = Redact(exception.Message, secrets);
            App.LogFailure(exception, "cli");
        }

        var report = new AutomationReport(success, command, output, outputs, error, Math.Round(timer.Elapsed.TotalSeconds, 3));
        try
        {
            await WriteReportAsync(reportPath, report);
            Console.Out.WriteLine(JsonSerializer.Serialize(report, JsonOptions));
        }
        catch (Exception exception)
        {
            App.LogFailure(exception, "report");
            Console.Error.WriteLine("无法写入 JSON 结果报告：" + Redact(exception.Message, secrets));
            return 1;
        }
        return success ? 0 : 1;
    }

    internal static async Task WriteAutomationReportAsync(string? reportPath, string command, string output,
        bool success, string error, double elapsedSeconds)
    {
        var report = new AutomationReport(success, command, output,
            success && output.Length > 0 ? [output] : Array.Empty<string>(), error, Math.Round(elapsedSeconds, 3));
        await WriteReportAsync(reportPath is null ? null : ValidateReportPath(reportPath), report);
    }

    private static async Task WriteReportAsync(string? path, AutomationReport report)
    {
        if (path is null)
            return;
        var directory = Path.GetDirectoryName(path)!;
        Directory.CreateDirectory(directory);
        var temporary = Path.Combine(directory, ".wordpdf-report-" + Guid.NewGuid().ToString("N") + ".tmp");
        var created = false;
        try
        {
            await using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None,
                             4096, FileOptions.Asynchronous))
            {
                created = true;
                await JsonSerializer.SerializeAsync(stream, report, JsonOptions);
                await stream.FlushAsync();
            }
            File.Move(temporary, path, overwrite: true);
        }
        finally
        {
            if (created && File.Exists(temporary))
                File.Delete(temporary);
        }
    }

    private static string ValidateReportPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !Path.IsPathFullyQualified(path))
            throw new ArgumentException("--report 必须指定 JSON 报告的绝对路径。");
        if (!Path.GetExtension(path).Equals(".json", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("--report 的文件扩展名必须是 .json，以免覆盖文档。");
        return Path.GetFullPath(path);
    }

    private static void ProtectReport(string? reportPath, IReadOnlyList<string> protectedPaths)
    {
        if (reportPath is not null && protectedPaths.Any(path =>
                Path.GetFullPath(path).Equals(reportPath, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("JSON 报告不能与输入或输出文件使用同一路径。");
    }

    private static string Redact(string message, IReadOnlyList<string> secrets)
    {
        foreach (var secret in secrets.Where(value => !string.IsNullOrEmpty(value)))
            message = message.Replace(secret, "[已隐藏]", StringComparison.Ordinal);
        return message;
    }

    private static void RequireExactly(IReadOnlyCollection<string> arguments, int count, string syntax)
    {
        if (arguments.Count != count)
            throw new ArgumentException("参数数量不正确。用法：--cli " + syntax);
    }

    private static void RequireAtLeast(IReadOnlyCollection<string> arguments, int count, string syntax)
    {
        if (arguments.Count < count)
            throw new ArgumentException("参数不足。用法：--cli " + syntax);
    }

    private sealed record AutomationReport(bool Success, string Command, string Output,
        IReadOnlyList<string> Outputs, string Error, double ElapsedSeconds);
}
