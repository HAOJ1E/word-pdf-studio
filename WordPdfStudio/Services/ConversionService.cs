using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace WordPdfStudio.Services;

public sealed class ConversionOptions
{
    public int Quality { get; set; } = 90;
    public bool PdfA { get; set; }
    public string PageRange { get; set; } = "";
    public bool ExportBookmarks { get; set; } = true;
}

/// <summary>One isolated headless Writer process per document; originals are opened from a private copy.</summary>
public sealed class ConversionService
{
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
        { ".doc", ".docx", ".rtf", ".odt" };
    private static readonly Regex PageRangePattern = new(@"^[1-9]\d*(?:-[1-9]\d*)?(?:,[1-9]\d*(?:-[1-9]\d*)?)*$",
        RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
    private readonly EngineManager _engine;
    private readonly string _jobsRoot = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordPdfStudio", "jobs");

    public ConversionService(EngineManager engine) => _engine = engine ?? throw new ArgumentNullException(nameof(engine));

    public static bool IsSupportedFile(string path) => SupportedExtensions.Contains(Path.GetExtension(path));

    public async Task ConvertAsync(string inputPath, string outputPath, ConversionOptions options,
        IProgress<EngineProgress>? progress = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(options);
        cancellationToken.ThrowIfCancellationRequested();
        inputPath = Path.GetFullPath(inputPath);
        outputPath = Path.GetFullPath(outputPath);
        if (!IsSupportedFile(inputPath))
            throw new ArgumentException("支持的文档格式为 DOC、DOCX、RTF 和 ODT。", nameof(inputPath));
        if (!File.Exists(inputPath))
            throw new FileNotFoundException("找不到源文档，文件可能已被移动或删除。", inputPath);
        if (new FileInfo(inputPath).Length == 0)
            throw new InvalidDataException("源文档为空，无法转换。");
        if (!Path.GetExtension(outputPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("输出文件扩展名必须是 .pdf。", nameof(outputPath));
        if (File.Exists(outputPath) || Directory.Exists(outputPath))
            throw new IOException("目标文件已存在。请选择其他文件名，或在应用中选择自动重命名。");
        if (options.Quality is < 1 or > 100)
            throw new ArgumentOutOfRangeException(nameof(options), "图片质量必须在 1 到 100 之间。");

        // Snapshot the settings so UI changes cannot affect a conversion already in progress.
        var filter = BuildPdfFilter(options.Quality, options.PdfA, options.ExportBookmarks,
            NormalizePageRange(options.PageRange));
        await ValidateInputAsync(inputPath, cancellationToken).ConfigureAwait(false);
        var executable = await _engine.EnsureReadyAsync(
            new InlineProgress(p => progress?.Report(new(p.Fraction * 0.2, p.Message))), cancellationToken)
            .ConfigureAwait(false);
        var outputDirectory = Path.GetDirectoryName(outputPath)!;
        Directory.CreateDirectory(outputDirectory);
        Directory.CreateDirectory(_jobsRoot);
        var jobDirectory = Path.Combine(_jobsRoot, Guid.NewGuid().ToString("N"));
        var profileDirectory = Path.Combine(jobDirectory, "profile");
        var resultDirectory = Path.Combine(jobDirectory, "result");
        var privateInput = Path.Combine(jobDirectory, "document" + Path.GetExtension(inputPath).ToLowerInvariant());
        var transactionalOutput = Path.Combine(outputDirectory, ".wordpdf-" + Guid.NewGuid().ToString("N") + ".tmp");

        try
        {
            Directory.CreateDirectory(profileDirectory);
            Directory.CreateDirectory(resultDirectory);
            await WriteSafeProfileAsync(profileDirectory, cancellationToken).ConfigureAwait(false);
            progress?.Report(new(0.22, "正在读取文档…"));
            await CopyFileAsync(inputPath, privateInput, cancellationToken).ConfigureAwait(false);

            var startInfo = new ProcessStartInfo
            {
                FileName = executable,
                WorkingDirectory = jobDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };
            var bundledFonts = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(executable)!,
                "..", "share", "fonts", "truetype"));
            if (Directory.Exists(bundledFonts))
                startInfo.Environment["SAL_FONTPATH"] = bundledFonts;
            startInfo.ArgumentList.Add("-env:UserInstallation=" + new Uri(profileDirectory + Path.DirectorySeparatorChar).AbsoluteUri);
            startInfo.ArgumentList.Add("--headless");
            startInfo.ArgumentList.Add("--nologo");
            startInfo.ArgumentList.Add("--nodefault");
            startInfo.ArgumentList.Add("--norestore");
            startInfo.ArgumentList.Add("--quickstart=no");
            startInfo.ArgumentList.Add("--convert-to");
            startInfo.ArgumentList.Add(filter);
            startInfo.ArgumentList.Add("--outdir");
            startInfo.ArgumentList.Add(resultDirectory);
            startInfo.ArgumentList.Add(privateInput);

            progress?.Report(new(0.35, "正在排版并转换为 PDF…"));
            await RunConverterAsync(startInfo, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();

            var convertedPdf = Path.Combine(resultDirectory, "document.pdf");
            await ValidatePdfAsync(convertedPdf, cancellationToken).ConfigureAwait(false);
            progress?.Report(new(0.9, "正在保存并检查 PDF…"));
            // Copy beside the destination, then rename on that same volume. Never replace an existing file.
            await CopyFileAsync(convertedPdf, transactionalOutput, cancellationToken).ConfigureAwait(false);
            await ValidatePdfAsync(transactionalOutput, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(transactionalOutput, outputPath, overwrite: false);
            progress?.Report(new(1, "转换完成"));
        }
        finally
        {
            try { if (File.Exists(transactionalOutput)) File.Delete(transactionalOutput); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            EngineManager.TryDeletePrivateDirectory(jobDirectory, _jobsRoot);
        }
    }

    public static string NormalizePageRange(string? pageRange)
    {
        if (string.IsNullOrWhiteSpace(pageRange))
            return "";
        var value = Regex.Replace(pageRange.Trim().Replace('，', ','), @"\s+", "");
        if (value.Length > 256 || !PageRangePattern.IsMatch(value))
            throw new ArgumentException("页码范围格式不正确，请使用 1-5,8,10-12；留空表示全部页。");
        foreach (var part in value.Split(','))
        {
            var endpoints = part.Split('-');
            if (!int.TryParse(endpoints[0], NumberStyles.None, CultureInfo.InvariantCulture, out var start) ||
                (endpoints.Length == 2 &&
                 (!int.TryParse(endpoints[1], NumberStyles.None, CultureInfo.InvariantCulture, out var end) || end < start)))
                throw new ArgumentException("页码必须为有效的正整数，范围结束页不能小于开始页。");
        }
        return value;
    }

    private static string BuildPdfFilter(int quality, bool pdfA, bool exportBookmarks, string pageRange)
    {
        var data = new Dictionary<string, object>
        {
            ["Quality"] = new { type = "long", value = quality.ToString(CultureInfo.InvariantCulture) },
            ["UseLosslessCompression"] = new { type = "boolean", value = "false" },
            ["SelectPdfVersion"] = new { type = "long", value = pdfA ? "2" : "0" },
            ["ExportBookmarks"] = new { type = "boolean", value = exportBookmarks ? "true" : "false" },
            ["UseTaggedPDF"] = new { type = "boolean", value = "true" },
            ["IsAddStream"] = new { type = "boolean", value = "false" },
            ["ExportFormFields"] = new { type = "boolean", value = "false" }
        };
        if (pageRange.Length > 0)
            data["PageRange"] = new { type = "string", value = pageRange };
        return "pdf:writer_pdf_Export:" + JsonSerializer.Serialize(data);
    }

    private static async Task WriteSafeProfileAsync(string profileDirectory, CancellationToken cancellationToken)
    {
        var userDirectory = Path.Combine(profileDirectory, "user");
        Directory.CreateDirectory(userDirectory);
        // These are official LibreOffice registry keys. This is a fresh, untrusted-document profile.
        const string registry = """
            <?xml version="1.0" encoding="UTF-8"?>
            <oor:items xmlns:oor="http://openoffice.org/2001/registry">
              <item oor:path="/org.openoffice.Office.Common/Security/Scripting">
                <prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop>
                <prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop>
                <prop oor:name="DisableActiveContent" oor:op="fuse"><value>true</value></prop>
                <prop oor:name="DisableOLEAutomation" oor:op="fuse"><value>true</value></prop>
                <prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse"><value>true</value></prop>
              </item>
              <item oor:path="/org.openoffice.Office.Writer/Content/Update">
                <prop oor:name="Link" oor:op="fuse"><value>2</value></prop>
                <prop oor:name="Field" oor:op="fuse"><value>false</value></prop>
                <prop oor:name="Chart" oor:op="fuse"><value>false</value></prop>
              </item>
              <item oor:path="/org.openoffice.Office.Common/Load">
                <prop oor:name="UserDefinedSettings" oor:op="fuse"><value>false</value></prop>
                <prop oor:name="DetectWebDAVRedirection" oor:op="fuse"><value>false</value></prop>
              </item>
              <item oor:path="/org.openoffice.Office.Writer/Filter/Import/VBA">
                <prop oor:name="Load" oor:op="fuse"><value>false</value></prop>
                <prop oor:name="Executable" oor:op="fuse"><value>false</value></prop>
                <prop oor:name="Save" oor:op="fuse"><value>false</value></prop>
              </item>
              <item oor:path="/org.openoffice.Office.Common/Misc">
                <prop oor:name="FirstRun" oor:op="fuse"><value>false</value></prop>
              </item>
              <item oor:path="/org.openoffice.Office.Jobs/Jobs/org.openoffice.Office.Jobs:Job['UpdateCheck']/Arguments">
                <prop oor:name="AutoCheckEnabled" oor:op="fuse"><value>false</value></prop>
                <prop oor:name="AutoDownloadEnabled" oor:op="fuse"><value>false</value></prop>
              </item>
              <item oor:path="/org.openoffice.Office.Update/Update">
                <prop oor:name="Enabled" oor:op="fuse"><value>false</value></prop>
              </item>
            </oor:items>
            """;
        await File.WriteAllTextAsync(Path.Combine(userDirectory, "registrymodifications.xcu"),
            registry, new UTF8Encoding(false), cancellationToken).ConfigureAwait(false);
    }

    private static async Task CopyFileAsync(string source, string destination, CancellationToken cancellationToken)
    {
        await using var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read,
            128 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan);
        await using var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None,
            128 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan);
        await input.CopyToAsync(output, 128 * 1024, cancellationToken).ConfigureAwait(false);
        await output.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private static async Task ValidateInputAsync(string path, CancellationToken cancellationToken)
    {
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
        var extension = Path.GetExtension(path).ToLowerInvariant();
        var header = new byte[16];
        var count = await stream.ReadAsync(header, cancellationToken).ConfigureAwait(false);
        var isCompoundDocument = count >= 8 && header.AsSpan(0, 8).SequenceEqual(
            new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 });
        var prefix = Encoding.ASCII.GetString(header, 0, count).TrimStart();
        var isRtf = prefix.StartsWith(@"{\rtf", StringComparison.Ordinal);
        if (extension == ".rtf")
        {
            if (!isRtf)
                throw new InvalidDataException("RTF 文档格式不正确，文件可能已损坏或使用了错误的扩展名。");
            return;
        }
        if (extension == ".doc")
        {
            // Binary Word 97–2003 and RTF documents saved by Word with the DOC extension.
            if (!isCompoundDocument && !isRtf)
                throw new InvalidDataException("DOC 文档格式不正确。请使用 Word 97–2003 文档或另存为 DOCX 后重试。");
            return;
        }
        if (isCompoundDocument)
            throw new InvalidDataException("此文档可能受密码保护。请先在文档编辑器中取消打开密码再转换。");
        stream.Position = 0;
        try
        {
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            if (extension == ".docx")
            {
                var document = archive.GetEntry("word/document.xml");
                if (archive.GetEntry("[Content_Types].xml") is null || archive.GetEntry("_rels/.rels") is null || document is null)
                    throw new InvalidDataException("DOCX 文档缺少必要内容，文件可能已损坏。");
                using var xmlStream = document.Open();
                using var reader = XmlReader.Create(xmlStream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    Async = true,
                    MaxCharactersInDocument = 100_000_000
                });
                await reader.MoveToContentAsync().ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.LocalName != "document" ||
                    reader.NamespaceURI is not ("http://schemas.openxmlformats.org/wordprocessingml/2006/main" or
                        "http://purl.oclc.org/ooxml/wordprocessingml/main"))
                    throw new InvalidDataException("DOCX 的正文结构无效，文件可能已损坏。");
            }
            else
            {
                var mimeType = archive.GetEntry("mimetype");
                if (mimeType is null || mimeType.Length > 128 || archive.GetEntry("content.xml") is null)
                    throw new InvalidDataException("ODT 文档缺少必要内容，文件可能已损坏。");
                using var mimeStream = mimeType.Open();
                using var reader = new StreamReader(mimeStream, Encoding.UTF8);
                if ((await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false)).Trim() !=
                    "application/vnd.oasis.opendocument.text")
                    throw new InvalidDataException("此文件不是有效的 ODT 文字文档。");
            }
        }
        catch (XmlException ex)
        {
            throw new InvalidDataException("文档正文格式无效，文件可能已损坏。", ex);
        }
        catch (InvalidDataException ex)
        {
            throw new InvalidDataException("无法读取文档内容，文件可能已损坏或使用了错误的扩展名。", ex);
        }
    }

    private static async Task ValidatePdfAsync(string path, CancellationToken cancellationToken)
    {
        if (!File.Exists(path) || new FileInfo(path).Length < 16)
            throw new InvalidDataException("未生成有效的 PDF。文档可能已损坏、受密码保护，或所选页码不在文档范围内。");
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            4096, FileOptions.Asynchronous);
        var header = new byte[5];
        await stream.ReadExactlyAsync(header, cancellationToken).ConfigureAwait(false);
        if (!header.AsSpan().SequenceEqual("%PDF-"u8))
            throw new InvalidDataException("转换结果不是有效的 PDF 文件。");
    }

    private static async Task RunConverterAsync(ProcessStartInfo startInfo, CancellationToken cancellationToken)
    {
        using var process = new Process { StartInfo = startInfo };
        try
        {
            if (!process.Start())
                throw new InvalidOperationException("无法启动转换引擎。");
        }
        catch (Win32Exception ex)
        {
            throw new InvalidOperationException("无法启动离线转换引擎。请检查程序完整性和系统权限。", ex);
        }

        // Kill-on-close keeps our own converter children from surviving an application crash.
        using var job = ProcessJob.TryAttach(process);
        using var readCancellation = new CancellationTokenSource();
        var stdout = ReadBoundedAsync(process.StandardOutput, readCancellation.Token);
        var stderr = ReadBoundedAsync(process.StandardError, readCancellation.Token);
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(180));
        using var combined = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeout.Token);
        try
        {
            await process.WaitForExitAsync(combined.Token).ConfigureAwait(false);
            var log = await stdout.WaitAsync(TimeSpan.FromSeconds(5)).ConfigureAwait(false);
            var error = await stderr.WaitAsync(TimeSpan.FromSeconds(5)).ConfigureAwait(false);
            if (process.ExitCode != 0)
                throw new InvalidOperationException("转换引擎未能完成任务。" + FormatError(error, log));
        }
        catch (OperationCanceledException)
        {
            KillOwnProcess(process, job);
            await WaitForStopAsync(process).ConfigureAwait(false);
            if (cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(cancellationToken);
            throw new TimeoutException("文档转换超过 3 分钟。请检查文档是否过大、损坏或受密码保护。");
        }
        catch (TimeoutException)
        {
            KillOwnProcess(process, job);
            await WaitForStopAsync(process).ConfigureAwait(false);
            throw new TimeoutException("转换引擎未正常结束。请重试或检查文档。");
        }
        finally
        {
            readCancellation.Cancel();
            // Observe reader completion/failure without letting inherited handles keep the UI waiting.
            try { await Task.WhenAll(stdout, stderr).WaitAsync(TimeSpan.FromSeconds(2)).ConfigureAwait(false); }
            catch (OperationCanceledException) { }
            catch (IOException) { }
            catch (TimeoutException) { }
        }
    }

    private static async Task<string> ReadBoundedAsync(StreamReader reader, CancellationToken cancellationToken)
    {
        const int maximumCharacters = 8192;
        var result = new StringBuilder();
        var buffer = new char[1024];
        int count;
        while ((count = await reader.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false)) > 0)
        {
            result.Append(buffer, 0, count);
            if (result.Length > maximumCharacters)
                result.Remove(0, result.Length - maximumCharacters);
        }
        return result.ToString();
    }

    private static string FormatError(string error, string log)
    {
        var message = string.IsNullOrWhiteSpace(error) ? log : error;
        message = message.Trim();
        if (message.Length > 1000)
            message = message[..1000] + "…";
        return message.Length == 0 ? "" : "\n" + message;
    }

    private static void KillOwnProcess(Process process, ProcessJob? job)
    {
        job?.Terminate();
        try { if (!process.HasExited) process.Kill(entireProcessTree: true); }
        catch (InvalidOperationException) { }
        catch (Win32Exception) { }
    }

    private static async Task WaitForStopAsync(Process process)
    {
        try { await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10)).ConfigureAwait(false); }
        catch (InvalidOperationException) { }
        catch (TimeoutException) { }
    }

    private sealed class InlineProgress(Action<EngineProgress> action) : IProgress<EngineProgress>
    {
        public void Report(EngineProgress value) => action(value);
    }

    private sealed class ProcessJob : IDisposable
    {
        private IntPtr _handle;
        private ProcessJob(IntPtr handle) => _handle = handle;

        public static ProcessJob? TryAttach(Process process)
        {
            if (!OperatingSystem.IsWindows())
                return null;
            var handle = CreateJobObject(IntPtr.Zero, null);
            if (handle == IntPtr.Zero)
                return null;
            var information = new JobExtendedLimitInformation();
            information.BasicLimitInformation.LimitFlags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            if (!SetInformationJobObject(handle, 9, ref information,
                    (uint)Marshal.SizeOf<JobExtendedLimitInformation>()) ||
                !AssignProcessToJobObject(handle, process.Handle))
            {
                CloseHandle(handle);
                return null;
            }
            return new ProcessJob(handle);
        }

        public void Terminate()
        {
            if (_handle != IntPtr.Zero)
                TerminateJobObject(_handle, 1);
        }

        public void Dispose()
        {
            if (_handle != IntPtr.Zero)
            {
                CloseHandle(_handle);
                _handle = IntPtr.Zero;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JobBasicLimitInformation
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IoCounters
        {
            public ulong ReadOperationCount;
            public ulong WriteOperationCount;
            public ulong OtherOperationCount;
            public ulong ReadTransferCount;
            public ulong WriteTransferCount;
            public ulong OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JobExtendedLimitInformation
        {
            public JobBasicLimitInformation BasicLimitInformation;
            public IoCounters IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateJobObject(IntPtr attributes, string? name);
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetInformationJobObject(IntPtr job, int informationClass,
            ref JobExtendedLimitInformation information, uint informationLength);
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool TerminateJobObject(IntPtr job, uint exitCode);
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr handle);
    }
}
