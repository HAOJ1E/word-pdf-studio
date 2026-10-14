using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace WordPdfStudio.Services;

public sealed record EngineProgress(double Fraction, string Message);

/// <summary>Prepares the private, embedded converter without installing Office or changing its profile.</summary>
public sealed class EngineManager
{
    public const string EngineVersion = "26.2.6";
    private const string ResourceName = "WordPdfStudio.Engine.libreoffice.zip";
    private const string CompletionMarker = ".wordpdf-engine-complete";
    private const long MinimumFreeSpace = 1_400_000_000;
    private readonly SemaphoreSlim _gate = new(1, 1);

    public string CacheDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "WordPdfStudio", "engine", EngineVersion);

    public bool IsReady
    {
        get
        {
            try { return TryGetOverride(out _) || TryGetCachedExecutable(out _); }
            catch (ArgumentException) { return false; }
            catch (InvalidOperationException) { return false; }
            catch (IOException) { return false; }
            catch (UnauthorizedAccessException) { return false; }
        }
    }

    public async Task<string> EnsureReadyAsync(IProgress<EngineProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (TryGetOverride(out var customExecutable))
        {
            progress?.Report(new(1, "转换引擎已就绪"));
            return customExecutable!;
        }

        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (TryGetCachedExecutable(out var cachedExecutable))
            {
                progress?.Report(new(1, "转换引擎已就绪"));
                return cachedExecutable!;
            }

            var engineRoot = Directory.GetParent(CacheDirectory)!.FullName;
            Directory.CreateDirectory(engineRoot);
            // An exclusive file handle survives await/thread switches, unlike a thread-affine Mutex.
            using var extractionLock = await AcquireLockAsync(
                Path.Combine(engineRoot, "extract-" + EngineVersion + ".lock"),
                progress, cancellationToken).ConfigureAwait(false);
            if (TryGetCachedExecutable(out cachedExecutable))
            {
                progress?.Report(new(1, "转换引擎已就绪"));
                return cachedExecutable!;
            }

            // A hard application stop may leave a partial staging tree. The lock proves none is active now.
            foreach (var abandoned in Directory.EnumerateDirectories(engineRoot,
                         ".extract-" + EngineVersion + "-*", SearchOption.TopDirectoryOnly))
                TryDeletePrivateDirectory(abandoned, engineRoot);

            using var archiveStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName)
                ?? throw new InvalidOperationException("程序未包含转换引擎。请使用完整的便携版程序。");
            using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read, leaveOpen: false);
            var uncompressedBytes = archive.Entries.Sum(entry => entry.Length);
            if (uncompressedBytes <= 0)
                throw new InvalidDataException("内置转换引擎为空或已损坏。");
            EnsureFreeSpace(engineRoot, Math.Max(MinimumFreeSpace,
                checked(uncompressedBytes + 200_000_000)));

            var staging = Path.Combine(engineRoot, ".extract-" + EngineVersion + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(staging);
            try
            {
                progress?.Report(new(0, "首次运行：正在准备离线转换引擎…"));
                await ExtractAsync(archive, staging, uncompressedBytes, progress, cancellationToken)
                    .ConfigureAwait(false);
                var executable = FindExecutable(staging);
                if (executable is null)
                    throw new InvalidDataException("内置转换引擎文件不完整。请重新获取完整程序。");

                cancellationToken.ThrowIfCancellationRequested();
                await File.WriteAllTextAsync(Path.Combine(staging, CompletionMarker),
                    EngineVersion, cancellationToken).ConfigureAwait(false);
                // The complete tree appears at its final location in one rename, while holding the lock.
                if (Directory.Exists(CacheDirectory))
                    DeletePrivateDirectory(CacheDirectory, engineRoot);
                Directory.Move(staging, CacheDirectory);
                progress?.Report(new(1, "离线转换引擎准备完成"));
                return FindExecutable(CacheDirectory)!;
            }
            finally
            {
                TryDeletePrivateDirectory(staging, engineRoot);
            }
        }
        finally
        {
            _gate.Release();
        }
    }

    private bool TryGetOverride(out string? executable)
    {
        executable = null;
        var value = Environment.GetEnvironmentVariable("WORDPDF_ENGINE_PATH");
        if (string.IsNullOrWhiteSpace(value))
            return false;

        var fullPath = Path.GetFullPath(value.Trim().Trim('"'));
        if (Directory.Exists(fullPath))
        {
            executable = FindExecutable(fullPath) ?? FindExecutable(Directory.GetParent(fullPath)?.FullName ?? fullPath);
        }
        else if (File.Exists(fullPath) &&
                 (Path.GetFileName(fullPath).Equals("soffice.com", StringComparison.OrdinalIgnoreCase) ||
                  Path.GetFileName(fullPath).Equals("soffice.exe", StringComparison.OrdinalIgnoreCase)))
        {
            var programDirectory = Path.GetDirectoryName(fullPath)!;
            var preferred = Path.Combine(programDirectory, "soffice.com");
            if (HasRuntimeFiles(programDirectory))
                executable = File.Exists(preferred) ? preferred : fullPath;
        }

        if (executable is null)
            throw new InvalidOperationException("开发转换引擎路径无效：WORDPDF_ENGINE_PATH 必须指向完整 LibreOffice 目录或 soffice 程序。");
        return true;
    }

    private bool TryGetCachedExecutable(out string? executable)
    {
        executable = null;
        try
        {
            var marker = Path.Combine(CacheDirectory, CompletionMarker);
            if (!File.Exists(marker) || File.ReadAllText(marker).Trim() != EngineVersion)
                return false;
            executable = FindExecutable(CacheDirectory);
            return executable is not null;
        }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
    }

    private static string? FindExecutable(string root)
    {
        var programDirectory = Path.Combine(root, "program");
        if (!HasRuntimeFiles(programDirectory))
            return null;
        var consoleExecutable = Path.Combine(programDirectory, "soffice.com");
        if (File.Exists(consoleExecutable))
            return consoleExecutable;
        var executable = Path.Combine(programDirectory, "soffice.exe");
        return File.Exists(executable) ? executable : null;
    }

    private static bool HasRuntimeFiles(string programDirectory) =>
        File.Exists(Path.Combine(programDirectory, "soffice.bin")) &&
        File.Exists(Path.Combine(programDirectory, "bootstrap.ini"));

    private static async Task<FileStream> AcquireLockAsync(string lockPath,
        IProgress<EngineProgress>? progress, CancellationToken cancellationToken)
    {
        var firstAttempt = true;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite,
                    FileShare.None, 1, FileOptions.None);
            }
            catch (IOException ex) when ((ex.HResult & 0xffff) is 32 or 33)
            {
                if (firstAttempt)
                {
                    progress?.Report(new(0, "另一个窗口正在准备引擎，等待完成…"));
                    firstAttempt = false;
                }
                await Task.Delay(250, cancellationToken).ConfigureAwait(false);
            }
        }
    }

    private static void EnsureFreeSpace(string directory, long requiredBytes)
    {
        var drive = new DriveInfo(Path.GetPathRoot(Path.GetFullPath(directory))!);
        if (drive.IsReady && drive.AvailableFreeSpace < requiredBytes)
            throw new IOException($"系统盘可用空间不足。首次运行至少需要 {requiredBytes / 1_000_000_000d:0.0} GB 可用空间。");
    }

    private static async Task ExtractAsync(ZipArchive archive, string staging, long totalBytes,
        IProgress<EngineProgress>? progress, CancellationToken cancellationToken)
    {
        var prefix = Path.GetFullPath(staging) + Path.DirectorySeparatorChar;
        var buffer = new byte[128 * 1024];
        long completedBytes = 0;
        var lastReport = Stopwatch.StartNew();
        foreach (var entry in archive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // Reject symlinks and absolute/traversing entries before touching the file system.
            if (((entry.ExternalAttributes >> 16) & 0xf000) == 0xa000 ||
                Path.IsPathRooted(entry.FullName) || entry.FullName.Contains(':'))
                throw new InvalidDataException("内置转换引擎包含不安全的文件路径。");
            var relativePath = entry.FullName.Replace('/', Path.DirectorySeparatorChar);
            var destination = Path.GetFullPath(Path.Combine(staging, relativePath));
            if (!destination.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("内置转换引擎包含越界文件路径。");
            if (string.IsNullOrEmpty(entry.Name))
            {
                Directory.CreateDirectory(destination);
                continue;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            using var source = entry.Open();
            await using var target = new FileStream(destination, FileMode.CreateNew,
                FileAccess.Write, FileShare.None, buffer.Length, FileOptions.Asynchronous | FileOptions.SequentialScan);
            int count;
            while ((count = await source.ReadAsync(buffer, cancellationToken).ConfigureAwait(false)) > 0)
            {
                await target.WriteAsync(buffer.AsMemory(0, count), cancellationToken).ConfigureAwait(false);
                completedBytes += count;
                if (lastReport.ElapsedMilliseconds >= 150)
                {
                    var fraction = Math.Clamp((double)completedBytes / totalBytes, 0, 0.99);
                    progress?.Report(new(fraction, $"首次运行：正在准备离线转换引擎 {fraction:P0}"));
                    lastReport.Restart();
                }
            }
        }
    }

    internal static void TryDeletePrivateDirectory(string path, string parent)
    {
        try { DeletePrivateDirectory(path, parent); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    internal static void DeletePrivateDirectory(string path, string parent)
    {
        var fullPath = Path.GetFullPath(path);
        var prefix = Path.GetFullPath(parent).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!fullPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("拒绝清理应用目录之外的文件。");
        if (!Directory.Exists(fullPath))
            return;
        // Delete reparse points themselves, never recurse through them.
        if ((File.GetAttributes(fullPath) & FileAttributes.ReparsePoint) != 0)
        {
            Directory.Delete(fullPath);
            return;
        }
        foreach (var directory in Directory.EnumerateDirectories(fullPath))
            DeletePrivateDirectory(directory, parent);
        foreach (var file in Directory.EnumerateFiles(fullPath))
        {
            File.SetAttributes(file, FileAttributes.Normal);
            File.Delete(file);
        }
        Directory.Delete(fullPath);
    }
}
