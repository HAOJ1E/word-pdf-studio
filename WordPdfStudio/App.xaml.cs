using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace WordPdfStudio;

public partial class App : Application
{
    public static bool IsScreenshotMode { get; private set; }
    public static bool IsAutomationMode { get; private set; }

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        var args = e.Args;
        var cli = args.Length > 0 && args[0].Equals("--cli", StringComparison.OrdinalIgnoreCase);
        IsScreenshotMode = args.Any(value => value.Equals("--screenshot", StringComparison.OrdinalIgnoreCase));
        IsAutomationMode = cli || IsScreenshotMode;
        DispatcherUnhandledException += HandleDispatcherException;
        if (IsAutomationMode)
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
        try
        {
            if (cli)
            {
                var exitCode = await HeadlessRunner.RunAsync(args.Skip(1).ToArray());
                Shutdown(exitCode);
                return;
            }

            var window = new MainWindow();
            MainWindow = window;
            if (IsScreenshotMode)
            {
                await CaptureScreenshotAsync(window, args);
                window.Close();
                Shutdown(0);
                return;
            }

            var files = args.Where(File.Exists).ToArray();
            if (files.Length > 0)
                window.AddPaths(files);
            window.Show();
        }
        catch (Exception exception)
        {
            LogFailure(exception, "startup");
            if (IsAutomationMode)
            {
                var report = args.LastOrDefault(value => value.StartsWith("--report=", StringComparison.OrdinalIgnoreCase));
                try
                {
                    await HeadlessRunner.WriteAutomationReportAsync(report is null ? null : report[9..],
                        cli ? "startup" : "screenshot", "", false, "启动失败：" + exception.Message, 0);
                }
                catch (Exception reportException) { LogFailure(reportException, "startup-report"); }
                Shutdown(1);
            }
            else
            {
                MessageBox.Show("程序启动失败：\n" + exception.Message, "文页 PDF", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown(1);
            }
        }
    }

    private static async Task CaptureScreenshotAsync(MainWindow window, IReadOnlyList<string> args)
    {
        var timer = Stopwatch.StartNew();
        var index = args.ToList().FindIndex(value => value.Equals("--screenshot", StringComparison.OrdinalIgnoreCase));
        if (index < 0 || index + 1 >= args.Count || !Path.IsPathFullyQualified(args[index + 1]))
            throw new ArgumentException("--screenshot 后必须指定 PNG 文件的绝对路径。");
        var destination = Path.GetFullPath(args[index + 1]);
        if (!Path.GetExtension(destination).Equals(".png", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("截图文件的扩展名必须是 .png。");
        if (File.Exists(destination) || Directory.Exists(destination))
            throw new IOException("截图文件已存在，请使用新的文件名。");

        window.WindowStartupLocation = WindowStartupLocation.Manual;
        window.Left = -10000;
        window.Top = 0;
        window.ShowActivated = false;
        window.ShowInTaskbar = false;
        var loaded = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        window.Loaded += (_, _) => loaded.TrySetResult();
        window.Show();
        await loaded.Task.WaitAsync(TimeSpan.FromSeconds(15));

        var uiPageOption = args.LastOrDefault(value => value.StartsWith("--ui-page=", StringComparison.OrdinalIgnoreCase));
        var page = uiPageOption is null ? "convert" : uiPageOption[10..].ToLowerInvariant();
        if (page is not ("convert" or "tools" or "history" or "about"))
            throw new ArgumentException("截图页面可选 convert、tools、history 或 about。");
        window.Navigate(page);
        var addFiles = args.Where(value => value.StartsWith("--add=", StringComparison.OrdinalIgnoreCase))
            .Select(value => value[6..]).ToArray();
        if (addFiles.Length > 0)
        {
            if (addFiles.Any(path => !File.Exists(path)))
                throw new FileNotFoundException("截图预览的 --add 文件不存在。");
            window.AddPaths(addFiles);
        }

        await window.Dispatcher.InvokeAsync(() => window.UpdateLayout(), DispatcherPriority.Loaded);
        await Task.Delay(200);
        await window.Dispatcher.InvokeAsync(() => window.UpdateLayout(), DispatcherPriority.Render);
        var root = window.RootGrid;
        root.Background ??= window.Background;
        var width = root.ActualWidth;
        var height = root.ActualHeight;
        if (width <= 0 || height <= 0)
            throw new InvalidOperationException("窗口尚未完成布局，无法截图。");
        root.Measure(new Size(width, height));
        root.Arrange(new Rect(0, 0, width, height));
        root.UpdateLayout();
        var bitmap = new RenderTargetBitmap((int)Math.Ceiling(width), (int)Math.Ceiling(height), 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(root);
        bitmap.Freeze();
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        using (var stream = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            encoder.Save(stream);

        var report = args.LastOrDefault(value => value.StartsWith("--report=", StringComparison.OrdinalIgnoreCase));
        await HeadlessRunner.WriteAutomationReportAsync(report is null ? null : report[9..], "screenshot",
            destination, true, "", timer.Elapsed.TotalSeconds);
    }

    private void HandleDispatcherException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        LogFailure(e.Exception, "dispatcher");
        e.Handled = true;
        if (IsAutomationMode)
        {
            Shutdown(1);
            return;
        }
        MessageBox.Show("操作未能完成：\n" + e.Exception.Message + "\n\n请检查文件或保存位置后重试。",
            "文页 PDF", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    internal static void LogFailure(Exception exception, string context)
    {
        try
        {
            // Exception messages and command arguments may contain credentials. Log only safe diagnostics.
            var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WordPdfStudio", "logs");
            Directory.CreateDirectory(directory);
            File.AppendAllText(Path.Combine(directory, "errors.log"),
                $"{DateTimeOffset.UtcNow:O} {context} {exception.GetType().FullName}\n{exception.StackTrace}\n");
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
