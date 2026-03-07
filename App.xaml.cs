using System.Windows;
using WindowPilot.Diagnostics;

namespace WindowPilot;

public partial class App : Application
{
    private static Mutex? _singleInstanceMutex;

    protected override void OnStartup(StartupEventArgs e)
    {
        // ── 初始化日志系统（最先执行）──────────────────────────
        Logger.Initialize(
            minConsoleLevel: LogLevel.Debug,
            minFileLevel:    LogLevel.Trace);

        Logger.Separator("App Startup");
        Logger.Info("WindowPilot 启动中…", "App");
        Logger.Debug($"工作目录: {Environment.CurrentDirectory}", "App");
        Logger.Debug($"可执行路径: {AppContext.BaseDirectory}", "App");
        Logger.Debug($"命令行参数: {string.Join(" ", e.Args)}", "App");

        // 单实例检查
        _singleInstanceMutex = new Mutex(true, "WindowPilot_SingleInstance", out bool isNewInstance);
        if (!isNewInstance)
        {
            Logger.Warning("检测到已有实例在运行，本次启动将被终止。", "App");
            MessageBox.Show("WindowPilot 已经在运行中。", "WindowPilot",
                MessageBoxButton.OK, MessageBoxImage.Information);
            Logger.Shutdown();
            Shutdown();
            return;
        }

        Logger.Info("单实例互斥锁获取成功，继续启动。", "App");
        Logger.Debug($"日志文件: {Logger.CurrentLogFilePath}", "App");

        base.OnStartup(e);

        Logger.Info("App.OnStartup 完成。", "App");
    }

    protected override void OnExit(ExitEventArgs e)
    {
        Logger.Separator("App Shutdown");
        Logger.Info($"程序退出，ExitCode = {e.ApplicationExitCode}", "App");

        _singleInstanceMutex?.ReleaseMutex();
        _singleInstanceMutex?.Dispose();
        Logger.Debug("单实例互斥锁已释放。", "App");

        base.OnExit(e);

        // 关闭日志系统（最后执行，确保所有日志都写入文件）
        Logger.Shutdown();
    }
}