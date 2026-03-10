using System.Windows;
using WindowPilot.Diagnostics;

namespace WindowPilot;

public partial class App : Application
{
    // 单实例互斥锁，防止程序重复启动
    private static Mutex? _singleInstanceMutex;

    /// <summary>
    /// 应用程序启动入口，初始化日志系统并执行单实例检查。
    /// </summary>
    /// <param name="e">启动事件参数，包含命令行参数列表。</param>
    protected override void OnStartup(StartupEventArgs e)
    {
        // 日志系统必须最先初始化，后续所有模块依赖它
        Logger.Initialize(
            minConsoleLevel: LogLevel.Debug,
            minFileLevel:    LogLevel.Trace);

        Logger.Separator("App Startup");
        Logger.Info("WindowPilot 启动中…", "App");
        Logger.Debug($"工作目录: {Environment.CurrentDirectory}", "App");
        Logger.Debug($"可执行路径: {AppContext.BaseDirectory}", "App");
        Logger.Debug($"命令行参数: {string.Join(" ", e.Args)}", "App");

        // 尝试获取命名互斥锁，isNewInstance 为 false 表示已有实例在运行
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

    /// <summary>
    /// 应用程序退出时释放资源并关闭日志系统。
    /// </summary>
    /// <param name="e">退出事件参数，包含退出代码。</param>
    protected override void OnExit(ExitEventArgs e)
    {
        Logger.Separator("App Shutdown");
        Logger.Info($"程序退出，ExitCode = {e.ApplicationExitCode}", "App");

        // 先释放互斥锁，允许新实例启动
        _singleInstanceMutex?.ReleaseMutex();
        _singleInstanceMutex?.Dispose();
        Logger.Debug("单实例互斥锁已释放。", "App");

        base.OnExit(e);

        // 日志系统最后关闭，确保所有挂起日志都写入文件
        Logger.Shutdown();
    }
}