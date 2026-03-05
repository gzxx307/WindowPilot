using System.Windows;

namespace WindowPilot;

public partial class App : Application
{
    private static Mutex? _singleInstanceMutex;

    protected override void OnStartup(StartupEventArgs e)
    {
        // 单实例检查
        _singleInstanceMutex = new Mutex(true, "WindowPilot_SingleInstance", out bool isNewInstance);
        if (!isNewInstance)
        {
            MessageBox.Show("WindowPilot 已经在运行中。", "WindowPilot",
                MessageBoxButton.OK, MessageBoxImage.Information);
            Shutdown();
            return;
        }

        base.OnStartup(e);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _singleInstanceMutex?.ReleaseMutex();
        _singleInstanceMutex?.Dispose();
        base.OnExit(e);
    }
}
