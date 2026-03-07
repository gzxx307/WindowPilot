using WindowPilot.Diagnostics;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 监听系统窗口事件（移动开始/结束、窗口销毁、前景变化等）
/// </summary>
public class WinEventHookService : IDisposable
{
    private const string Cat = "WinEventHook";

    private readonly List<IntPtr> _hookHandles = new();
    private readonly NativeMethods.WinEventDelegate _winEventDelegate;

    // ── 高频事件节流（LocationChange 非常密集，只在 Debug 模式下完整输出）──
    private int _locationChangeCount;
    private int _moveSizeCount;

    // ── 事件 ──
    public event Action<IntPtr>? WindowMoveSizeStarted;
    public event Action<IntPtr>? WindowMoveSizeEnded;
    public event Action<IntPtr>? WindowDestroyed;
    public event Action<IntPtr>? WindowMinimized;
    public event Action<IntPtr>? WindowRestored;
    public event Action<IntPtr>? WindowForegroundChanged;
    public event Action<IntPtr>? WindowLocationChanged;
    public event Action<IntPtr>? WindowTitleChanged;

    public WinEventHookService()
    {
        Logger.Debug("WinEventHookService 构造中…", Cat);
        // 必须持有 delegate，防止被 GC 回收
        _winEventDelegate = WinEventCallback;
        Logger.Trace("WinEventDelegate 已创建并固定引用。", Cat);
    }

    /// <summary>
    /// 启动所有事件监听
    /// </summary>
    public void Start()
    {
        Logger.Info("安装 WinEvent 钩子…", Cat);

        InstallHook(NativeConstants.EVENT_SYSTEM_MOVESIZESTART,
                    NativeConstants.EVENT_SYSTEM_MOVESIZEEND,
                    "MoveSize Start/End");

        InstallHook(NativeConstants.EVENT_OBJECT_DESTROY,
                    NativeConstants.EVENT_OBJECT_DESTROY,
                    "Object Destroy");

        InstallHook(NativeConstants.EVENT_SYSTEM_MINIMIZESTART,
                    NativeConstants.EVENT_SYSTEM_MINIMIZEEND,
                    "Minimize Start/End");

        InstallHook(NativeConstants.EVENT_SYSTEM_FOREGROUND,
                    NativeConstants.EVENT_SYSTEM_FOREGROUND,
                    "Foreground Change");

        InstallHook(NativeConstants.EVENT_OBJECT_LOCATIONCHANGE,
                    NativeConstants.EVENT_OBJECT_LOCATIONCHANGE,
                    "Location Change");

        InstallHook(NativeConstants.EVENT_OBJECT_NAMECHANGE,
                    NativeConstants.EVENT_OBJECT_NAMECHANGE,
                    "Name/Title Change");

        Logger.Info($"WinEvent 钩子安装完毕，共 {_hookHandles.Count} 个有效钩子。", Cat);
    }

    private void InstallHook(uint eventMin, uint eventMax, string description)
    {
        Logger.Trace($"InstallHook: [{description}] eventMin=0x{eventMin:X4} eventMax=0x{eventMax:X4}", Cat);

        var handle = NativeMethods.SetWinEventHook(
            eventMin, eventMax,
            IntPtr.Zero,
            _winEventDelegate,
            0, 0,
            NativeConstants.WINEVENT_OUTOFCONTEXT | NativeConstants.WINEVENT_SKIPOWNPROCESS);

        if (handle != IntPtr.Zero)
        {
            _hookHandles.Add(handle);
            Logger.Debug($"  ✓ 钩子安装成功: [{description}] handle=0x{handle:X}", Cat);
        }
        else
        {
            Logger.Error($"  ✗ 钩子安装失败: [{description}]", Cat);
        }
    }

    private void WinEventCallback(
        IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        // 只关注窗口对象本身
        if (idObject != NativeConstants.OBJID_WINDOW || idChild != 0)
            return;

        if (hwnd == IntPtr.Zero)
            return;

        switch (eventType)
        {
            case NativeConstants.EVENT_SYSTEM_MOVESIZESTART:
                int msCount = Interlocked.Increment(ref _moveSizeCount);
                Logger.Debug($"EVENT_SYSTEM_MOVESIZESTART  hwnd=0x{hwnd:X}  (第{msCount}次)", Cat);
                WindowMoveSizeStarted?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_SYSTEM_MOVESIZEEND:
                Logger.Debug($"EVENT_SYSTEM_MOVESIZEEND    hwnd=0x{hwnd:X}", Cat);
                WindowMoveSizeEnded?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_OBJECT_DESTROY:
                Logger.Trace($"EVENT_OBJECT_DESTROY        hwnd=0x{hwnd:X}", Cat);
                WindowDestroyed?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_SYSTEM_MINIMIZESTART:
                Logger.Debug($"EVENT_SYSTEM_MINIMIZESTART  hwnd=0x{hwnd:X}", Cat);
                WindowMinimized?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_SYSTEM_MINIMIZEEND:
                Logger.Debug($"EVENT_SYSTEM_MINIMIZEEND    hwnd=0x{hwnd:X}", Cat);
                WindowRestored?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_SYSTEM_FOREGROUND:
                Logger.Trace($"EVENT_SYSTEM_FOREGROUND     hwnd=0x{hwnd:X}", Cat);
                WindowForegroundChanged?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_OBJECT_LOCATIONCHANGE:
                // LocationChange 非常密集（鼠标拖动时每帧触发），
                // 只写文件 Trace，不输出到控制台以避免刷屏
                int lcCount = Interlocked.Increment(ref _locationChangeCount);
                if (lcCount % 120 == 0) // 每 120 次输出一条控制台摘要
                    Logger.Debug($"EVENT_OBJECT_LOCATIONCHANGE hwnd=0x{hwnd:X}  (累计 {lcCount} 次)", Cat);
                else
                    Logger.Trace($"EVENT_OBJECT_LOCATIONCHANGE hwnd=0x{hwnd:X}  (累计 {lcCount} 次)", Cat);
                WindowLocationChanged?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_OBJECT_NAMECHANGE:
                Logger.Trace($"EVENT_OBJECT_NAMECHANGE     hwnd=0x{hwnd:X}", Cat);
                WindowTitleChanged?.Invoke(hwnd);
                break;

            default:
                Logger.Trace($"未处理事件 0x{eventType:X4}  hwnd=0x{hwnd:X}", Cat);
                break;
        }
    }

    /// <summary>
    /// 停止所有监听
    /// </summary>
    public void Stop()
    {
        Logger.Info($"卸载 {_hookHandles.Count} 个 WinEvent 钩子…", Cat);
        foreach (var handle in _hookHandles)
        {
            bool ok = NativeMethods.UnhookWinEvent(handle);
            Logger.Debug($"  UnhookWinEvent(0x{handle:X}) = {ok}", Cat);
        }
        _hookHandles.Clear();
        Logger.Info("所有 WinEvent 钩子已卸载。", Cat);
        Logger.Debug($"统计 — MoveSize触发: {_moveSizeCount}次  LocationChange触发: {_locationChangeCount}次", Cat);
    }

    public void Dispose()
    {
        Logger.Debug("WinEventHookService.Dispose()", Cat);
        Stop();
        GC.SuppressFinalize(this);
    }
}