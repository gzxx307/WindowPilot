using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 监听系统窗口事件（移动开始/结束、窗口销毁、前景变化等）
/// </summary>
public class WinEventHookService : IDisposable
{
    private readonly List<IntPtr> _hookHandles = new();
    private readonly NativeMethods.WinEventDelegate _winEventDelegate;

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
        // 必须持有delegate，防止被GC回收
        _winEventDelegate = WinEventCallback;
    }

    /// <summary>
    /// 启动所有事件监听
    /// </summary>
    public void Start()
    {
        // 窗口拖动开始/结束
        InstallHook(NativeConstants.EVENT_SYSTEM_MOVESIZESTART, NativeConstants.EVENT_SYSTEM_MOVESIZEEND);

        // 窗口销毁
        InstallHook(NativeConstants.EVENT_OBJECT_DESTROY, NativeConstants.EVENT_OBJECT_DESTROY);

        // 最小化/还原
        InstallHook(NativeConstants.EVENT_SYSTEM_MINIMIZESTART, NativeConstants.EVENT_SYSTEM_MINIMIZEEND);

        // 前景窗口变化
        InstallHook(NativeConstants.EVENT_SYSTEM_FOREGROUND, NativeConstants.EVENT_SYSTEM_FOREGROUND);

        // 窗口位置变化
        InstallHook(NativeConstants.EVENT_OBJECT_LOCATIONCHANGE, NativeConstants.EVENT_OBJECT_LOCATIONCHANGE);

        // 窗口标题变化
        InstallHook(NativeConstants.EVENT_OBJECT_NAMECHANGE, NativeConstants.EVENT_OBJECT_NAMECHANGE);
    }

    private void InstallHook(uint eventMin, uint eventMax)
    {
        var handle = NativeMethods.SetWinEventHook(
            eventMin, eventMax,
            IntPtr.Zero,
            _winEventDelegate,
            0, 0, // 所有进程、所有线程
            NativeConstants.WINEVENT_OUTOFCONTEXT | NativeConstants.WINEVENT_SKIPOWNPROCESS);

        if (handle != IntPtr.Zero)
            _hookHandles.Add(handle);
    }

    private void WinEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
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
                WindowMoveSizeStarted?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_SYSTEM_MOVESIZEEND:
                WindowMoveSizeEnded?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_OBJECT_DESTROY:
                WindowDestroyed?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_SYSTEM_MINIMIZESTART:
                WindowMinimized?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_SYSTEM_MINIMIZEEND:
                WindowRestored?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_SYSTEM_FOREGROUND:
                WindowForegroundChanged?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_OBJECT_LOCATIONCHANGE:
                WindowLocationChanged?.Invoke(hwnd);
                break;

            case NativeConstants.EVENT_OBJECT_NAMECHANGE:
                WindowTitleChanged?.Invoke(hwnd);
                break;
        }
    }

    /// <summary>
    /// 停止所有监听
    /// </summary>
    public void Stop()
    {
        foreach (var handle in _hookHandles)
        {
            NativeMethods.UnhookWinEvent(handle);
        }
        _hookHandles.Clear();
    }

    public void Dispose()
    {
        Stop();
        GC.SuppressFinalize(this);
    }
}
