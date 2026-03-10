using WindowPilot.Diagnostics;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 监听系统 WinEvent 事件，将底层回调转换为强类型 C# 事件对外暴露。
/// 调用 <see cref="Start"/> 安装钩子，程序退出时通过 <see cref="Dispose"/> 卸载。
/// </summary>
public class WinEventHookService : IDisposable
{
    private const string Cat = "WinEventHook";

    // 所有已安装钩子的句柄列表，卸载时逐一释放
    private readonly List<IntPtr> _hookHandles = new();
    // 钩子委托实例必须持有强引用，防止被 GC 回收导致回调崩溃
    private readonly NativeMethods.WinEventDelegate _winEventDelegate;

    // 高频事件节流计数器
    private int _locationChangeCount; // LocationChange 事件触发次数
    private int _moveSizeCount;       // MoveSize 事件触发次数

    // 对外暴露的事件
    public event Action<IntPtr>? WindowMoveSizeStarted;    // 窗口开始移动或调整大小
    public event Action<IntPtr>? WindowMoveSizeEnded;      // 窗口完成移动或调整大小
    public event Action<IntPtr>? WindowDestroyed;          // 窗口被销毁
    public event Action<IntPtr>? WindowMinimized;          // 窗口开始最小化
    public event Action<IntPtr>? WindowRestored;           // 窗口从最小化还原
    public event Action<IntPtr>? WindowForegroundChanged;  // 前景窗口发生变化
    public event Action<IntPtr>? WindowLocationChanged;    // 窗口位置或大小发生变化（高频）
    public event Action<IntPtr>? WindowTitleChanged;       // 窗口标题文字发生变化

    // 构造函数，创建并固定委托引用
    public WinEventHookService()
    {
        Logger.Debug("WinEventHookService 构造中…", Cat);
        // 必须在此保存委托引用，若仅作为参数传入则可能被 GC 提前回收
        _winEventDelegate = WinEventCallback;
        Logger.Trace("WinEventDelegate 已创建并固定引用。", Cat);
    }

    // 安装所有需要监听的 WinEvent 钩子
    public void Start()
    {
        Logger.Info("安装 WinEvent 钩子…", Cat);

        // 移动/调整大小 开始与结束
        InstallHook(NativeConstants.EVENT_SYSTEM_MOVESIZESTART,
                    NativeConstants.EVENT_SYSTEM_MOVESIZEEND,
                    "MoveSize Start/End");

        // 窗口对象销毁
        InstallHook(NativeConstants.EVENT_OBJECT_DESTROY,
                    NativeConstants.EVENT_OBJECT_DESTROY,
                    "Object Destroy");

        // 最小化 开始与结束
        InstallHook(NativeConstants.EVENT_SYSTEM_MINIMIZESTART,
                    NativeConstants.EVENT_SYSTEM_MINIMIZEEND,
                    "Minimize Start/End");

        // 前景窗口变化
        InstallHook(NativeConstants.EVENT_SYSTEM_FOREGROUND,
                    NativeConstants.EVENT_SYSTEM_FOREGROUND,
                    "Foreground Change");

        // 位置/大小变化（高频，每帧可能触发多次）
        InstallHook(NativeConstants.EVENT_OBJECT_LOCATIONCHANGE,
                    NativeConstants.EVENT_OBJECT_LOCATIONCHANGE,
                    "Location Change");

        // 窗口标题变化
        InstallHook(NativeConstants.EVENT_OBJECT_NAMECHANGE,
                    NativeConstants.EVENT_OBJECT_NAMECHANGE,
                    "Name/Title Change");

        Logger.Info($"WinEvent 钩子安装完毕，共 {_hookHandles.Count} 个有效钩子。", Cat);
    }

    /// <summary>
    /// 安装单个 WinEvent 范围钩子。
    /// </summary>
    /// <param name="eventMin">监听的事件类型下限，对应 EVENT_* 常量。</param>
    /// <param name="eventMax">监听的事件类型上限，与 <paramref name="eventMin"/> 相同时只监听单一事件。</param>
    /// <param name="description">钩子描述，仅用于日志输出。</param>
    private void InstallHook(uint eventMin, uint eventMax, string description)
    {
        Logger.Trace($"InstallHook: [{description}] eventMin=0x{eventMin:X4} eventMax=0x{eventMax:X4}", Cat);

        var handle = NativeMethods.SetWinEventHook(
            eventMin, eventMax,
            IntPtr.Zero,          // 钩子 DLL，OUTOFCONTEXT 模式下传 Zero
            _winEventDelegate,
            0, 0,                 // 监听所有进程和线程
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

    /// <summary>
    /// 系统 WinEvent 回调，将原始事件分发为强类型 C# 事件。
    /// </summary>
    /// <param name="hWinEventHook">钩子句柄，通常不使用。</param>
    /// <param name="eventType">事件类型，对应 EVENT_* 常量。</param>
    /// <param name="hwnd">产生事件的窗口句柄。</param>
    /// <param name="idObject">事件对象标识，OBJID_WINDOW 表示窗口本身。</param>
    /// <param name="idChild">子对象标识，0 表示对象本身而非子元素。</param>
    /// <param name="dwEventThread">产生事件的线程 ID。</param>
    /// <param name="dwmsEventTime">事件发生时的系统时间戳（毫秒）。</param>
    private void WinEventCallback(
        IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        // 只处理窗口对象本身，忽略子控件和非窗口对象
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
                // LocationChange 在鼠标拖动期间每帧触发，只每 120 次输出一条控制台摘要以避免刷屏
                int lcCount = Interlocked.Increment(ref _locationChangeCount);
                if (lcCount % 120 == 0)
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

    // 卸载所有已安装的 WinEvent 钩子
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

    // 释放资源，调用 Stop 卸载钩子
    public void Dispose()
    {
        Logger.Debug("WinEventHookService.Dispose()", Cat);
        Stop();
        GC.SuppressFinalize(this);
    }
}