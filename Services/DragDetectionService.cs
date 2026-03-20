using System.Windows;
using WindowPilot.Diagnostics;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 检测窗口拖拽行为，同时处理外部窗口拖入和已托管窗口的拖拽（互换位置或拖至侧边栏释放）。
/// 订阅 <see cref="WinEventHookService"/> 的移动事件，通过 DispatcherTimer 以 16ms 间隔轮询光标位置。
/// </summary>
public class DragDetectionService : IDisposable
{
    private const string Cat = "DragDetection";

    private readonly WinEventHookService _winEventHook;
    private IntPtr _draggingWindow = IntPtr.Zero; // 当前正在被拖拽的窗口句柄
    private System.Windows.Threading.DispatcherTimer? _trackTimer; // 拖拽期间的光标位置轮询计时器

    // 管理区域（侧边栏的屏幕物理像素矩形）
    private Rect _dropZone = Rect.Empty;

    // 已托管窗口拖拽状态
    private bool _isManagedDrag;
    private int  _managedDragOriginalSlotIndex = -1; // 拖拽开始时的原始槽位，-1 表示未记录

    // 拖拽跟踪统计
    private int _moveTickCount;    // 当前会话的 Tick 累计数
    private int _dragSessionCount; // 历史拖拽会话总数
    /// <summary>
    /// 记录外部窗口在拖拽开始瞬间（尚未移动时）的屏幕位置。
    /// 键为窗口句柄，值为 GetWindowRect 的结果。
    /// 托管成功后由 MainWindow 读取并写入 ManagedWindow.OriginalRect，之后删除条目。
    /// </summary>
    public Dictionary<IntPtr, RECT> PreDragRects { get; } = new();
    
    
    // 外部窗口拖入事件
    public event Action<IntPtr>?        ExternalDragStarted;    // 外部窗口开始被拖拽
    public event Action<IntPtr, Point>? DragMoved;              // 拖拽中光标位置更新（物理像素）
    public event Action<IntPtr>?        WindowDroppedInZone;    // 外部窗口在区域内释放，应纳入托管
    public event Action<IntPtr>?        WindowDraggedOutOfZone; // 已托管窗口被拖出管理区域，应释放
    public event Action<IntPtr>?        DragEnded;              // 拖拽结束，无论是否落入区域

    // 已托管窗口拖拽事件
    public event Action<IntPtr>?        ManagedDragStarted;           // 托管窗口开始被拖拽
    public event Action<IntPtr, Point>? ManagedDragMoved;             // 托管窗口拖拽中光标位置更新
    public event Action<IntPtr>?        ManagedWindowDroppedOnSidebar; // 托管窗口拖至侧边栏，应解除托管
    /// <summary>
    /// 托管窗口被拖到另一个槽位上，触发位置互换。
    /// 参数：(被拖拽的窗口句柄, 目标槽位索引)。
    /// </summary>
    public event Action<IntPtr, int>?   ManagedWindowSwapRequested;
    public event Action<IntPtr>?        ManagedDragCancelled; // 托管窗口拖拽取消，应回到原位
    public event Action<IntPtr>?        ManagedDragEnded;     // 托管窗口拖拽结束（总是触发，用于清理覆盖层）

    public IntPtr DraggingWindow => _draggingWindow;
    public bool   IsDragging     => _draggingWindow != IntPtr.Zero;
    public bool   IsManagedDrag  => _isManagedDrag;

    // 由外部维护的当前托管窗口句柄集合，用于判断拖拽是否来自已托管窗口
    public HashSet<IntPtr> ManagedWindows { get; } = new();

    /// <summary>
    /// 构造拖拽检测服务，订阅 WinEvent 移动事件。
    /// </summary>
    /// <param name="winEventHook">WinEvent 钩子服务实例，提供 MoveSizeStarted/Ended 事件。</param>
    public DragDetectionService(WinEventHookService winEventHook)
    {
        Logger.Debug("DragDetectionService 构造中…", Cat);
        _winEventHook = winEventHook;
        _winEventHook.WindowMoveSizeStarted += OnMoveSizeStarted;
        _winEventHook.WindowMoveSizeEnded   += OnMoveSizeEnded;
        Logger.Debug("已订阅 WindowMoveSizeStarted / WindowMoveSizeEnded 事件。", Cat);
    }

    /// <summary>
    /// 设置放置区域的屏幕物理像素矩形。坐标系与 GetCursorPos 一致。
    /// </summary>
    /// <param name="screenPhysicalRect">侧边栏在屏幕上的物理像素矩形，用于拖拽落点判断。</param>
    public void SetDropZone(Rect screenPhysicalRect)
    {
        Logger.Debug(
            $"SetDropZone: Left={screenPhysicalRect.Left:F0} Top={screenPhysicalRect.Top:F0} " +
            $"W={screenPhysicalRect.Width:F0} H={screenPhysicalRect.Height:F0}", Cat);
        _dropZone = screenPhysicalRect;
    }

    /// <summary>
    /// 响应 EVENT_SYSTEM_MOVESIZESTART，区分托管窗口拖拽与外部窗口拖拽并分别处理。
    /// </summary>
    /// <param name="hwnd">开始移动或调整大小的窗口句柄。</param>
    private void OnMoveSizeStarted(IntPtr hwnd)
    {
        // 优先检测：托管窗口拖拽
        if (ManagedWindows.Contains(hwnd))
        {
            int session = Interlocked.Increment(ref _dragSessionCount);
            Logger.Info($"[Session #{session}] 托管窗口拖拽开始  hwnd=0x{hwnd:X}", Cat);

            _draggingWindow = hwnd;
            _isManagedDrag  = true;
            _managedDragOriginalSlotIndex = -1;
            _moveTickCount = 0;

            ManagedDragStarted?.Invoke(hwnd);

            // 在 UI 线程上启动计时器，DispatcherTimer 必须在 UI 线程创建
            Application.Current?.Dispatcher.BeginInvoke(() =>
            {
                _trackTimer?.Stop();
                _trackTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(16) // 约 60fps 的轮询频率
                };
                _trackTimer.Tick += ManagedTrackTimer_Tick;
                _trackTimer.Start();
                Logger.Trace("托管拖拽跟踪计时器已启动（16ms）。", Cat);
            });

            return;
        }

        // 外部窗口拖拽，过滤不符合托管条件的窗口
        if (!WindowEnumerator.ShouldInclude(hwnd))
        {
            Logger.Trace($"OnMoveSizeStarted: hwnd=0x{hwnd:X} 不在候选范围，忽略。", Cat);
            return;
        }

        int extSession = Interlocked.Increment(ref _dragSessionCount);
        Logger.Info($"[Session #{extSession}] 外部窗口拖拽开始  hwnd=0x{hwnd:X}", Cat);

        _draggingWindow = hwnd;
        _isManagedDrag  = false;
        _moveTickCount  = 0;

        // 在拖拽刚开始（窗口尚未被移动）时立即快照位置，
        // 供托管成功后写入 OriginalRect，避免 Windows 松手时先移动窗口导致位置错误
        if (NativeMethods.GetWindowRect(hwnd, out RECT preDragRect))
        {
            PreDragRects[hwnd] = preDragRect;
            Logger.Debug(
                $"  PreDragRect 已记录: ({preDragRect.Left},{preDragRect.Top}," +
                $"{preDragRect.Width}×{preDragRect.Height})", Cat);
        }

        ExternalDragStarted?.Invoke(hwnd);

        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            _trackTimer?.Stop();
            _trackTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(16)
            };
            _trackTimer.Tick += TrackTimer_Tick;
            _trackTimer.Start();
            Logger.Trace("外部拖拽跟踪计时器已启动（16ms）。", Cat);
        });
    }

    // 外部窗口拖拽的光标位置轮询
    private void TrackTimer_Tick(object? sender, EventArgs e)
    {
        if (_draggingWindow == IntPtr.Zero) return;

        NativeMethods.GetCursorPos(out POINT pt);
        int tick = Interlocked.Increment(ref _moveTickCount);

        // 每 60 帧（约 1 秒）输出一次摘要，避免控制台被高频日志刷屏
        if (tick % 60 == 0)
            Logger.Trace($"拖拽跟踪 Tick#{tick}  hwnd=0x{_draggingWindow:X}  cursor=({pt.X},{pt.Y})", Cat);

        DragMoved?.Invoke(_draggingWindow, new Point(pt.X, pt.Y));
    }

    // 托管窗口拖拽的光标位置轮询
    private void ManagedTrackTimer_Tick(object? sender, EventArgs e)
    {
        if (_draggingWindow == IntPtr.Zero) return;

        NativeMethods.GetCursorPos(out POINT pt);
        int tick = Interlocked.Increment(ref _moveTickCount);

        if (tick % 60 == 0)
            Logger.Trace($"托管拖拽 Tick#{tick}  hwnd=0x{_draggingWindow:X}  cursor=({pt.X},{pt.Y})", Cat);

        ManagedDragMoved?.Invoke(_draggingWindow, new Point(pt.X, pt.Y));
    }

    /// <summary>
    /// 响应 EVENT_SYSTEM_MOVESIZEEND，停止计时器并根据拖拽类型分发结束处理。
    /// </summary>
    /// <param name="hwnd">完成移动或调整大小的窗口句柄。</param>
    private void OnMoveSizeEnded(IntPtr hwnd)
    {
        Logger.Debug($"OnMoveSizeEnded  hwnd=0x{hwnd:X}  dragging=0x{_draggingWindow:X}", Cat);

        // 防御性处理：若结束的窗口与记录的不一致，强制使用记录值
        if (hwnd != _draggingWindow && _draggingWindow != IntPtr.Zero)
        {
            Logger.Warning(
                $"MoveSizeEnded hwnd(0x{hwnd:X}) != draggingWindow(0x{_draggingWindow:X})，强制使用 draggingWindow。", Cat);
            hwnd = _draggingWindow;
        }

        // 在 UI 线程上停止计时器
        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            _trackTimer?.Stop();
            _trackTimer = null;
            Logger.Trace("拖拽跟踪计时器已停止。", Cat);
        });

        if (hwnd == IntPtr.Zero)
        {
            Logger.Warning("OnMoveSizeEnded: hwnd == IntPtr.Zero，跳过处理。", Cat);
            return;
        }

        Logger.Debug(
            $"拖拽结束处理  isManagedDrag={_isManagedDrag}  totalTicks={_moveTickCount}", Cat);

        if (_isManagedDrag)
            HandleManagedDragEnd(hwnd);
        else
            HandleExternalDragEnd(hwnd);

        // 重置拖拽状态
        _draggingWindow = IntPtr.Zero;
        _isManagedDrag  = false;
        _managedDragOriginalSlotIndex = -1;
    }

    /// <summary>
    /// 处理外部窗口拖拽结束：判断是否落入区域，触发对应事件。
    /// </summary>
    /// <param name="hwnd">完成拖拽的外部窗口句柄。</param>
    private void HandleExternalDragEnd(IntPtr hwnd)
    {
        bool wasManagedBefore = ManagedWindows.Contains(hwnd);

        NativeMethods.GetCursorPos(out POINT cursorPt);
        var cursorPoint = new Point(cursorPt.X, cursorPt.Y);

        // 判断鼠标是否在侧边栏区域内松开
        bool isInZone = _dropZone != Rect.Empty && _dropZone.Contains(cursorPoint);

        Logger.Debug(
            $"HandleExternalDragEnd  hwnd=0x{hwnd:X}  " +
            $"wasManaged={wasManagedBefore}  cursor=({cursorPt.X},{cursorPt.Y})  " +
            $"isInZone={isInZone}  zone={_dropZone}", Cat);

        if (!wasManagedBefore && isInZone)
        {
            // 新窗口拖入区域，应纳入托管
            Logger.Info($"外部窗口 0x{hwnd:X} 已拖入管理区域 → 触发 WindowDroppedInZone。", Cat);
            WindowDroppedInZone?.Invoke(hwnd);
        }
        else if (wasManagedBefore && !isInZone)
        {
            // 已托管窗口被拖到区域外，应解除托管
            Logger.Info($"已托管窗口 0x{hwnd:X} 被拖出管理区域 → 触发 WindowDraggedOutOfZone。", Cat);
            WindowDraggedOutOfZone?.Invoke(hwnd);
        }
        else
        {
            Logger.Debug($"外部窗口 0x{hwnd:X} 拖拽结束，无区域匹配变化。", Cat);
        }

        // 无论是否托管，拖拽结束后都清理预存的位置快照
        PreDragRects.Remove(hwnd);

        DragEnded?.Invoke(hwnd);
    }

    /// <summary>
    /// 处理已托管窗口拖拽结束：判断是否落在侧边栏上，触发解除托管或取消事件。
    /// </summary>
    /// <param name="hwnd">完成拖拽的托管窗口句柄。</param>
    private void HandleManagedDragEnd(IntPtr hwnd)
    {
        NativeMethods.GetCursorPos(out POINT cursorPt);
        var cursorPoint = new Point(cursorPt.X, cursorPt.Y);

        bool isOverSidebar = _dropZone != Rect.Empty && _dropZone.Contains(cursorPoint);

        Logger.Debug(
            $"HandleManagedDragEnd  hwnd=0x{hwnd:X}  " +
            $"cursor=({cursorPt.X},{cursorPt.Y})  isOverSidebar={isOverSidebar}", Cat);

        if (isOverSidebar)
        {
            // 落在侧边栏，解除托管
            Logger.Info($"托管窗口 0x{hwnd:X} 拖至侧边栏 → 触发 ManagedWindowDroppedOnSidebar。", Cat);
            ManagedWindowDroppedOnSidebar?.Invoke(hwnd);
        }
        else
        {
            // 未落在侧边栏，取消拖拽，窗口将回到原位
            Logger.Debug($"托管窗口 0x{hwnd:X} 拖拽结束，未至侧边栏 → 触发 ManagedDragCancelled。", Cat);
            ManagedDragCancelled?.Invoke(hwnd);
        }

        // ManagedDragEnded 总是触发，供调用方统一清理覆盖层等 UI
        ManagedDragEnded?.Invoke(hwnd);
    }

    /// <summary>
    /// 编程式通知拖拽已开始，用于热键发起的拖拽。
    /// 调用方应在 TemporaryUnembed 并 PostMessage SC_MOVE 之后立即调用此方法。
    /// </summary>
    /// <param name="hwnd">被拖拽的托管窗口句柄。</param>
    /// <param name="originalSlotIndex">拖拽开始前该窗口所在的槽位索引，取消时用于回位。</param>
    public void NotifyProgrammaticDragStarted(IntPtr hwnd, int originalSlotIndex)
    {
        int session = Interlocked.Increment(ref _dragSessionCount);
        Logger.Info(
            $"[Session #{session}] 编程式拖拽通知  hwnd=0x{hwnd:X}  originalSlot={originalSlotIndex}", Cat);

        _draggingWindow = hwnd;
        _isManagedDrag  = true;
        _managedDragOriginalSlotIndex = originalSlotIndex;
        _moveTickCount  = 0;

        ManagedDragStarted?.Invoke(hwnd);

        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            _trackTimer?.Stop();
            _trackTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(16)
            };
            _trackTimer.Tick += ManagedTrackTimer_Tick;
            _trackTimer.Start();
            Logger.Trace("编程式拖拽跟踪计时器已启动（16ms）。", Cat);
        });
    }

    // 取消订阅事件并停止计时器
    public void Dispose()
    {
        Logger.Debug("DragDetectionService.Dispose()", Cat);
        Application.Current?.Dispatcher.BeginInvoke(() => _trackTimer?.Stop());
        _winEventHook.WindowMoveSizeStarted -= OnMoveSizeStarted;
        _winEventHook.WindowMoveSizeEnded   -= OnMoveSizeEnded;
        Logger.Debug(
            $"统计 — 拖拽会话: {_dragSessionCount}次  总跟踪帧数: {_moveTickCount}", Cat);
        GC.SuppressFinalize(this);
    }
}