using System.Windows;
using WindowPilot.Diagnostics;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 检测窗口拖拽行为——同时处理外部窗口拖入和已托管窗口拖拽（互换 / 释放）
/// </summary>
public class DragDetectionService : IDisposable
{
    private const string Cat = "DragDetection";

    private readonly WinEventHookService _winEventHook;
    private IntPtr _draggingWindow = IntPtr.Zero;
    private System.Windows.Threading.DispatcherTimer? _trackTimer;

    // ── 管理区域 ──
    private Rect _dropZone = Rect.Empty;

    // ── 已托管窗口拖拽状态 ──
    private bool _isManagedDrag;
    private int  _managedDragOriginalSlotIndex = -1;

    // ── 拖拽跟踪统计 ──
    private int  _moveTickCount;
    private int  _dragSessionCount;

    // ── 事件（外部窗口拖入） ──

    /// <summary>外部窗口开始被拖拽</summary>
    public event Action<IntPtr>? ExternalDragStarted;

    /// <summary>拖拽中鼠标位置更新（物理像素）</summary>
    public event Action<IntPtr, Point>? DragMoved;

    /// <summary>窗口被拖入管理区域并释放（鼠标落点在区域内）</summary>
    public event Action<IntPtr>? WindowDroppedInZone;

    /// <summary>被管理的窗口拖出了管理区域</summary>
    public event Action<IntPtr>? WindowDraggedOutOfZone;

    /// <summary>拖拽结束（无论是否落入区域）</summary>
    public event Action<IntPtr>? DragEnded;

    // ── 事件（已托管窗口拖拽） ──

    /// <summary>已托管窗口开始被拖拽（通过原生移动或编程式发起）</summary>
    public event Action<IntPtr>? ManagedDragStarted;

    /// <summary>已托管窗口拖拽中鼠标位置更新</summary>
    public event Action<IntPtr, Point>? ManagedDragMoved;

    /// <summary>已托管窗口被拖入侧边栏（释放区域），应解除托管</summary>
    public event Action<IntPtr>? ManagedWindowDroppedOnSidebar;

    /// <summary>
    /// 已托管窗口被拖到另一个窗口的槽位上，应互换位置。
    /// 参数：(被拖拽的窗口句柄, 目标槽位索引)
    /// </summary>
    public event Action<IntPtr, int>? ManagedWindowSwapRequested;

    /// <summary>已托管窗口拖拽结束但未触发任何操作（窗口应回到原位）</summary>
    public event Action<IntPtr>? ManagedDragCancelled;

    /// <summary>已托管窗口拖拽结束（总是触发，用于清理覆盖层等）</summary>
    public event Action<IntPtr>? ManagedDragEnded;

    public IntPtr DraggingWindow => _draggingWindow;
    public bool IsDragging       => _draggingWindow != IntPtr.Zero;
    public bool IsManagedDrag    => _isManagedDrag;

    /// <summary>
    /// 当前被管理的窗口集合（由外部设置，用于判断拖出 / 判断是否为已托管窗口拖拽）
    /// </summary>
    public HashSet<IntPtr> ManagedWindows { get; } = new();

    public DragDetectionService(WinEventHookService winEventHook)
    {
        Logger.Debug("DragDetectionService 构造中…", Cat);
        _winEventHook = winEventHook;
        _winEventHook.WindowMoveSizeStarted += OnMoveSizeStarted;
        _winEventHook.WindowMoveSizeEnded   += OnMoveSizeEnded;
        Logger.Debug("已订阅 WindowMoveSizeStarted / WindowMoveSizeEnded 事件。", Cat);
    }

    /// <summary>
    /// 设置放置区域的屏幕物理像素矩形。
    /// 与 GetCursorPos / PointToScreen 返回值为同一坐标系。
    /// </summary>
    public void SetDropZone(Rect screenPhysicalRect)
    {
        Logger.Debug(
            $"SetDropZone: Left={screenPhysicalRect.Left:F0} Top={screenPhysicalRect.Top:F0} " +
            $"W={screenPhysicalRect.Width:F0} H={screenPhysicalRect.Height:F0}", Cat);
        _dropZone = screenPhysicalRect;
    }

    private void OnMoveSizeStarted(IntPtr hwnd)
    {
        // ── 优先检测：是否为已托管窗口 ──
        if (ManagedWindows.Contains(hwnd))
        {
            int session = Interlocked.Increment(ref _dragSessionCount);
            Logger.Info($"[Session #{session}] 托管窗口拖拽开始  hwnd=0x{hwnd:X}", Cat);

            _draggingWindow = hwnd;
            _isManagedDrag  = true;
            _managedDragOriginalSlotIndex = -1;
            _moveTickCount = 0;

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
                Logger.Trace("托管拖拽跟踪计时器已启动（16ms）。", Cat);
            });

            return;
        }

        // ── 外部窗口拖拽 ──
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

    // ── 外部窗口拖拽跟踪 ──
    private void TrackTimer_Tick(object? sender, EventArgs e)
    {
        if (_draggingWindow == IntPtr.Zero) return;

        NativeMethods.GetCursorPos(out POINT pt);
        int tick = Interlocked.Increment(ref _moveTickCount);

        // 每 60 帧（约 1 秒）输出一次位置摘要，避免控制台刷屏
        if (tick % 60 == 0)
            Logger.Trace($"拖拽跟踪 Tick#{tick}  hwnd=0x{_draggingWindow:X}  cursor=({pt.X},{pt.Y})", Cat);

        DragMoved?.Invoke(_draggingWindow, new Point(pt.X, pt.Y));
    }

    // ── 已托管窗口拖拽跟踪 ──
    private void ManagedTrackTimer_Tick(object? sender, EventArgs e)
    {
        if (_draggingWindow == IntPtr.Zero) return;

        NativeMethods.GetCursorPos(out POINT pt);
        int tick = Interlocked.Increment(ref _moveTickCount);

        if (tick % 60 == 0)
            Logger.Trace($"托管拖拽 Tick#{tick}  hwnd=0x{_draggingWindow:X}  cursor=({pt.X},{pt.Y})", Cat);

        ManagedDragMoved?.Invoke(_draggingWindow, new Point(pt.X, pt.Y));
    }

    private void OnMoveSizeEnded(IntPtr hwnd)
    {
        Logger.Debug($"OnMoveSizeEnded  hwnd=0x{hwnd:X}  dragging=0x{_draggingWindow:X}", Cat);

        // 确保处理的是同一个窗口
        if (hwnd != _draggingWindow && _draggingWindow != IntPtr.Zero)
        {
            Logger.Warning(
                $"MoveSizeEnded hwnd(0x{hwnd:X}) != draggingWindow(0x{_draggingWindow:X})，强制使用 draggingWindow。", Cat);
            hwnd = _draggingWindow;
        }

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

        _draggingWindow = IntPtr.Zero;
        _isManagedDrag  = false;
        _managedDragOriginalSlotIndex = -1;
    }

    // ── 外部窗口拖拽结束处理 ──
    private void HandleExternalDragEnd(IntPtr hwnd)
    {
        bool wasManagedBefore = ManagedWindows.Contains(hwnd);

        NativeMethods.GetCursorPos(out POINT cursorPt);
        var cursorPoint = new Point(cursorPt.X, cursorPt.Y);

        bool isInZone = _dropZone != Rect.Empty && _dropZone.Contains(cursorPoint);

        Logger.Debug(
            $"HandleExternalDragEnd  hwnd=0x{hwnd:X}  " +
            $"wasManaged={wasManagedBefore}  cursor=({cursorPt.X},{cursorPt.Y})  " +
            $"isInZone={isInZone}  zone={_dropZone}", Cat);

        if (!wasManagedBefore && isInZone)
        {
            Logger.Info($"外部窗口 0x{hwnd:X} 已拖入管理区域 → 触发 WindowDroppedInZone。", Cat);
            WindowDroppedInZone?.Invoke(hwnd);
        }
        else if (wasManagedBefore && !isInZone)
        {
            Logger.Info($"已托管窗口 0x{hwnd:X} 被拖出管理区域 → 触发 WindowDraggedOutOfZone。", Cat);
            WindowDraggedOutOfZone?.Invoke(hwnd);
        }
        else
        {
            Logger.Debug($"外部窗口 0x{hwnd:X} 拖拽结束，无区域匹配变化。", Cat);
        }

        DragEnded?.Invoke(hwnd);
    }

    // ── 已托管窗口拖拽结束处理 ──
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
            Logger.Info($"托管窗口 0x{hwnd:X} 拖至侧边栏 → 触发 ManagedWindowDroppedOnSidebar。", Cat);
            ManagedWindowDroppedOnSidebar?.Invoke(hwnd);
        }
        else
        {
            Logger.Debug($"托管窗口 0x{hwnd:X} 拖拽结束，未至侧边栏 → 触发 ManagedDragCancelled。", Cat);
            ManagedDragCancelled?.Invoke(hwnd);
        }

        ManagedDragEnded?.Invoke(hwnd);
    }

    /// <summary>
    /// 由外部调用，编程式通知拖拽已开始（用于热键发起的拖拽）。
    /// 此时窗口已被 TemporaryUnembed 并进入 SC_MOVE 循环。
    /// </summary>
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