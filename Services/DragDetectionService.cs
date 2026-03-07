using System.Windows;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 检测窗口拖拽行为——同时处理外部窗口拖入和已托管窗口拖拽（互换 / 释放）
/// </summary>
public class DragDetectionService : IDisposable
{
    private readonly WinEventHookService _winEventHook;
    private IntPtr _draggingWindow = IntPtr.Zero;
    private System.Windows.Threading.DispatcherTimer? _trackTimer;

    // ── 管理区域 ──
    private Rect _dropZone = Rect.Empty;

    // ── 已托管窗口拖拽状态 ──
    private bool _isManagedDrag;
    private int  _managedDragOriginalSlotIndex = -1;

    // ── 事件（外部窗口拖入） ──

    /// <summary>
    /// 外部窗口开始被拖拽
    /// </summary>
    public event Action<IntPtr>? ExternalDragStarted;

    /// <summary>
    /// 拖拽中鼠标位置更新（物理像素）
    /// </summary>
    public event Action<IntPtr, Point>? DragMoved;

    /// <summary>
    /// 窗口被拖入管理区域并释放（鼠标落点在区域内）
    /// </summary>
    public event Action<IntPtr>? WindowDroppedInZone;

    /// <summary>
    /// 被管理的窗口拖出了管理区域
    /// </summary>
    public event Action<IntPtr>? WindowDraggedOutOfZone;

    /// <summary>
    /// 拖拽结束（无论是否落入区域）
    /// </summary>
    public event Action<IntPtr>? DragEnded;

    // ── 事件（已托管窗口拖拽） ──

    /// <summary>
    /// 已托管窗口开始被拖拽（通过原生移动或编程式发起）
    /// </summary>
    public event Action<IntPtr>? ManagedDragStarted;

    /// <summary>
    /// 已托管窗口拖拽中鼠标位置更新
    /// </summary>
    public event Action<IntPtr, Point>? ManagedDragMoved;

    /// <summary>
    /// 已托管窗口被拖入侧边栏（释放区域），应解除托管
    /// </summary>
    public event Action<IntPtr>? ManagedWindowDroppedOnSidebar;

    /// <summary>
    /// 已托管窗口被拖到另一个窗口的槽位上，应互换位置。
    /// 参数：(被拖拽的窗口句柄, 目标槽位索引)
    /// </summary>
    public event Action<IntPtr, int>? ManagedWindowSwapRequested;

    /// <summary>
    /// 已托管窗口拖拽结束但未触发任何操作（窗口应回到原位）
    /// </summary>
    public event Action<IntPtr>? ManagedDragCancelled;

    /// <summary>
    /// 已托管窗口拖拽结束（总是触发，用于清理覆盖层等）
    /// </summary>
    public event Action<IntPtr>? ManagedDragEnded;

    public IntPtr DraggingWindow => _draggingWindow;
    public bool IsDragging => _draggingWindow != IntPtr.Zero;
    public bool IsManagedDrag => _isManagedDrag;

    /// <summary>
    /// 当前被管理的窗口集合（由外部设置，用于判断拖出 / 判断是否为已托管窗口拖拽）
    /// </summary>
    public HashSet<IntPtr> ManagedWindows { get; } = new();

    public DragDetectionService(WinEventHookService winEventHook)
    {
        _winEventHook = winEventHook;
        _winEventHook.WindowMoveSizeStarted += OnMoveSizeStarted;
        _winEventHook.WindowMoveSizeEnded += OnMoveSizeEnded;
    }

    /// <summary>
    /// 设置放置区域的屏幕物理像素矩形。
    /// 与 GetCursorPos / PointToScreen 返回值为同一坐标系。
    /// </summary>
    public void SetDropZone(Rect screenPhysicalRect)
    {
        _dropZone = screenPhysicalRect;
    }

    private void OnMoveSizeStarted(IntPtr hwnd)
    {
        // ── 优先检测：是否为已托管窗口 ──
        if (ManagedWindows.Contains(hwnd))
        {
            _draggingWindow = hwnd;
            _isManagedDrag = true;
            _managedDragOriginalSlotIndex = -1; // 由外部在事件处理中设置

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
            });

            return;
        }

        // ── 外部窗口拖拽 ──
        if (!WindowEnumerator.ShouldInclude(hwnd))
            return;

        _draggingWindow = hwnd;
        _isManagedDrag = false;
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
        });
    }

    // ── 外部窗口拖拽跟踪 ──
    private void TrackTimer_Tick(object? sender, EventArgs e)
    {
        if (_draggingWindow == IntPtr.Zero) return;

        NativeMethods.GetCursorPos(out POINT pt);
        DragMoved?.Invoke(_draggingWindow, new Point(pt.X, pt.Y));
    }

    // ── 已托管窗口拖拽跟踪 ──
    private void ManagedTrackTimer_Tick(object? sender, EventArgs e)
    {
        if (_draggingWindow == IntPtr.Zero) return;

        NativeMethods.GetCursorPos(out POINT pt);
        ManagedDragMoved?.Invoke(_draggingWindow, new Point(pt.X, pt.Y));
    }

    private void OnMoveSizeEnded(IntPtr hwnd)
    {
        // 确保处理的是同一个窗口
        if (hwnd != _draggingWindow && _draggingWindow != IntPtr.Zero)
            hwnd = _draggingWindow;

        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            _trackTimer?.Stop();
            _trackTimer = null;
        });

        if (hwnd == IntPtr.Zero) return;

        if (_isManagedDrag)
        {
            HandleManagedDragEnd(hwnd);
        }
        else
        {
            HandleExternalDragEnd(hwnd);
        }

        _draggingWindow = IntPtr.Zero;
        _isManagedDrag = false;
        _managedDragOriginalSlotIndex = -1;
    }

    // ── 外部窗口拖拽结束处理 ──
    private void HandleExternalDragEnd(IntPtr hwnd)
    {
        bool wasManagedBefore = ManagedWindows.Contains(hwnd);

        NativeMethods.GetCursorPos(out POINT cursorPt);
        var cursorPoint = new Point(cursorPt.X, cursorPt.Y);

        bool isInZone = _dropZone != Rect.Empty && _dropZone.Contains(cursorPoint);

        if (!wasManagedBefore && isInZone)
        {
            WindowDroppedInZone?.Invoke(hwnd);
        }
        else if (wasManagedBefore && !isInZone)
        {
            WindowDraggedOutOfZone?.Invoke(hwnd);
        }

        DragEnded?.Invoke(hwnd);
    }

    // ── 已托管窗口拖拽结束处理 ──
    private void HandleManagedDragEnd(IntPtr hwnd)
    {
        NativeMethods.GetCursorPos(out POINT cursorPt);
        var cursorPoint = new Point(cursorPt.X, cursorPt.Y);

        bool isOverSidebar = _dropZone != Rect.Empty && _dropZone.Contains(cursorPoint);

        if (isOverSidebar)
        {
            // 拖拽到侧边栏 → 解除托管
            ManagedWindowDroppedOnSidebar?.Invoke(hwnd);
        }
        else
        {
            // 交由 MainWindow 根据鼠标位置判断是互换还是取消
            // MainWindow 会通过 LayoutService.GetSlotIndexAtScreenPoint 查找目标槽位
            // 这里只触发取消 / 互换事件由 MainWindow 在 ManagedDragEnded 中统一处理
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
        _draggingWindow = hwnd;
        _isManagedDrag = true;
        _managedDragOriginalSlotIndex = originalSlotIndex;

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
        });
    }

    public void Dispose()
    {
        Application.Current?.Dispatcher.BeginInvoke(() => _trackTimer?.Stop());
        _winEventHook.WindowMoveSizeStarted -= OnMoveSizeStarted;
        _winEventHook.WindowMoveSizeEnded -= OnMoveSizeEnded;
        GC.SuppressFinalize(this);
    }
}