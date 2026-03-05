using System.Windows;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 检测窗口拖拽行为
/// </summary>
public class DragDetectionService : IDisposable
{
    private readonly WinEventHookService _winEventHook;
    private IntPtr _draggingWindow = IntPtr.Zero;
    private System.Windows.Threading.DispatcherTimer? _trackTimer;

    // ── 管理区域 ──
    private Rect _dropZone = Rect.Empty;

    // ── 事件 ──

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

    public IntPtr DraggingWindow => _draggingWindow;
    public bool IsDragging => _draggingWindow != IntPtr.Zero;

    /// <summary>
    /// 当前被管理的窗口集合（由外部设置，用于判断拖出）
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
        if (!WindowEnumerator.ShouldInclude(hwnd))
            return;

        _draggingWindow = hwnd;
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

    private void TrackTimer_Tick(object? sender, EventArgs e)
    {
        if (_draggingWindow == IntPtr.Zero) return;

        NativeMethods.GetCursorPos(out POINT pt);
        DragMoved?.Invoke(_draggingWindow, new Point(pt.X, pt.Y));
    }

    private void OnMoveSizeEnded(IntPtr hwnd)
    {
        if (hwnd != _draggingWindow && _draggingWindow != IntPtr.Zero)
            hwnd = _draggingWindow;

        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            _trackTimer?.Stop();
            _trackTimer = null;
        });

        if (hwnd == IntPtr.Zero) return;

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
        _draggingWindow = IntPtr.Zero;
    }

    public void Dispose()
    {
        Application.Current?.Dispatcher.BeginInvoke(() => _trackTimer?.Stop());
        _winEventHook.WindowMoveSizeStarted -= OnMoveSizeStarted;
        _winEventHook.WindowMoveSizeEnded -= OnMoveSizeEnded;
        GC.SuppressFinalize(this);
    }
}