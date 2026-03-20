using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using WindowPilot.Controls;
using WindowPilot.Diagnostics;
using WindowPilot.Models;
using WindowPilot.Native;
using WindowPilot.Services;

namespace WindowPilot;

public partial class MainWindow : Window
{
    private const string Cat = "MainWindow";

    // 服务层
    private readonly WinEventHookService  _winEventHook  = new();
    private readonly WindowManagerService _windowManager;
    private readonly DragDetectionService _dragDetection;
    private readonly LayoutService        _layout;
    private readonly HotkeyService        _hotkey        = new();

    // 外部窗口拖入时显示在侧边栏上方的蓝色提示覆盖层
    private DropZoneOverlay?    _overlay;

    // 托管窗口拖拽时显示在侧边栏上方的红色解除提示覆盖层
    private ReleaseZoneOverlay? _releaseOverlay;

    // 缩略图预览服务（鼠标悬停侧边栏条目时显示实时预览浮窗）
    private ThumbnailPreviewService? _thumbnailPreview;

    // 悬停防抖计时器：鼠标停留超过阈值后才真正显示预览，避免快速滑过时频繁弹出
    private System.Windows.Threading.DispatcherTimer? _hoverTimer;

    // 当前正在悬停的目标窗口及其屏幕坐标（由 SidebarItem_MouseEnter 更新）
    private ManagedWindow? _hoverTarget;
    private Point          _hoverItemScreenTL;   // 悬停条目左上角屏幕坐标
    private double         _sidebarScreenRight;  // 侧边栏右边界屏幕 X 坐标

    // 侧边栏折叠状态
    private bool   _sidebarExpanded   = true;
    private double _savedSidebarWidth = 200; // 折叠前的侧边栏宽度，用于展开时恢复

    // 侧边栏条目的 WPF DragDrop 状态
    private Point _dragItemStartPoint;    // 鼠标按下时的坐标，用于判断是否超过拖拽阈值
    private bool  _isItemDragInProgress;  // 是否有条目拖拽正在进行

    // 构造函数

    public MainWindow()
    {
        Logger.Separator("MainWindow Init");
        Logger.Debug("MainWindow 构造中…", Cat);

        InitializeComponent();

        // 服务依赖关系：WinEventHook → WindowManager → Layout/DragDetection
        _windowManager = new WindowManagerService(_winEventHook);
        _dragDetection = new DragDetectionService(_winEventHook);
        _layout        = new LayoutService(_windowManager);

        // 将托管列表绑定到侧边栏 ListBox
        WindowListBox.ItemsSource = _windowManager.ManagedWindows;
        _windowManager.ManagedWindows.CollectionChanged += OnManagedWindowsChanged;

        // 窗口管理器事件订阅
        _windowManager.WindowManaged  += w => Dispatcher.BeginInvoke(() =>
        {
            Logger.Info($"WindowManaged 回调: \"{w.Title}\"", Cat);
            SetStatus($"已接管：{w.Title}");
        });
        _windowManager.WindowReleased += w => Dispatcher.BeginInvoke(() =>
        {
            Logger.Info($"WindowReleased 回调: \"{w.Title}\"", Cat);
            SetStatus($"已释放：{w.Title}");
        });
        _windowManager.LayoutChanged  += () => Dispatcher.BeginInvoke(() =>
        {
            Logger.Trace("LayoutChanged 触发 → ApplyLayout", Cat);
            _layout.ApplyLayout();
        });
        _windowManager.ManageFailed   += (_, msg) => Dispatcher.BeginInvoke(() =>
        {
            Logger.Warning($"ManageFailed: {msg}", Cat);
            SetStatus(msg);
        });

        // 外部窗口拖入侧边栏事件订阅
        _dragDetection.ExternalDragStarted    += _ => Dispatcher.BeginInvoke(() =>
        {
            Logger.Debug("ExternalDragStarted → ShowDropOverlay", Cat);
            ShowDropOverlay();
        });
        _dragDetection.DragEnded              += _ => Dispatcher.BeginInvoke(() =>
        {
            Logger.Debug("DragEnded → HideDropOverlay", Cat);
            HideDropOverlay();
        });
        _dragDetection.DragMoved              += (_, _) => { /* 鼠标跟踪仅写日志，无 UI 更新 */ };
        _dragDetection.WindowDroppedInZone    += OnWindowDroppedInZone;
        _dragDetection.WindowDraggedOutOfZone += OnWindowDraggedOutOfZone;

        // 托管窗口拖拽事件订阅（互换位置 / 拖至侧边栏释放）
        _dragDetection.ManagedDragStarted            += OnManagedDragStarted;
        _dragDetection.ManagedDragMoved              += OnManagedDragMoved;
        _dragDetection.ManagedWindowDroppedOnSidebar += OnManagedWindowDroppedOnSidebar;
        _dragDetection.ManagedDragCancelled          += OnManagedDragCancelled;
        _dragDetection.ManagedDragEnded              += OnManagedDragEnded;

        Logger.Debug("所有事件订阅完毕。", Cat);
    }

    // 生命周期

    /// <summary>
    /// 窗口加载完成后初始化宿主句柄、WinEvent 钩子、全局热键和覆盖层。
    /// </summary>
    /// <param name="sender">事件发送者。</param>
    /// <param name="e">路由事件参数。</param>
    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        Logger.Separator("MainWindow Loaded");

        // 获取宿主 HWND，窗口 Loaded 后句柄才可用
        _windowManager.HostHwnd = new WindowInteropHelper(this).Handle;
        Logger.Info($"HostHwnd = 0x{_windowManager.HostHwnd:X}", Cat);

        _winEventHook.Start();

        // 注册全局热键，必须在窗口 Loaded 后调用 Attach
        _hotkey.Attach(this);

        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x47,
            GrabForegroundWindow, "Ctrl+Alt+G 抓取前台窗口");

        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x25,
            () => Dispatcher.BeginInvoke(_layout.SwitchToPrevious), "Ctrl+Alt+← 切换上一个");

        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x27,
            () => Dispatcher.BeginInvoke(_layout.SwitchToNext), "Ctrl+Alt+→ 切换下一个");

        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x51,
            () => Dispatcher.BeginInvoke(() => SetLayoutMode(LayoutService.LayoutMode.QuadSplit)),
            "Ctrl+Alt+Q 四分区布局");

        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x4D,
            StartActiveWindowDrag, "Ctrl+Alt+M 编程式拖拽");

        // 预先创建覆盖层窗口，拖拽时直接显示/隐藏避免创建延迟
        _overlay        = new DropZoneOverlay();
        _releaseOverlay = new ReleaseZoneOverlay();
        Logger.Debug("覆盖层窗口已创建: DropZone / ReleaseZone", Cat);

        // 初始化缩略图预览服务
        _thumbnailPreview = new ThumbnailPreviewService();

        // 悬停防抖计时器：350ms 后触发预览，期间移走则取消
        _hoverTimer          = new System.Windows.Threading.DispatcherTimer();
        _hoverTimer.Interval = TimeSpan.FromMilliseconds(350);
        _hoverTimer.Tick    += HoverTimer_Tick;
        Logger.Debug("缩略图预览服务及悬停计时器已初始化。", Cat);

        // 延迟一帧后计算初始布局，确保控件尺寸已确定
        Dispatcher.BeginInvoke(() => { UpdateHostArea(); UpdateDropZone(); });

        Logger.Info("MainWindow 已加载完毕，程序就绪。", Cat);
        Logger.Debug($"日志文件位置: {Logger.CurrentLogFilePath}", Cat);

        var hwndSource = HwndSource.FromHwnd(_windowManager.HostHwnd);
        hwndSource?.AddHook(HostWndProc);
        Logger.Info("HostWndProc 钩子已安装。", Cat);
        
        SetStatus("就绪 — 拖拽窗口到侧边栏以接管，或按 Ctrl+Alt+G 抓取当前窗口 │ Ctrl+Alt+M 拖拽移动窗口");
    }

    /// <summary>
    /// 主窗口关闭时按依赖顺序逐一释放所有资源。
    /// </summary>
    /// <param name="sender">事件发送者。</param>
    /// <param name="e">取消关闭事件参数，可设置 Cancel 阻止关闭。</param>
    private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        Logger.Separator("MainWindow Closing");
        Logger.Info("主窗口正在关闭，开始清理…", Cat);

        _hotkey.Dispose();
        Logger.Debug("HotkeyService 已释放。", Cat);

        // 停止悬停计时器并释放缩略图预览
        _hoverTimer?.Stop();
        _thumbnailPreview?.Dispose();
        Logger.Debug("ThumbnailPreviewService 已释放。", Cat);

        // ReleaseAll 会还原所有托管窗口，必须在 WinEventHook 停止前执行
        _windowManager.ReleaseAll();
        _windowManager.Dispose();
        Logger.Debug("WindowManagerService 已释放。", Cat);

        _dragDetection.Dispose();
        Logger.Debug("DragDetectionService 已释放。", Cat);

        _winEventHook.Dispose();
        Logger.Debug("WinEventHookService 已释放。", Cat);

        _overlay?.Close();
        _releaseOverlay?.Close();
        Logger.Debug("覆盖层窗口已关闭。", Cat);

        Logger.Info("MainWindow 清理完成。", Cat);
    }

    /// <summary>
    /// 主窗口大小变化时重新计算宿主区域和放置区域坐标。
    /// </summary>
    /// <param name="sender">事件发送者。</param>
    /// <param name="e">包含新旧尺寸信息的参数。</param>
    private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        Logger.Trace(
            $"MainWindow SizeChanged: {e.PreviousSize.Width:F0}×{e.PreviousSize.Height:F0} → " +
            $"{e.NewSize.Width:F0}×{e.NewSize.Height:F0}", Cat);
        Dispatcher.BeginInvoke(() => { UpdateHostArea(); UpdateDropZone(); });
    }

    /// <summary>
    /// HostPanel 尺寸变化时重新计算宿主区域，分隔条拖动时会频繁触发。
    /// </summary>
    /// <param name="sender">事件发送者。</param>
    /// <param name="e">包含新尺寸信息的参数。</param>
    private void HostPanel_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        Logger.Trace($"HostPanel SizeChanged: {e.NewSize.Width:F0}×{e.NewSize.Height:F0}", Cat);
        Dispatcher.BeginInvoke(UpdateHostArea);
    }

    // 宿主区域与放置区域

    // 将 HostPanel 的屏幕坐标换算为宿主客户区坐标，更新到 LayoutService
    private void UpdateHostArea()
    {
        if (!IsLoaded || HostPanel.ActualWidth <= 0) return;

        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero) return;

        // 获取 HostPanel 左上角和右下角的屏幕坐标
        var screenTL = HostPanel.PointToScreen(new Point(0, 0));
        var screenBR = HostPanel.PointToScreen(new Point(HostPanel.ActualWidth, HostPanel.ActualHeight));

        // 将屏幕左上角转换为宿主客户区坐标
        var clientTL = new POINT { X = (int)screenTL.X, Y = (int)screenTL.Y };
        NativeMethods.ScreenToClient(hwnd, ref clientTL);

        // 宽高从屏幕坐标差计算，保持物理像素精度
        int w = (int)(screenBR.X - screenTL.X);
        int h = (int)(screenBR.Y - screenTL.Y);

        Logger.Debug($"UpdateHostArea: client=({clientTL.X},{clientTL.Y}) {w}×{h}px", Cat);

        _layout.SetHostArea(clientTL.X, clientTL.Y, w, h);
        _layout.ApplyLayout();
    }

    // 计算侧边栏的屏幕物理像素矩形，更新到 DragDetectionService 作为放置区域
    private void UpdateDropZone()
    {
        if (!IsLoaded || SidebarBorder.ActualWidth <= 0) return;

        var tl   = SidebarBorder.PointToScreen(new Point(0, 0));
        var br   = SidebarBorder.PointToScreen(
            new Point(SidebarBorder.ActualWidth, SidebarBorder.ActualHeight));
        var zone = new Rect(tl, br);

        Logger.Debug(
            $"UpdateDropZone: ({tl.X:F0},{tl.Y:F0}) {(br.X - tl.X):F0}×{(br.Y - tl.Y):F0}px", Cat);

        _dragDetection.SetDropZone(zone);
        SyncManagedWindowsToDetection();
    }

    // 将 WindowManagerService 中的托管句柄同步到 DragDetectionService，保持两者一致
    private void SyncManagedWindowsToDetection()
    {
        int before = _dragDetection.ManagedWindows.Count;
        _dragDetection.ManagedWindows.Clear();
        foreach (var w in _windowManager.ManagedWindows)
            _dragDetection.ManagedWindows.Add(w.Handle);

        Logger.Trace(
            $"SyncManagedWindowsToDetection: {before} → {_dragDetection.ManagedWindows.Count} 个句柄", Cat);
    }

    // 覆盖层（外部窗口拖入）

    // 在侧边栏位置显示蓝色拖入提示覆盖层
    private void ShowDropOverlay()
    {
        if (_overlay == null || !IsLoaded || SidebarBorder.ActualWidth <= 0) return;
        var tl   = SidebarBorder.PointToScreen(new Point(0, 0));
        var rect = new Rect(tl, new Size(SidebarBorder.ActualWidth, SidebarBorder.ActualHeight));
        Logger.Trace($"ShowDropOverlay at ({rect.Left:F0},{rect.Top:F0})", Cat);
        _overlay.ShowAtRect(rect);
    }

    // 隐藏蓝色拖入提示覆盖层
    private void HideDropOverlay()
    {
        Logger.Trace("HideDropOverlay", Cat);
        _overlay?.HideOverlay();
    }

    // 覆盖层（托管窗口拖拽）

    // 在侧边栏位置显示红色解除托管提示覆盖层
    private void ShowReleaseOverlay()
    {
        if (_releaseOverlay == null || !IsLoaded || SidebarBorder.ActualWidth <= 0) return;
        var tl   = SidebarBorder.PointToScreen(new Point(0, 0));
        var rect = new Rect(tl, new Size(SidebarBorder.ActualWidth, SidebarBorder.ActualHeight));
        Logger.Trace($"ShowReleaseOverlay at ({rect.Left:F0},{rect.Top:F0})", Cat);
        _releaseOverlay.ShowAtRect(rect);
    }

    // 隐藏红色解除托管提示覆盖层
    private void HideReleaseOverlay()
    {
        Logger.Trace("HideReleaseOverlay", Cat);
        _releaseOverlay?.HideOverlay();
    }

    // 外部拖拽检测回调

    /// <summary>
    /// 外部窗口被拖入放置区域后释放，将其纳入托管并同步句柄集合。
    /// </summary>
    /// <param name="hwnd">被拖入的外部窗口句柄。</param>
    private void OnWindowDroppedInZone(IntPtr hwnd)
    {
        Logger.Info($"OnWindowDroppedInZone  hwnd=0x{hwnd:X}", Cat);

        bool hasPre = _dragDetection.PreDragRects.TryGetValue(hwnd, out RECT preDragRect);
        Logger.Debug(hasPre
            ? $"  PreDragRect 捕获成功: ({preDragRect.Left},{preDragRect.Top},{preDragRect.Width}×{preDragRect.Height})"
            : "  PreDragRect 未找到（将使用 SaveOriginalState 记录的位置）", Cat);

        Dispatcher.BeginInvoke(() =>
        {
            _windowManager.TryManageWindow(hwnd);
            _dragDetection.ManagedWindows.Add(hwnd);

            // hasPre / preDragRect 已在外部同步捕获，此处直接使用
            if (hasPre)
            {
                var window = _windowManager.FindByHandle(hwnd);
                if (window != null)
                {
                    window.OriginalRect = preDragRect;
                    Logger.Debug(
                        $"  OriginalRect 已更正为拖拽前位置: " +
                        $"({preDragRect.Left},{preDragRect.Top},{preDragRect.Width}×{preDragRect.Height})", Cat);
                }
            }
        });
    }
 
    /// <summary>
    /// 已托管窗口通过 MoveSize 事件被拖出管理区域，解除其托管。
    /// </summary>
    /// <param name="hwnd">被拖出区域的窗口句柄。</param>
    private void OnWindowDraggedOutOfZone(IntPtr hwnd)
    {
        Logger.Info($"OnWindowDraggedOutOfZone  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            _windowManager.ReleaseWindow(hwnd);
            _dragDetection.ManagedWindows.Remove(hwnd);
        });
    }

    // 托管窗口拖拽回调

    /// <summary>
    /// 托管窗口开始被拖拽时显示释放覆盖层并更新状态栏提示。
    /// </summary>
    /// <param name="hwnd">开始拖拽的托管窗口句柄。</param>
    private void OnManagedDragStarted(IntPtr hwnd)
    {
        Logger.Debug($"OnManagedDragStarted  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            ShowReleaseOverlay();
            SetStatus("拖拽中 — 移至侧边栏松开可解除托管");
        });
    }

    /// <summary>
    /// 托管窗口拖拽过程中的光标位置更新，当前仅由 DragDetectionService 记录日志，无额外 UI 操作。
    /// </summary>
    /// <param name="hwnd">正在被拖拽的窗口句柄。</param>
    /// <param name="screenPt">当前光标的屏幕物理像素坐标。</param>
    private void OnManagedDragMoved(IntPtr hwnd, Point screenPt)
    {
        // 位置信息已在 DragDetectionService 以 Trace 级别持续记录，此处无需额外处理
    }

    /// <summary>
    /// 托管窗口被拖至侧边栏区域释放，解除托管并更新状态栏。
    /// </summary>
    /// <param name="hwnd">被释放的托管窗口句柄。</param>
    private void OnManagedWindowDroppedOnSidebar(IntPtr hwnd)
    {
        Logger.Info($"OnManagedWindowDroppedOnSidebar  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            _windowManager.ReleaseWindow(hwnd);
            _dragDetection.ManagedWindows.Remove(hwnd);
            var mw = _windowManager.FindByHandle(hwnd);
            SetStatus($"已释放：{mw?.Title ?? "窗口"}");
        });
    }

    /// <summary>
    /// 托管窗口拖拽未落在有效目标上，尝试互换位置或将窗口归位。
    /// </summary>
    /// <param name="hwnd">拖拽取消的托管窗口句柄。</param>
    private void OnManagedDragCancelled(IntPtr hwnd)
    {
        Logger.Debug($"OnManagedDragCancelled  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            // 堆叠模式下无槽位概念，直接重嵌并重排
            if (_layout.CurrentMode == LayoutService.LayoutMode.Stacked)
            {
                Logger.Debug("  堆叠模式下取消拖拽 → ReEmbedAndRelayout", Cat);
                ReEmbedAndRelayout(hwnd);
                return;
            }

            // 获取鼠标松开时的屏幕坐标，判断落在哪个槽位
            NativeMethods.GetCursorPos(out POINT cursorPt);
            var cursorPoint = new Point(cursorPt.X, cursorPt.Y);
            var hostHwnd    = new WindowInteropHelper(this).Handle;
            int targetSlot  = _layout.GetSlotIndexAtScreenPoint(cursorPoint, hostHwnd);

            Logger.Debug(
                $"  松开位置: ({cursorPt.X},{cursorPt.Y})  targetSlot={targetSlot}", Cat);

            var draggedWindow = _windowManager.FindByHandle(hwnd);
            if (draggedWindow == null)
            {
                Logger.Warning("  FindByHandle 返回 null，放弃处理。", Cat);
                return;
            }

            int originalSlot = draggedWindow.SlotIndex;

            // 落在不同槽位上且该槽位有其他窗口，执行互换
            if (targetSlot >= 0 && targetSlot != originalSlot)
            {
                var targetWindow = _layout.GetWindowAtSlotIndex(targetSlot);
                if (targetWindow != null && targetWindow.Handle != hwnd)
                {
                    Logger.Info(
                        $"  互换: \"{draggedWindow.Title}\"[{originalSlot}] ⇄ " +
                        $"\"{targetWindow.Title}\"[{targetSlot}]", Cat);

                    // 若被拖拽窗口处于临时脱嵌状态，先重嵌再互换
                    if (!draggedWindow.IsEmbedded)
                        _windowManager.ReEmbed(hwnd);

                    _windowManager.SwapOrder(hwnd, targetWindow.Handle);
                    _layout.ApplyLayout();
                    SyncManagedWindowsToDetection();

                    SetStatus($"已互换位置：{draggedWindow.Title}  ⇄  {targetWindow.Title}");
                    return;
                }
            }

            // 未命中有效目标，将窗口送回原位
            Logger.Debug("  未命中有效目标 → ReEmbedAndRelayout（回到原位）", Cat);
            ReEmbedAndRelayout(hwnd);
            SetStatus("拖拽已取消，窗口回到原位");
        });
    }

    /// <summary>
    /// 托管窗口拖拽结束（无论结果如何），隐藏释放覆盖层。
    /// </summary>
    /// <param name="hwnd">完成拖拽的托管窗口句柄。</param>
    private void OnManagedDragEnded(IntPtr hwnd)
    {
        Logger.Debug($"OnManagedDragEnded  hwnd=0x{hwnd:X}  → 隐藏释放覆盖层", Cat);
        Dispatcher.BeginInvoke(HideReleaseOverlay);
    }

    /// <summary>
    /// 将临时脱嵌的窗口重新嵌入宿主并重新应用布局，用于拖拽取消时的回位。
    /// </summary>
    /// <param name="hwnd">需要重嵌的窗口句柄。</param>
    private void ReEmbedAndRelayout(IntPtr hwnd)
    {
        Logger.Debug($"ReEmbedAndRelayout  hwnd=0x{hwnd:X}", Cat);
        var window = _windowManager.FindByHandle(hwnd);
        if (window != null && !window.IsEmbedded)
        {
            Logger.Debug($"  调用 ReEmbed: \"{window.Title}\"", Cat);
            _windowManager.ReEmbed(hwnd);
        }
        _layout.ApplyLayout();
        SyncManagedWindowsToDetection();
    }

    // 编程式活跃窗口拖拽（Ctrl+Alt+M）

    // 热键触发的编程式拖拽：找到当前激活的托管窗口，临时脱嵌并进入系统 Move 循环
    private void StartActiveWindowDrag()
    {
        Logger.Debug("StartActiveWindowDrag 热键触发", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            // 堆叠模式下只有一个窗口时拖拽无意义，直接提示
            if (_layout.CurrentMode == LayoutService.LayoutMode.Stacked &&
                _windowManager.ManagedWindows.Count <= 1)
            {
                Logger.Warning("堆叠模式仅有一个窗口，无需拖拽。", Cat);
                SetStatus("堆叠模式下仅有一个窗口，无需拖拽");
                return;
            }

            // 优先使用激活窗口，没有激活窗口则取列表第一个
            var activeWindow = _windowManager.ManagedWindows.FirstOrDefault(w => w.IsActive)
                             ?? _windowManager.ManagedWindows.FirstOrDefault();

            if (activeWindow == null)
            {
                Logger.Warning("没有可拖拽的托管窗口。", Cat);
                SetStatus("没有可拖拽的托管窗口");
                return;
            }

            int originalSlot = activeWindow.SlotIndex;
            Logger.Info(
                $"发起编程式拖拽: \"{activeWindow.Title}\"  hwnd=0x{activeWindow.Handle:X}  " +
                $"slot={originalSlot}", Cat);

            if (_windowManager.StartProgrammaticDrag(activeWindow.Handle))
            {
                // 通知 DragDetectionService 当前为编程式拖拽，使其正确处理后续事件
                _dragDetection.NotifyProgrammaticDragStarted(activeWindow.Handle, originalSlot);
                SetStatus($"拖拽 \"{activeWindow.Title}\" — 移动到目标位置后松开鼠标");
            }
            else
            {
                Logger.Error($"StartProgrammaticDrag 失败: \"{activeWindow.Title}\"", Cat);
                SetStatus($"无法发起拖拽：\"{activeWindow.Title}\"");
            }
        });
    }

    // 侧边栏条目拖拽（WPF DragDrop）

    /// <summary>
    /// 记录鼠标按下时的坐标，用于后续判断是否超过拖拽启动阈值。
    /// </summary>
    /// <param name="sender">触发事件的边框元素。</param>
    /// <param name="e">鼠标按下事件参数，包含坐标信息。</param>
    private void WindowItemBorder_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        => _dragItemStartPoint = e.GetPosition(null);

    /// <summary>
    /// 鼠标移动超过系统拖拽阈值后，启动 WPF DragDrop 操作。
    /// </summary>
    /// <param name="sender">触发事件的边框元素，DataContext 应为 <see cref="ManagedWindow"/>。</param>
    /// <param name="e">鼠标移动事件参数。</param>
    private void WindowItemBorder_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _isItemDragInProgress) return;

        // 计算移动距离，未超过系统阈值时不启动拖拽
        var diff = e.GetPosition(null) - _dragItemStartPoint;
        if (Math.Abs(diff.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(diff.Y) < SystemParameters.MinimumVerticalDragDistance) return;

        if (sender is not Border border || border.DataContext is not ManagedWindow mw) return;

        Logger.Debug($"侧边栏条目拖拽开始: \"{mw.Title}\"  hwnd=0x{mw.Handle:X}", Cat);
        _isItemDragInProgress = true;

        // 显示底部释放区，DoDragDrop 期间用户可以将条目拖至此处解除托管
        SidebarReleaseZone.Visibility = Visibility.Visible;

        // DoDragDrop 是同步阻塞调用，直到用户松开鼠标才返回
        DragDrop.DoDragDrop(border,
            new DataObject(typeof(ManagedWindow), mw),
            DragDropEffects.Move);

        // DoDragDrop 返回后恢复释放区的隐藏状态
        SidebarReleaseZone.Visibility = Visibility.Collapsed;
        _isItemDragInProgress = false;
        Logger.Trace("侧边栏条目拖拽结束（DoDragDrop 返回）。", Cat);
    }

    /// <summary>
    /// 拖拽条目悬停在目标条目上时，验证可互换性并给予视觉反馈。
    /// </summary>
    /// <param name="sender">被悬停的目标条目边框。</param>
    /// <param name="e">拖拽事件参数，包含拖拽数据和效果设置。</param>
    private void WindowItemBorder_DragOver(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(typeof(ManagedWindow)))
        {
            e.Effects = DragDropEffects.None;
            e.Handled = true;
            return;
        }

        if (sender is not Border border || border.DataContext is not ManagedWindow target)
        {
            e.Handled = true;
            return;
        }

        var dragged  = e.Data.GetData(typeof(ManagedWindow)) as ManagedWindow;
        // 拖拽的来源和目标必须不同才允许互换
        bool canSwap = dragged != null && dragged != target;

        if (canSwap)
        {
            e.Effects         = DragDropEffects.Move;
            // 高亮目标条目背景作为视觉反馈
            border.Background = (Brush)FindResource("BgHover");
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    /// <summary>
    /// 拖拽离开目标条目时清除悬停高亮背景。
    /// </summary>
    /// <param name="sender">离开的目标条目边框。</param>
    /// <param name="e">拖拽事件参数。</param>
    private void WindowItemBorder_DragLeave(object sender, DragEventArgs e)
    {
        // 清除高亮，恢复默认背景
        if (sender is Border border)
            border.ClearValue(Border.BackgroundProperty);
    }

    /// <summary>
    /// 在目标条目上松开鼠标时，执行两个窗口的位置互换。
    /// </summary>
    /// <param name="sender">接收放置的目标条目边框。</param>
    /// <param name="e">包含被拖拽数据的拖拽事件参数。</param>
    private void WindowItemBorder_Drop(object sender, DragEventArgs e)
    {
        // 清除悬停高亮
        if (sender is Border b) b.ClearValue(Border.BackgroundProperty);

        if (!e.Data.GetDataPresent(typeof(ManagedWindow))) return;

        var dragged = e.Data.GetData(typeof(ManagedWindow)) as ManagedWindow;
        var target  = (sender as FrameworkElement)?.DataContext as ManagedWindow;

        if (dragged == null || target == null || dragged == target) return;

        Logger.Info(
            $"侧边栏 Drop 互换: \"{dragged.Title}\" ⇄ \"{target.Title}\"  " +
            $"mode={_layout.CurrentMode}", Cat);

        _windowManager.SwapOrder(dragged.Handle, target.Handle);
        _layout.ApplyLayout();
        SyncManagedWindowsToDetection();

        // 堆叠模式下的互换调整的是切换顺序，而非视觉位置
        string modeNote = _layout.CurrentMode == LayoutService.LayoutMode.Stacked
            ? "（堆叠模式：切换顺序已调整）" : string.Empty;
        SetStatus($"已交换位置：{dragged.Title}  ⇄  {target.Title}  {modeNote}".TrimEnd());

        e.Handled = true;
    }

    /// <summary>
    /// 拖拽条目悬停在底部释放区时，允许放置操作。
    /// </summary>
    /// <param name="sender">底部释放区边框。</param>
    /// <param name="e">拖拽事件参数。</param>
    private void SidebarReleaseZone_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(typeof(ManagedWindow))
            ? DragDropEffects.Move
            : DragDropEffects.None;
        e.Handled = true;
    }

    /// <summary>
    /// 条目拖拽至底部释放区时，解除对应窗口的托管。
    /// </summary>
    /// <param name="sender">底部释放区边框。</param>
    /// <param name="e">包含被拖拽数据的拖拽事件参数。</param>
    private void SidebarReleaseZone_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(typeof(ManagedWindow)) is ManagedWindow mw)
        {
            Logger.Info($"SidebarReleaseZone Drop: 释放 \"{mw.Title}\"", Cat);
            _windowManager.ReleaseWindow(mw.Handle);
            SetStatus($"已释放：{mw.Title}");
            e.Handled = true;
        }
    }

    // 工具栏按钮事件

    /// <summary>
    /// 工具栏布局模式按钮点击，从 Tag 解析目标布局模式并应用。
    /// </summary>
    /// <param name="sender">被点击的按钮，Tag 属性对应 <see cref="LayoutService.LayoutMode"/> 枚举名。</param>
    /// <param name="e">路由事件参数。</param>
    private void LayoutMode_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn &&
            Enum.TryParse<LayoutService.LayoutMode>(btn.Tag?.ToString(), out var mode))
        {
            Logger.Debug($"LayoutMode 按钮点击: {mode}", Cat);
            SetLayoutMode(mode);
        }
    }

    /// <summary>
    /// 切换布局模式，重新应用布局并高亮对应工具栏按钮。
    /// </summary>
    /// <param name="mode">要切换到的目标布局模式。</param>
    private void SetLayoutMode(LayoutService.LayoutMode mode)
    {
        Logger.Info($"SetLayoutMode: {mode}", Cat);
        _layout.CurrentMode = mode;
        _layout.ApplyLayout();
        HighlightLayoutButton(mode);
        SetStatus($"布局模式：{GetLayoutName(mode)}");
    }

    /// <summary>
    /// 更新工具栏布局模式按钮的激活高亮，当前模式按钮显示强调色。
    /// </summary>
    /// <param name="mode">当前激活的布局模式。</param>
    private void HighlightLayoutButton(LayoutService.LayoutMode mode)
    {
        var defaultBg = (Brush)FindResource("BgTertiary");
        var activeBg  = (Brush)FindResource("AccentBrush");
        // 仅当前模式的按钮使用强调色，其余恢复默认背景
        BtnStacked.Background   = mode == LayoutService.LayoutMode.Stacked   ? activeBg : defaultBg;
        BtnQuad.Background      = mode == LayoutService.LayoutMode.QuadSplit  ? activeBg : defaultBg;
        BtnLeftRight.Background = mode == LayoutService.LayoutMode.LeftRight  ? activeBg : defaultBg;
        BtnTopBottom.Background = mode == LayoutService.LayoutMode.TopBottom  ? activeBg : defaultBg;
    }

    /// <summary>
    /// 返回布局模式对应的中文可读名称，用于状态栏显示。
    /// </summary>
    /// <param name="mode">目标布局模式。</param>
    /// <returns>中文名称字符串。</returns>
    private static string GetLayoutName(LayoutService.LayoutMode mode) => mode switch
    {
        LayoutService.LayoutMode.Stacked   => "堆叠抽屉",
        LayoutService.LayoutMode.QuadSplit => "四分区",
        LayoutService.LayoutMode.LeftRight => "左右分屏",
        LayoutService.LayoutMode.TopBottom => "上下分屏",
        _                                  => mode.ToString()
    };

    /// <summary>
    /// 重排按钮点击，重新计算宿主区域并应用布局。
    /// </summary>
    /// <param name="sender">触发事件的按钮。</param>
    /// <param name="e">路由事件参数。</param>
    private void RefreshLayout_Click(object sender, RoutedEventArgs e)
    {
        Logger.Debug("RefreshLayout 按钮点击", Cat);
        UpdateHostArea();
        SetStatus("布局已重排");
    }

    /// <summary>
    /// 全部释放按钮点击，解除所有托管窗口。
    /// </summary>
    /// <param name="sender">触发事件的按钮。</param>
    /// <param name="e">路由事件参数。</param>
    private void ReleaseAll_Click(object sender, RoutedEventArgs e)
    {
        Logger.Info("ReleaseAll 按钮点击", Cat);
        _windowManager.ReleaseAll();
        SetStatus("已释放所有窗口");
    }

    /// <summary>
    /// 侧边栏折叠/展开按钮点击，切换侧边栏的可见状态。
    /// </summary>
    /// <param name="sender">触发事件的按钮。</param>
    /// <param name="e">路由事件参数。</param>
    private void ToggleSidebar_Click(object sender, RoutedEventArgs e)
    {
        Logger.Debug($"ToggleSidebar: expanded={_sidebarExpanded} → {!_sidebarExpanded}", Cat);
        if (_sidebarExpanded)
        {
            // 折叠：保存当前宽度，将侧边栏收窄为仅显示折叠按钮的宽度
            _savedSidebarWidth = SidebarColumn.ActualWidth;
            SidebarColumn.Width    = new GridLength(28);
            SidebarColumn.MinWidth = 0;
            SidebarContentPanel.Visibility = Visibility.Collapsed;
            SidebarTitleText.Visibility    = Visibility.Collapsed;
            BtnToggleSidebar.Content = "▶";
            BtnToggleSidebar.ToolTip = "展开侧边栏";
            _sidebarExpanded = false;
        }
        else
        {
            // 展开：恢复到折叠前保存的宽度
            SidebarColumn.Width    = new GridLength(_savedSidebarWidth);
            SidebarColumn.MinWidth = 28;
            SidebarContentPanel.Visibility = Visibility.Visible;
            SidebarTitleText.Visibility    = Visibility.Visible;
            BtnToggleSidebar.Content = "◀";
            BtnToggleSidebar.ToolTip = "折叠侧边栏";
            _sidebarExpanded = true;
        }
        // 侧边栏宽度变化后，宿主区域和放置区域坐标也随之改变
        Dispatcher.BeginInvoke(() => { UpdateHostArea(); UpdateDropZone(); });
    }

    // 窗口列表条目交互

    /// <summary>
    /// 点击侧边栏条目时切换到对应窗口，拖拽进行中时忽略点击。
    /// </summary>
    /// <param name="sender">触发事件的框架元素，DataContext 应为 <see cref="ManagedWindow"/>。</param>
    /// <param name="e">鼠标按钮事件参数。</param>
    private void WindowItem_Click(object sender, MouseButtonEventArgs e)
    {
        // 拖拽进行中时忽略点击，避免误触发切换
        if (_isItemDragInProgress) return;
        if (sender is FrameworkElement fe && fe.DataContext is ManagedWindow mw)
        {
            Logger.Debug($"WindowItem 点击切换: \"{mw.Title}\"  hwnd=0x{mw.Handle:X}", Cat);
            _layout.SwitchToWindow(mw);
        }
    }

    /// <summary>
    /// 条目右侧释放按钮点击，解除对应窗口的托管。
    /// </summary>
    /// <param name="sender">触发事件的按钮，Tag 属性保存窗口句柄。</param>
    /// <param name="e">路由事件参数，标记为已处理以防止冒泡触发行点击。</param>
    private void ReleaseButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is IntPtr hwnd)
        {
            Logger.Debug($"ReleaseButton 点击: hwnd=0x{hwnd:X}", Cat);
            _windowManager.ReleaseWindow(hwnd);
        }
        // 阻止事件冒泡，防止触发 WindowItem_Click
        e.Handled = true;
    }

    // 热键动作

    // 抓取当前前台窗口，排除宿主自身
    private void GrabForegroundWindow()
    {
        var hwnd = NativeMethods.GetForegroundWindow();
        var self = new WindowInteropHelper(this).Handle;
        Logger.Debug($"GrabForegroundWindow: foreground=0x{hwnd:X}  self=0x{self:X}", Cat);

        if (hwnd == IntPtr.Zero || hwnd == self)
        {
            Logger.Warning("GrabForegroundWindow: 前台窗口为空或为自身，忽略。", Cat);
            return;
        }
        Dispatcher.BeginInvoke(() => _windowManager.TryManageWindow(hwnd));
    }

    // 缩略图预览交互

    /// <summary>
    /// 鼠标进入侧边栏条目时，记录目标及坐标，并启动防抖计时器。
    /// 若计时器已运行（快速在条目间移动），则重置倒计时并更新目标。
    /// </summary>
    /// <param name="sender">被悬停的 <see cref="Border"/>，DataContext 为 <see cref="ManagedWindow"/>。</param>
    /// <param name="e">鼠标事件参数。</param>
    private void SidebarItem_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (sender is not FrameworkElement fe || fe.DataContext is not ManagedWindow mw) return;

        // 拖拽期间不显示预览，避免视觉干扰
        if (_isItemDragInProgress) return;

        _hoverTarget       = mw;
        _hoverItemScreenTL = fe.PointToScreen(new Point(0, 0));
        _sidebarScreenRight = SidebarBorder.PointToScreen(
            new Point(SidebarBorder.ActualWidth, 0)).X;

        // 重置计时器（保证每次悬停都等满防抖时间再弹出）
        _hoverTimer?.Stop();
        _hoverTimer?.Start();

        Logger.Trace(
            $"SidebarItem_MouseEnter: \"{mw.Title}\"  " +
            $"screenTL=({_hoverItemScreenTL.X:F0},{_hoverItemScreenTL.Y:F0})", Cat);
    }

    /// <summary>
    /// 鼠标离开单个条目时不立即隐藏预览（可能只是移向相邻条目）；
    /// 由 <see cref="WindowListBox_MouseLeave"/> 在彻底离开列表时统一隐藏。
    /// </summary>
    private void SidebarItem_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
    {
        // 无需操作：只要鼠标还在 ListBox 范围内，预览应保持可见。
        // 完全离开列表后 WindowListBox_MouseLeave 会负责清理。
    }

    /// <summary>
    /// 鼠标完全离开窗口列表时，停止防抖计时器并隐藏预览浮窗。
    /// </summary>
    private void WindowListBox_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
    {
        Logger.Trace("WindowListBox_MouseLeave → 停止计时器 + 隐藏预览", Cat);
        _hoverTimer?.Stop();
        _hoverTarget = null;
        _thumbnailPreview?.Hide();
    }

    /// <summary>
    /// 防抖计时器触发时，显示当前悬停目标的缩略图预览。
    /// 计时器为单次触发（Tick 后立即停止）。
    /// </summary>
    private void HoverTimer_Tick(object? sender, EventArgs e)
    {
        // 计时器是持续触发的 DispatcherTimer，手动停止使其仅触发一次
        _hoverTimer?.Stop();

        if (_hoverTarget == null || _thumbnailPreview == null) return;

        // 从 PresentationSource 读取 DPI 缩放比例
        // PointToScreen 返回物理像素，Window.Left/Top 需要逻辑像素，必须换算
        var source   = PresentationSource.FromVisual(this);
        double dpiScale = source?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;

        Logger.Debug(
            $"HoverTimer 触发 → 显示预览: \"{_hoverTarget.Title}\"  " +
            $"dpiScale={dpiScale:F2}  sidebarRight={_sidebarScreenRight:F0}", Cat);

        _thumbnailPreview.Show(_hoverTarget, _hoverItemScreenTL, _sidebarScreenRight, dpiScale);
    }
    
    // 焦点修复
 
    /// <summary>
    /// 宿主窗口的 WndProc 钩子。
    /// 拦截 <c>WM_PARENTNOTIFY(WM_LBUTTONDOWN)</c>：当用户点击任意嵌入子窗口时，
    /// 立即调用 <see cref="WindowManagerService.FocusEmbeddedWindow"/> 将键盘焦点
    /// 转移到该子窗口，解决跨进程嵌入后无法键盘输入的问题。
    /// </summary>
    private IntPtr HostWndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == NativeConstants.WM_PARENTNOTIFY)
        {
            // WM_PARENTNOTIFY 的 wParam 低字节是子消息类型
            int childMsg = wParam.ToInt32() & 0xFFFF;
            if (childMsg == NativeConstants.WM_LBUTTONDOWN)
            {
                // lParam 包含鼠标在宿主客户区的坐标（物理像素）
                int clientX = (short)(lParam.ToInt64() & 0xFFFF);
                int clientY = (short)((lParam.ToInt64() >> 16) & 0xFFFF);
 
                // 找出该坐标下被点击的托管窗口
                var target = FindManagedWindowAtClientPoint(clientX, clientY);
                if (target != null)
                {
                    Logger.Debug(
                        $"WM_PARENTNOTIFY click → FocusEmbedded: \"{target.Title}\"  " +
                        $"client=({clientX},{clientY})", Cat);
                    _windowManager.FocusEmbeddedWindow(target.Handle);
 
                    // 同步激活状态
                    foreach (var w in _windowManager.ManagedWindows)
                        w.IsActive = w.Handle == target.Handle;
                }
            }
        }
        return IntPtr.Zero;
    }
 
    /// <summary>
    /// 根据宿主客户区坐标，向上遍历窗口父链，找到对应的托管窗口。
    /// 嵌入窗口内部可能有多层子控件（如文本框、标签页等），需要逐级向上查找。
    /// </summary>
    /// <param name="clientX">宿主客户区 X 坐标（物理像素）。</param>
    /// <param name="clientY">宿主客户区 Y 坐标（物理像素）。</param>
    /// <returns>找到的托管窗口，未找到时返回 null。</returns>
    private ManagedWindow? FindManagedWindowAtClientPoint(int clientX, int clientY)
    {
        // 将宿主客户区坐标转换为屏幕坐标，再用 WindowFromPoint 定位实际被点击的 HWND
        var pt = new POINT { X = clientX, Y = clientY };
        NativeMethods.ClientToScreen(_windowManager.HostHwnd, ref pt);
        IntPtr clickedHwnd = NativeMethods.WindowFromPoint(pt);
 
        if (clickedHwnd == IntPtr.Zero) return null;
 
        // 向上遍历父链，直到找到托管窗口或到达宿主
        IntPtr current = clickedHwnd;
        while (current != IntPtr.Zero && current != _windowManager.HostHwnd)
        {
            var managed = _windowManager.FindByHandle(current);
            if (managed != null) return managed;
            current = NativeMethods.GetParent(current);
        }
        return null;
    }

    // UI 辅助

    /// <summary>
    /// 托管窗口列表内容变化时更新侧边栏的空状态提示、计数文字和句柄同步。
    /// </summary>
    /// <param name="sender">发生变化的 ObservableCollection。</param>
    /// <param name="e">集合变化事件参数，包含变化类型和受影响的元素。</param>
    private void OnManagedWindowsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        Logger.Trace(
            $"ManagedWindows CollectionChanged: action={e.Action}  " +
            $"count={_windowManager.ManagedWindows.Count}", Cat);

        Dispatcher.BeginInvoke(() =>
        {
            bool has = _windowManager.ManagedWindows.Count > 0;
            // 有窗口时显示列表，隐藏空状态提示；无窗口时反之
            WindowListBox.Visibility   = has ? Visibility.Visible   : Visibility.Collapsed;
            EmptyStatePanel.Visibility = has ? Visibility.Collapsed : Visibility.Visible;
            DropHintPanel.Visibility   = has ? Visibility.Collapsed : Visibility.Visible;
            WindowCountText.Text = has
                ? $"{_windowManager.ManagedWindows.Count} 个窗口"
                : "0 个窗口";

            // 列表变化后同步句柄到拖拽检测服务
            SyncManagedWindowsToDetection();
        });
    }

    /// <summary>
    /// 更新底部状态栏文字。
    /// </summary>
    /// <param name="msg">要显示的状态消息。</param>
    private void SetStatus(string msg)
    {
        Logger.Trace($"SetStatus: {msg}", Cat);
        StatusText.Text = msg;
    }
}