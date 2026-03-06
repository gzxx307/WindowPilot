using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using WindowPilot.Controls;
using WindowPilot.Models;
using WindowPilot.Native;
using WindowPilot.Services;

namespace WindowPilot;

public partial class MainWindow : Window
{
    // ── 服务 ──
    private readonly WinEventHookService   _winEventHook  = new();
    private readonly WindowManagerService  _windowManager;
    private readonly DragDetectionService  _dragDetection;
    private readonly LayoutService         _layout;
    private readonly HotkeyService         _hotkey        = new();

    // ── 覆盖层（拖拽外部窗口时显示在侧边栏上方） ──
    private DropZoneOverlay? _overlay;

    // ── 侧边栏状态 ──
    private bool   _sidebarExpanded    = true;
    private double _savedSidebarWidth  = 200;

    // ── 侧边栏条目拖拽状态 ──
    private Point _dragItemStartPoint;
    private bool  _isItemDragInProgress;

    // ────────────────────────── 构造函数 ──────────────────────────

    public MainWindow()
    {
        InitializeComponent();

        _windowManager = new WindowManagerService(_winEventHook);
        _dragDetection = new DragDetectionService(_winEventHook);
        _layout        = new LayoutService(_windowManager);

        // 绑定窗口列表
        WindowListBox.ItemsSource = _windowManager.ManagedWindows;
        _windowManager.ManagedWindows.CollectionChanged += OnManagedWindowsChanged;

        // 管理器事件
        _windowManager.WindowManaged  += w => Dispatcher.BeginInvoke(() => SetStatus($"已接管：{w.Title}"));
        _windowManager.WindowReleased += w => Dispatcher.BeginInvoke(() => SetStatus($"已释放：{w.Title}"));
        _windowManager.LayoutChanged  += () => Dispatcher.BeginInvoke(() => _layout.ApplyLayout());
        _windowManager.ManageFailed   += (_, msg) => Dispatcher.BeginInvoke(() => SetStatus(msg));

        // 拖拽检测事件（外部窗口拖入侧边栏）
        _dragDetection.ExternalDragStarted    += _ => Dispatcher.BeginInvoke(ShowDropOverlay);
        _dragDetection.DragEnded              += _ => Dispatcher.BeginInvoke(HideDropOverlay);
        _dragDetection.DragMoved              += (_, pt) => Dispatcher.BeginInvoke(() => UpdateDebugMousePos(pt));
        _dragDetection.WindowDroppedInZone    += OnWindowDroppedInZone;
        _dragDetection.WindowDraggedOutOfZone += OnWindowDraggedOutOfZone;
    }

    // ────────────────────────── 生命周期 ──────────────────────────

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        _windowManager.HostHwnd = new WindowInteropHelper(this).Handle;
        _winEventHook.Start();

        // 注册全局热键
        _hotkey.Attach(this);
        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x47 /* G */,
            GrabForegroundWindow);
        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x25 /* ← */,
            () => Dispatcher.BeginInvoke(_layout.SwitchToPrevious));
        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x27 /* → */,
            () => Dispatcher.BeginInvoke(_layout.SwitchToNext));
        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x51 /* Q */,
            () => Dispatcher.BeginInvoke(() => SetLayoutMode(LayoutService.LayoutMode.QuadSplit)));

        _overlay = new DropZoneOverlay();

        // 初始化宿主区域和放置区域
        Dispatcher.BeginInvoke(() => { UpdateHostArea(); UpdateDropZone(); });

        SetStatus("就绪 — 拖拽窗口到侧边栏以接管，或按 Ctrl+Alt+G 抓取当前窗口");
    }

    private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        _hotkey.Dispose();
        _windowManager.ReleaseAll();
        _windowManager.Dispose();
        _dragDetection.Dispose();
        _winEventHook.Dispose();
        _overlay?.Close();
    }

    private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        => Dispatcher.BeginInvoke(() => { UpdateHostArea(); UpdateDropZone(); });

    private void HostPanel_SizeChanged(object sender, SizeChangedEventArgs e)
        => Dispatcher.BeginInvoke(UpdateHostArea);

    // ────────────────────────── 宿主区域 / 放置区域 ──────────────────────────

    private void UpdateHostArea()
    {
        if (!IsLoaded || HostPanel.ActualWidth <= 0) return;

        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero) return;

        // PointToScreen 返回物理像素屏幕坐标；ScreenToClient 转换为宿主客户区坐标
        var screenTL = HostPanel.PointToScreen(new Point(0, 0));
        var screenBR = HostPanel.PointToScreen(new Point(HostPanel.ActualWidth, HostPanel.ActualHeight));

        var clientTL = new POINT { X = (int)screenTL.X, Y = (int)screenTL.Y };
        NativeMethods.ScreenToClient(hwnd, ref clientTL);

        int w = (int)(screenBR.X - screenTL.X);
        int h = (int)(screenBR.Y - screenTL.Y);

        _layout.SetHostArea(clientTL.X, clientTL.Y, w, h);
        _layout.ApplyLayout();

        UpdateDebugHostArea();
    }

    private void UpdateDropZone()
    {
        if (!IsLoaded || SidebarBorder.ActualWidth <= 0) return;

        var tl = SidebarBorder.PointToScreen(new Point(0, 0));
        var br = SidebarBorder.PointToScreen(
            new Point(SidebarBorder.ActualWidth, SidebarBorder.ActualHeight));

        _dragDetection.SetDropZone(new Rect(tl, br));

        // 保持 ManagedWindows 同步
        _dragDetection.ManagedWindows.Clear();
        foreach (var w in _windowManager.ManagedWindows)
            _dragDetection.ManagedWindows.Add(w.Handle);

        UpdateDebugDropZone(new Rect(tl, br));
    }

    // ────────────────────────── 覆盖层 ──────────────────────────

    private void ShowDropOverlay()
    {
        if (_overlay == null || !IsLoaded || SidebarBorder.ActualWidth <= 0) return;
        var tl = SidebarBorder.PointToScreen(new Point(0, 0));
        _overlay.ShowAtRect(new Rect(tl,
            new Size(SidebarBorder.ActualWidth, SidebarBorder.ActualHeight)));
    }

    private void HideDropOverlay() => _overlay?.HideOverlay();

    // ────────────────────────── 外部拖拽检测回调 ──────────────────────────

    private void OnWindowDroppedInZone(IntPtr hwnd)
    {
        Dispatcher.BeginInvoke(() =>
        {
            _windowManager.TryManageWindow(hwnd);
            _dragDetection.ManagedWindows.Add(hwnd);
        });
    }

    private void OnWindowDraggedOutOfZone(IntPtr hwnd)
    {
        Dispatcher.BeginInvoke(() =>
        {
            _windowManager.ReleaseWindow(hwnd);
            _dragDetection.ManagedWindows.Remove(hwnd);
        });
    }

    // ════════════════════════════════════════════════════════════
    // ✨ 侧边栏条目拖拽：互换槽位 / 拖至释放区解除托管
    // ════════════════════════════════════════════════════════════

    /// <summary>记录鼠标按下位置，用于判断是否触发拖拽</summary>
    private void WindowItemBorder_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        => _dragItemStartPoint = e.GetPosition(null);

    /// <summary>
    /// 鼠标移动超过系统拖拽阈值后启动 WPF DragDrop。
    /// 拖拽期间底部出现红色释放区域。
    /// </summary>
    private void WindowItemBorder_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _isItemDragInProgress) return;

        var diff = e.GetPosition(null) - _dragItemStartPoint;
        if (Math.Abs(diff.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(diff.Y) < SystemParameters.MinimumVerticalDragDistance) return;

        if (sender is not Border border || border.DataContext is not ManagedWindow mw) return;

        _isItemDragInProgress = true;

        // 显示底部释放区域
        SidebarReleaseZone.Visibility = Visibility.Visible;

        // 启动拖拽（阻塞直到完成）
        DragDrop.DoDragDrop(border,
            new DataObject(typeof(ManagedWindow), mw),
            DragDropEffects.Move);

        // 拖拽结束后清理
        SidebarReleaseZone.Visibility = Visibility.Collapsed;
        _isItemDragInProgress = false;
    }

    /// <summary>拖拽经过某条目时高亮显示（堆叠模式下不允许换位）</summary>
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

        var dragged = e.Data.GetData(typeof(ManagedWindow)) as ManagedWindow;
        bool canSwap = dragged != null && dragged != target;

        if (canSwap)
        {
            e.Effects = DragDropEffects.Move;
            border.Background = (Brush)FindResource("BgHover");
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    /// <summary>鼠标离开时还原高亮（让 DataTrigger 的 IsActive 样式重新接管）</summary>
    private void WindowItemBorder_DragLeave(object sender, DragEventArgs e)
    {
        // ClearValue 移除本地值，DataTrigger（IsActive 高亮）可自动恢复
        if (sender is Border border)
            border.ClearValue(Border.BackgroundProperty);
    }

    /// <summary>
    /// 拖放到目标条目：交换两个窗口在列表中的顺序，重新应用布局。
    /// 所有布局模式均支持：
    ///   · 分屏模式 → 视觉槽位互换
    ///   · 堆叠模式 → 切换循环顺序互换
    /// </summary>
    private void WindowItemBorder_Drop(object sender, DragEventArgs e)
    {
        // 还原高亮
        if (sender is Border b) b.ClearValue(Border.BackgroundProperty);

        if (!e.Data.GetDataPresent(typeof(ManagedWindow))) return;

        var dragged = e.Data.GetData(typeof(ManagedWindow)) as ManagedWindow;
        var target  = (sender as FrameworkElement)?.DataContext as ManagedWindow;

        if (dragged == null || target == null || dragged == target) return;

        // 交换列表顺序并重新布局
        _windowManager.SwapOrder(dragged.Handle, target.Handle);
        _layout.ApplyLayout();

        string modeNote = _layout.CurrentMode == LayoutService.LayoutMode.Stacked
            ? "（堆叠模式：切换顺序已调整）"
            : string.Empty;
        SetStatus($"已交换位置：{dragged.Title}  ⇄  {target.Title}  {modeNote}".TrimEnd());

        e.Handled = true;
    }

    /// <summary>拖拽经过释放区时显示允许放置</summary>
    private void SidebarReleaseZone_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(typeof(ManagedWindow))
            ? DragDropEffects.Move
            : DragDropEffects.None;
        e.Handled = true;
    }

    /// <summary>放置到释放区：解除该窗口的托管，窗口回到原位置</summary>
    private void SidebarReleaseZone_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(typeof(ManagedWindow)) is ManagedWindow mw)
        {
            _windowManager.ReleaseWindow(mw.Handle);
            SetStatus($"已释放：{mw.Title}");
            e.Handled = true;
        }
    }

    // ────────────────────────── 工具栏按钮 ──────────────────────────

    private void LayoutMode_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn &&
            Enum.TryParse<LayoutService.LayoutMode>(btn.Tag?.ToString(), out var mode))
            SetLayoutMode(mode);
    }

    private void SetLayoutMode(LayoutService.LayoutMode mode)
    {
        _layout.CurrentMode = mode;
        _layout.ApplyLayout();
        HighlightLayoutButton(mode);
        SetStatus($"布局模式：{GetLayoutName(mode)}");
    }

    private void HighlightLayoutButton(LayoutService.LayoutMode mode)
    {
        var defaultBg = (Brush)FindResource("BgTertiary");
        var activeBg  = (Brush)FindResource("AccentBrush");
        BtnStacked.Background   = mode == LayoutService.LayoutMode.Stacked   ? activeBg : defaultBg;
        BtnQuad.Background      = mode == LayoutService.LayoutMode.QuadSplit  ? activeBg : defaultBg;
        BtnLeftRight.Background = mode == LayoutService.LayoutMode.LeftRight  ? activeBg : defaultBg;
        BtnTopBottom.Background = mode == LayoutService.LayoutMode.TopBottom  ? activeBg : defaultBg;
    }

    private static string GetLayoutName(LayoutService.LayoutMode mode) => mode switch
    {
        LayoutService.LayoutMode.Stacked   => "堆叠抽屉",
        LayoutService.LayoutMode.QuadSplit => "四分区",
        LayoutService.LayoutMode.LeftRight => "左右分屏",
        LayoutService.LayoutMode.TopBottom => "上下分屏",
        _                                  => mode.ToString()
    };

    private void RefreshLayout_Click(object sender, RoutedEventArgs e)
    {
        UpdateHostArea();
        SetStatus("布局已重排");
    }

    private void ReleaseAll_Click(object sender, RoutedEventArgs e)
    {
        _windowManager.ReleaseAll();
        SetStatus("已释放所有窗口");
    }

    private void ToggleDebug_Click(object sender, RoutedEventArgs e)
    {
        DebugPanel.Visibility = DebugPanel.Visibility == Visibility.Visible
            ? Visibility.Collapsed
            : Visibility.Visible;
    }

    private void ToggleSidebar_Click(object sender, RoutedEventArgs e)
    {
        if (_sidebarExpanded)
        {
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
            SidebarColumn.Width    = new GridLength(_savedSidebarWidth);
            SidebarColumn.MinWidth = 28;
            SidebarContentPanel.Visibility = Visibility.Visible;
            SidebarTitleText.Visibility    = Visibility.Visible;
            BtnToggleSidebar.Content = "◀";
            BtnToggleSidebar.ToolTip = "折叠侧边栏";
            _sidebarExpanded = true;
        }
        Dispatcher.BeginInvoke(() => { UpdateHostArea(); UpdateDropZone(); });
    }

    // ────────────────────────── 窗口列表条目 ──────────────────────────

    /// <summary>
    /// 点击（无拖拽）：切换到该窗口。
    /// 使用 MouseLeftButtonUp 以避免与拖拽手势冲突。
    /// </summary>
    private void WindowItem_Click(object sender, MouseButtonEventArgs e)
    {
        // 如果本次鼠标操作触发了拖拽，忽略点击
        if (_isItemDragInProgress) return;
        if (sender is FrameworkElement fe && fe.DataContext is ManagedWindow mw)
            _layout.SwitchToWindow(mw);
    }

    private void ReleaseButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is IntPtr hwnd)
            _windowManager.ReleaseWindow(hwnd);
        e.Handled = true; // 阻止冒泡到 WindowItem_Click
    }

    // ────────────────────────── 热键动作 ──────────────────────────

    private void GrabForegroundWindow()
    {
        var hwnd = NativeMethods.GetForegroundWindow();
        var self = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero || hwnd == self) return;
        Dispatcher.BeginInvoke(() => _windowManager.TryManageWindow(hwnd));
    }

    // ────────────────────────── UI 辅助 ──────────────────────────

    private void OnManagedWindowsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        Dispatcher.BeginInvoke(() =>
        {
            bool has = _windowManager.ManagedWindows.Count > 0;
            WindowListBox.Visibility   = has ? Visibility.Visible   : Visibility.Collapsed;
            EmptyStatePanel.Visibility = has ? Visibility.Collapsed : Visibility.Visible;
            DropHintPanel.Visibility   = has ? Visibility.Collapsed : Visibility.Visible;
            WindowCountText.Text = has
                ? $"{_windowManager.ManagedWindows.Count} 个窗口"
                : "0 个窗口";

            // 同步到拖拽检测服务
            _dragDetection.ManagedWindows.Clear();
            foreach (var w in _windowManager.ManagedWindows)
                _dragDetection.ManagedWindows.Add(w.Handle);
        });
    }

    private void SetStatus(string msg) => StatusText.Text = msg;

    // ────────────────────────── Debug 面板更新 ──────────────────────────

    private void UpdateDebugHostArea()
    {
        if (DebugPanel.Visibility != Visibility.Visible) return;
        var area = _layout.HostArea;
        DbgHostAreaText.Text = $"Left={area.Left}, Top={area.Top}";
        DbgHostSizeText.Text = $"{area.Width} × {area.Height} px";

        var hwnd = new WindowInteropHelper(this).Handle;
        uint dpi = hwnd != IntPtr.Zero ? NativeMethods.GetDpiForWindow(hwnd) : 96;
        DbgDpiText.Text = $"DPI={dpi}  Scale={dpi / 96.0:F2}×";
    }

    private void UpdateDebugMousePos(Point screenPt)
    {
        if (DebugPanel.Visibility != Visibility.Visible) return;
        DbgMousePhysText.Text = $"Mouse(phys) = {screenPt.X:F0}, {screenPt.Y:F0}";
    }

    private void UpdateDebugDropZone(Rect zone)
    {
        if (DebugPanel.Visibility != Visibility.Visible) return;
        DbgDropZonePhysText.Text = $"({zone.Left:F0},{zone.Top:F0}) {zone.Width:F0}×{zone.Height:F0}";
        var tl = SidebarBorder.PointToScreen(new Point(0, 0));
        DbgSidebarPhysText.Text = $"SidebarTL = {tl.X:F0}, {tl.Y:F0}";
    }
}