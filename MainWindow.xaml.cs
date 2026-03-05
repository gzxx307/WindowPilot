using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using WindowPilot.Controls;
using WindowPilot.Models;
using WindowPilot.Native;
using WindowPilot.Services;

namespace WindowPilot;

public partial class MainWindow : Window
{
    // ── 服务 ──
    private readonly WinEventHookService _winEventHook;
    private readonly DragDetectionService _dragDetection;
    private readonly WindowManagerService _windowManager;
    private readonly LayoutService _layoutService;
    private readonly HotkeyService _hotkeyService;

    // ── UI ──
    private DropZoneOverlay? _dropOverlay;
    private IntPtr _myHwnd;

    // ── DPI 缩放（从主窗口 PresentationSource 读取，物理px / DIP） ──
    private double _dpiX = 1.0;
    private double _dpiY = 1.0;

    // ── 当前放置区的物理像素矩形（供 Debug 用） ──
    private Rect _currentDropZonePhys = Rect.Empty;

    // ── 热键 ID ──
    private int _hotkeyGrab = -1;
    private int _hotkeyNext = -1;
    private int _hotkeyPrev = -1;
    private int _hotkeyQuad = -1;

    // ── 侧边栏折叠状态 ──
    private bool _sidebarExpanded = true;
    private double _savedSidebarWidth = 200;

    // ── Debug ──
    private bool _debugVisible = false;
    private DispatcherTimer? _debugTimer;

    public MainWindow()
    {
        InitializeComponent();

        _winEventHook = new WinEventHookService();
        _dragDetection = new DragDetectionService(_winEventHook);
        _windowManager = new WindowManagerService(_winEventHook);
        _layoutService = new LayoutService(_windowManager);
        _hotkeyService = new HotkeyService();

        WindowListBox.ItemsSource = _windowManager.ManagedWindows;

        _windowManager.ManagedWindows.CollectionChanged += ManagedWindows_Changed;
        _windowManager.ManageFailed += OnManageFailed;
        _windowManager.WindowManaged += OnWindowManaged;
        _windowManager.WindowReleased += OnWindowReleased;
        _windowManager.LayoutChanged += OnLayoutChanged;

        _dragDetection.ExternalDragStarted += OnExternalDragStarted;
        _dragDetection.DragMoved += OnDragMoved;
        _dragDetection.WindowDroppedInZone += OnWindowDroppedInZone;
        _dragDetection.DragEnded += OnDragEnded;

        Loaded += MainWindow_Loaded;
        Closing += MainWindow_Closing;

        HostPanel.SizeChanged += HostPanel_SizeChanged;
        SidebarBorder.SizeChanged += SidebarBorder_SizeChanged;
    }

    // ═══════════════════════════════════════════
    //  生命周期
    // ═══════════════════════════════════════════

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        _myHwnd = new WindowInteropHelper(this).Handle;
        _windowManager.HostHwnd = _myHwnd;

        // ── 读取 DPI 缩放比（物理像素 / DIP） ──
        // PresentationSource 是最可靠的来源；Per-Monitor DPI 变化时应重新读取
        RefreshDpiScale();

        long myStyle = NativeMethods.GetWindowLongSafe(_myHwnd, NativeConstants.GWL_STYLE);
        myStyle |= (long)NativeConstants.WS_CLIPCHILDREN;
        NativeMethods.SetWindowLongPtrSafe(_myHwnd, NativeConstants.GWL_STYLE, myStyle);

        _winEventHook.Start();
        _hotkeyService.Attach(this);
        RegisterHotkeys();

        _dropOverlay = new DropZoneOverlay();

        UpdateHostArea();
        UpdateDropZone();

        SetStatus("就绪 — 拖拽窗口到左侧侧边栏，或按 Ctrl+Alt+G 抓取当前活动窗口");
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        _debugTimer?.Stop();
        _windowManager.ReleaseAll();
        _hotkeyService.Dispose();
        _dragDetection.Dispose();
        _windowManager.Dispose();
        _winEventHook.Dispose();
        _dropOverlay?.Close();
    }

    // ═══════════════════════════════════════════
    //  DPI
    // ═══════════════════════════════════════════

    /// <summary>
    /// 从 PresentationSource 刷新 DPI 缩放比。
    /// TransformToDevice.M11 = 物理像素 / DIP（如 150% 时 = 1.5）。
    /// </summary>
    private void RefreshDpiScale()
    {
        var src = PresentationSource.FromVisual(this);
        if (src?.CompositionTarget != null)
        {
            _dpiX = src.CompositionTarget.TransformToDevice.M11;
            _dpiY = src.CompositionTarget.TransformToDevice.M22;
        }
    }

    // ═══════════════════════════════════════════
    //  宿主区域计算（精确像素坐标）
    // ═══════════════════════════════════════════

    private void HostPanel_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdateHostArea();
        _layoutService.ApplyLayout();
    }

    private void SidebarBorder_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdateDropZone();
    }

    /// <summary>
    /// 计算 HostPanel 在宿主窗口客户区中的精确物理像素矩形。
    ///
    /// 坐标换算链：
    ///   HostPanel 本地 DIP (0,0)
    ///   → PointToScreen → 屏幕物理像素
    ///   → ScreenToClient(_myHwnd) → 宿主客户区物理像素  ← MoveWindow 所用坐标系
    /// </summary>
    private void UpdateHostArea()
    {
        if (_myHwnd == IntPtr.Zero) return;
        try
        {
            var screenTL = HostPanel.PointToScreen(new Point(0, 0));
            var screenBR = HostPanel.PointToScreen(new Point(HostPanel.ActualWidth, HostPanel.ActualHeight));

            var ptTL = new POINT { X = (int)Math.Round(screenTL.X), Y = (int)Math.Round(screenTL.Y) };
            var ptBR = new POINT { X = (int)Math.Round(screenBR.X), Y = (int)Math.Round(screenBR.Y) };

            NativeMethods.ScreenToClient(_myHwnd, ref ptTL);
            NativeMethods.ScreenToClient(_myHwnd, ref ptBR);

            int pw = ptBR.X - ptTL.X;
            int ph = ptBR.Y - ptTL.Y;

            if (pw > 0 && ph > 0)
            {
                _layoutService.SetHostArea(ptTL.X, ptTL.Y, pw, ph);
                RefreshDebugHostArea(ptTL.X, ptTL.Y, pw, ph);
            }
        }
        catch { }
    }

    /// <summary>
    /// 将拖拽检测放置区设为侧边栏的屏幕区域（物理像素），
    /// 同时将 Overlay 定位到对应的 DIP 位置。
    ///
    /// ┌─────────────────────────────────────────────────────┐
    /// │ DragDetection._dropZone  → 物理像素（与 GetWindowRect 同系） │
    /// │ DropZoneOverlay.Left/Top → DIP     （WPF Window 坐标系）    │
    /// │ physicalRect / dpiX = dipRect                               │
    /// └─────────────────────────────────────────────────────┘
    /// </summary>
    private void UpdateDropZone()
    {
        if (_myHwnd == IntPtr.Zero) return;
        try
        {
            if (!_sidebarExpanded || SidebarBorder.ActualWidth < 30)
            {
                _dragDetection.SetDropZone(Rect.Empty);
                _currentDropZonePhys = Rect.Empty;
                return;
            }

            // PointToScreen 返回物理像素
            var screenTL = SidebarBorder.PointToScreen(new Point(0, 0));
            var screenBR = SidebarBorder.PointToScreen(
                new Point(SidebarBorder.ActualWidth, SidebarBorder.ActualHeight));

            var physicalRect = new Rect(screenTL, screenBR); // 物理像素

            // 拖拽判定用物理像素
            _dragDetection.SetDropZone(physicalRect);
            _currentDropZonePhys = physicalRect;

            // Overlay 定位用 DIP（物理像素 ÷ DPI缩放）
            _dropOverlay?.SetBounds(physicalRect, _dpiX, _dpiY);

            RefreshDebugDropZone(physicalRect);
        }
        catch { }
    }

    // ═══════════════════════════════════════════
    //  热键
    // ═══════════════════════════════════════════

    private void RegisterHotkeys()
    {
        _hotkeyGrab = _hotkeyService.Register(
            HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt,
            0x47, GrabActiveWindow);

        _hotkeyNext = _hotkeyService.Register(
            HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt,
            0x27, () => Dispatcher.BeginInvoke(_layoutService.SwitchToNext));

        _hotkeyPrev = _hotkeyService.Register(
            HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt,
            0x25, () => Dispatcher.BeginInvoke(_layoutService.SwitchToPrevious));

        _hotkeyQuad = _hotkeyService.Register(
            HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt,
            0x51, () => Dispatcher.BeginInvoke(() =>
            {
                _layoutService.CurrentMode = LayoutService.LayoutMode.QuadSplit;
                _layoutService.ApplyLayout();
                UpdateLayoutButtons();
            }));

        if (_hotkeyGrab == -1)
            SetStatus("⚠ 部分热键注册失败（可能与其他程序冲突）");
    }

    private void GrabActiveWindow()
    {
        Dispatcher.BeginInvoke(() =>
        {
            IntPtr fg = NativeMethods.GetForegroundWindow();
            if (fg == IntPtr.Zero || fg == _myHwnd) return;

            if (_windowManager.ManagedWindows.Any(w => w.Handle == fg))
            {
                SetStatus("该窗口已在管理中");
                return;
            }

            if (!WindowEnumerator.ShouldInclude(fg, excludeOwnerHwnd: _myHwnd))
            {
                SetStatus("⚠ 该窗口无法被管理（系统窗口或不可见）");
                return;
            }

            if (_windowManager.TryManageWindow(fg))
            {
                UpdateHostArea();
                _layoutService.ApplyLayout();
            }
        });
    }

    // ═══════════════════════════════════════════
    //  拖拽检测回调
    // ═══════════════════════════════════════════

    private void OnExternalDragStarted(IntPtr hwnd)
    {
        Dispatcher.BeginInvoke(() =>
        {
            if (hwnd == _myHwnd) return;
            if (_windowManager.ManagedWindows.Any(w => w.Handle == hwnd)) return;
            if (_sidebarExpanded)
                _dropOverlay?.ShowOverlay();
        });
    }

    private void OnDragMoved(IntPtr hwnd, Point mousePos)
    {
        Dispatcher.BeginInvoke(() =>
        {
            if (!_sidebarExpanded) return;
            bool inZone = !_currentDropZonePhys.IsEmpty && _currentDropZonePhys.Contains(mousePos);
            _dropOverlay?.UpdateHighlight(inZone);

            if (_debugVisible)
            {
                NativeMethods.GetCursorPos(out POINT pt);
                DbgMousePhysText.Text = $"鼠标(物理): ({pt.X}, {pt.Y})";
                DbgMouseInZoneText.Text = $"在区域内: {(inZone ? "✓ 是" : "✗ 否")}";
            }
        });
    }

    private void OnWindowDroppedInZone(IntPtr hwnd)
    {
        Dispatcher.BeginInvoke(() =>
        {
            if (_windowManager.TryManageWindow(hwnd))
            {
                UpdateHostArea();
                _layoutService.ApplyLayout();
                SetStatus($"已嵌入窗口: {WindowEnumerator.GetWindowTitle(hwnd)}");
            }
        });
    }

    private void OnDragEnded(IntPtr hwnd)
    {
        Dispatcher.BeginInvoke(() => _dropOverlay?.HideOverlay());
    }

    // ═══════════════════════════════════════════
    //  窗口管理回调
    // ═══════════════════════════════════════════

    private void OnManageFailed(IntPtr hwnd, string reason)
        => Dispatcher.BeginInvoke(() => SetStatus($"⚠ {reason}"));

    private void OnWindowManaged(ManagedWindow window)
        => Dispatcher.BeginInvoke(() => { SyncManagedWindowsToDetection(); SetStatus($"已嵌入: {window.Title}"); });

    private void OnWindowReleased(ManagedWindow window)
        => Dispatcher.BeginInvoke(SyncManagedWindowsToDetection);

    private void OnLayoutChanged()
        => Dispatcher.BeginInvoke(UpdateUI);

    // ═══════════════════════════════════════════
    //  UI 事件
    // ═══════════════════════════════════════════

    private void LayoutMode_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string modeStr &&
            Enum.TryParse<LayoutService.LayoutMode>(modeStr, out var mode))
        {
            _layoutService.CurrentMode = mode;
            _layoutService.ApplyLayout();
            UpdateLayoutButtons();
            SetStatus($"布局模式: {GetModeName(mode)}");
        }
    }

    private void RefreshLayout_Click(object sender, RoutedEventArgs e)
    {
        RefreshDpiScale();
        UpdateHostArea();
        UpdateDropZone();
        _layoutService.ApplyLayout();
        SetStatus("布局已刷新");
    }

    private void ReleaseAll_Click(object sender, RoutedEventArgs e)
    {
        _windowManager.ReleaseAll();
        SetStatus("已释放所有窗口");
    }

    private void WindowItem_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.DataContext is ManagedWindow window)
        {
            _layoutService.SwitchToWindow(window);
            SetStatus($"已切换到: {window.Title}");
        }
    }

    private void ReleaseButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is IntPtr handle)
        {
            var title = WindowEnumerator.GetWindowTitle(handle);
            _windowManager.ReleaseWindow(handle);
            SetStatus($"已释放: {title}");
        }
    }

    // ── 侧边栏折叠/展开 ──

    private void ToggleSidebar_Click(object sender, RoutedEventArgs e) => ToggleSidebar();

    private void ToggleSidebar()
    {
        var col = ContentGrid.ColumnDefinitions[0];
        var splitterCol = ContentGrid.ColumnDefinitions[1];

        if (_sidebarExpanded)
        {
            _savedSidebarWidth = col.ActualWidth > 30 ? col.ActualWidth : 200;
            col.Width = new GridLength(28);
            col.MinWidth = 28;
            col.MaxWidth = 28;
            splitterCol.Width = new GridLength(0);
            MainSplitter.IsEnabled = false;

            SidebarTitleText.Visibility = Visibility.Collapsed;
            SidebarContentPanel.Visibility = Visibility.Collapsed;
            BtnToggleSidebar.Content = "▶";
            BtnToggleSidebar.ToolTip = "展开侧边栏";

            _sidebarExpanded = false;
            _dragDetection.SetDropZone(Rect.Empty);
            _currentDropZonePhys = Rect.Empty;
        }
        else
        {
            col.MaxWidth = 360;
            col.MinWidth = 120;
            col.Width = new GridLength(_savedSidebarWidth);
            splitterCol.Width = new GridLength(4);
            MainSplitter.IsEnabled = true;

            SidebarTitleText.Visibility = Visibility.Visible;
            SidebarContentPanel.Visibility = Visibility.Visible;
            BtnToggleSidebar.Content = "◀";
            BtnToggleSidebar.ToolTip = "折叠侧边栏";

            _sidebarExpanded = true;

            Dispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(() =>
            {
                RefreshDpiScale();
                UpdateHostArea();
                UpdateDropZone();
                _layoutService.ApplyLayout();
            }));
        }
    }

    // ── Debug 面板 ──

    private void ToggleDebug_Click(object sender, RoutedEventArgs e)
    {
        _debugVisible = !_debugVisible;
        DebugPanel.Visibility = _debugVisible ? Visibility.Visible : Visibility.Collapsed;
        BtnDebug.BorderBrush = _debugVisible
            ? (Brush)FindResource("AccentBrush")
            : (Brush)FindResource("BorderBrush");

        if (_debugVisible)
        {
            // 立即刷新一次
            RefreshAllDebug();

            // 启动定时刷新（鼠标位置等实时数据）
            _debugTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
            _debugTimer.Tick += (_, _) => RefreshAllDebug();
            _debugTimer.Start();
        }
        else
        {
            _debugTimer?.Stop();
            _debugTimer = null;
        }
    }

    private void RefreshAllDebug()
    {
        if (!_debugVisible) return;

        // DPI
        DbgDpiText.Text = $"DPI倍率: X={_dpiX:F3}  Y={_dpiY:F3}  ({(int)(_dpiX * 96)} dpi)";

        // 鼠标（物理像素）
        NativeMethods.GetCursorPos(out POINT pt);
        DbgMousePhysText.Text = $"鼠标(物理px): ({pt.X}, {pt.Y})";

        // 在区域内？
        bool inZone = !_currentDropZonePhys.IsEmpty &&
                      _currentDropZonePhys.Contains(new Point(pt.X, pt.Y));
        DbgMouseInZoneText.Text = $"鼠标在DropZone内: {(inZone ? "✓ 是" : "✗ 否")}";

        RefreshDebugDropZone(_currentDropZonePhys);

        // Overlay 实际位置
        if (_dropOverlay != null)
        {
            DbgOverlayDipText.Text = $"Overlay L={_dropOverlay.Left:F1} T={_dropOverlay.Top:F1}";
            DbgSidebarPhysText.Text = $"       W={_dropOverlay.Width:F1} H={_dropOverlay.Height:F1}";
        }
    }

    private void RefreshDebugDropZone(Rect phys)
    {
        if (!_debugVisible) return;
        if (phys.IsEmpty)
        {
            DbgDropZonePhysText.Text = "DropZone: (空)";
        }
        else
        {
            DbgDropZonePhysText.Text =
                $"L={phys.Left:F0} T={phys.Top:F0}";
            // 显示 DIP 等价
            var dipL = phys.Left / _dpiX;
            var dipT = phys.Top / _dpiY;
            var dipW = phys.Width / _dpiX;
            var dipH = phys.Height / _dpiY;
            DbgMouseInZoneText.Text =
                $"W={phys.Width:F0} H={phys.Height:F0}  → DIP({dipL:F0},{dipT:F0} {dipW:F0}×{dipH:F0})";
        }
    }

    private void RefreshDebugHostArea(int x, int y, int w, int h)
    {
        if (!_debugVisible) return;
        DbgHostAreaText.Text = $"原点(客户区px): ({x}, {y})";
        DbgHostSizeText.Text = $"大小: {w} × {h} px";
    }

    // ═══════════════════════════════════════════
    //  辅助
    // ═══════════════════════════════════════════

    private void SyncManagedWindowsToDetection()
    {
        _dragDetection.ManagedWindows.Clear();
        foreach (var w in _windowManager.ManagedWindows)
            _dragDetection.ManagedWindows.Add(w.Handle);
    }

    private void UpdateUI()
    {
        int count = _windowManager.ManagedWindows.Count;
        bool hasWin = count > 0;
        WindowCountText.Text = $"{count} 个窗口";
        EmptyStatePanel.Visibility = hasWin ? Visibility.Collapsed : Visibility.Visible;
        WindowListBox.Visibility = hasWin ? Visibility.Visible : Visibility.Collapsed;
        DropHintPanel.Visibility = hasWin ? Visibility.Collapsed : Visibility.Visible;
    }

    private void ManagedWindows_Changed(object? sender, NotifyCollectionChangedEventArgs e)
        => Dispatcher.BeginInvoke(UpdateUI);

    private void SetStatus(string text) => StatusText.Text = text;

    private void UpdateLayoutButtons()
    {
        var mode = _layoutService.CurrentMode;
        var accent = (Brush)FindResource("AccentBrush");
        var border = (Brush)FindResource("BorderBrush");
        BtnStacked.BorderBrush = mode == LayoutService.LayoutMode.Stacked ? accent : border;
        BtnQuad.BorderBrush = mode == LayoutService.LayoutMode.QuadSplit ? accent : border;
        BtnLeftRight.BorderBrush = mode == LayoutService.LayoutMode.LeftRight ? accent : border;
        BtnTopBottom.BorderBrush = mode == LayoutService.LayoutMode.TopBottom ? accent : border;
    }

    private static string GetModeName(LayoutService.LayoutMode mode) => mode switch
    {
        LayoutService.LayoutMode.Stacked => "堆叠（抽屉切换）",
        LayoutService.LayoutMode.QuadSplit => "四分区",
        LayoutService.LayoutMode.LeftRight => "左右分屏",
        LayoutService.LayoutMode.TopBottom => "上下分屏",
        _ => mode.ToString()
    };
}