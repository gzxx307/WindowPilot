using System.Collections.Specialized;
using System.IO;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Web.WebView2.Core;
using WindowPilot.Bridge;
using WindowPilot.Controls;
using WindowPilot.Diagnostics;
using WindowPilot.Models;
using WindowPilot.Native;
using WindowPilot.Services;

namespace WindowPilot;

public partial class MainWindow : Window
{
    private const string Cat = "MainWindow";

    // ── 服务层 ────────────────────────────────────────────────────────
    private readonly WinEventHookService  _winEventHook  = new();
    private readonly WindowManagerService _windowManager;
    private readonly DragDetectionService _dragDetection;
    private readonly LayoutService        _layout;
    private readonly HotkeyService        _hotkey        = new();

    // ── UI 通信桥 ─────────────────────────────────────────────────────
    private readonly UIBridge _bridge;

    // ── 覆盖层（独立 WPF 窗口） ──────────────────────────────────────
    private DropZoneOverlay?         _overlay;
    private ReleaseZoneOverlay?      _releaseOverlay;

    // ── 缩略图预览 ────────────────────────────────────────────────────
    private ThumbnailPreviewService? _thumbnailPreview;
    private System.Windows.Threading.DispatcherTimer? _hoverTimer;
    private ManagedWindow? _hoverTarget;
    private Point          _hoverItemScreenTL;
    private double         _sidebarScreenRight;

    // ── 侧边栏折叠状态 ────────────────────────────────────────────────
    private bool   _sidebarExpanded   = true;
    private double _savedSidebarWidth = 200;

    // ── 构造函数 ──────────────────────────────────────────────────────

    public MainWindow()
    {
        Logger.Separator("MainWindow Init");
        Logger.Debug("MainWindow 构造中…", Cat);

        InitializeComponent();

        _windowManager = new WindowManagerService(_winEventHook);
        _dragDetection = new DragDetectionService(_winEventHook);
        _layout        = new LayoutService(_windowManager);

        _bridge = new UIBridge(_windowManager, _layout, Dispatcher);

        // UIBridge 发出的工具栏业务事件
        _bridge.LayoutModeRequested    += mode => Dispatcher.BeginInvoke(() => ApplyLayoutMode(mode));
        _bridge.RefreshLayoutRequested += () => Dispatcher.BeginInvoke(() =>
        {
            UpdateHostArea();
            _bridge.PushStatus("布局已重排");
        });
        _bridge.ReleaseAllRequested += () => Dispatcher.BeginInvoke(() =>
        {
            _windowManager.ReleaseAll();
            _bridge.PushStatus("已释放所有窗口");
        });
        _bridge.ToggleSidebarRequested += () => Dispatcher.BeginInvoke(ToggleSidebar);

        SubscribeServiceEvents();
    }

    // ── 服务事件订阅 ──────────────────────────────────────────────────

    private void SubscribeServiceEvents()
    {
        _windowManager.ManagedWindows.CollectionChanged += OnManagedWindowsChanged;

        _windowManager.WindowManaged += w => Dispatcher.BeginInvoke(() =>
        {
            Logger.Info($"WindowManaged 回调: \"{w.Title}\"", Cat);
            _bridge.PushStatus($"已接管：{w.Title}");
            _bridge.PushWindowList();
        });

        _windowManager.WindowReleased += w => Dispatcher.BeginInvoke(() =>
        {
            Logger.Info($"WindowReleased 回调: \"{w.Title}\"", Cat);
            _bridge.PushStatus($"已释放：{w.Title}");
            _bridge.PushWindowList();
        });

        _windowManager.LayoutChanged += () => Dispatcher.BeginInvoke(() =>
        {
            Logger.Trace("LayoutChanged 触发 → ApplyLayout", Cat);
            _layout.ApplyLayout();
            _bridge.PushWindowList();
        });

        _windowManager.ManageFailed += (_, msg) => Dispatcher.BeginInvoke(() =>
        {
            Logger.Warning($"ManageFailed: {msg}", Cat);
            _bridge.PushStatus(msg);
        });

        _dragDetection.ExternalDragStarted += _ => Dispatcher.BeginInvoke(() =>
        {
            Logger.Debug("ExternalDragStarted → ShowDropOverlay", Cat);
            ShowDropOverlay();
            _bridge.PushDropOverlay(true);
        });

        _dragDetection.DragEnded += _ => Dispatcher.BeginInvoke(() =>
        {
            Logger.Debug("DragEnded → HideDropOverlay", Cat);
            HideDropOverlay();
            _bridge.PushDropOverlay(false);
        });

        _dragDetection.DragMoved              += (_, _) => { };
        _dragDetection.WindowDroppedInZone    += OnWindowDroppedInZone;
        _dragDetection.WindowDraggedOutOfZone += OnWindowDraggedOutOfZone;

        _dragDetection.ManagedDragStarted            += OnManagedDragStarted;
        _dragDetection.ManagedDragMoved              += OnManagedDragMoved;
        _dragDetection.ManagedWindowDroppedOnSidebar += OnManagedWindowDroppedOnSidebar;
        _dragDetection.ManagedDragCancelled          += OnManagedDragCancelled;
        _dragDetection.ManagedDragEnded              += OnManagedDragEnded;

        Logger.Debug("所有事件订阅完毕。", Cat);
    }

    // ── 生命周期 ──────────────────────────────────────────────────────

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        Logger.Separator("MainWindow Loaded");

        _windowManager.HostHwnd = new WindowInteropHelper(this).Handle;
        Logger.Info($"HostHwnd = 0x{_windowManager.HostHwnd:X}", Cat);

        _winEventHook.Start();

        _hotkey.Attach(this);

        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x47,
            GrabForegroundWindow, "Ctrl+Alt+G 抓取前台窗口");

        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x25,
            () => Dispatcher.BeginInvoke(() =>
            {
                _layout.SwitchToPrevious();
                _bridge.PushWindowList();
            }), "Ctrl+Alt+← 切换上一个");

        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x27,
            () => Dispatcher.BeginInvoke(() =>
            {
                _layout.SwitchToNext();
                _bridge.PushWindowList();
            }), "Ctrl+Alt+→ 切换下一个");

        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x51,
            () => Dispatcher.BeginInvoke(() => ApplyLayoutMode(LayoutService.LayoutMode.QuadSplit)),
            "Ctrl+Alt+Q 四分区布局");

        _hotkey.Register(HotkeyService.Modifiers.Ctrl | HotkeyService.Modifiers.Alt, 0x4D,
            StartActiveWindowDrag, "Ctrl+Alt+M 编程式拖拽");

        _overlay        = new DropZoneOverlay();
        _releaseOverlay = new ReleaseZoneOverlay();
        Logger.Debug("覆盖层窗口已创建。", Cat);

        _thumbnailPreview = new ThumbnailPreviewService();
        _hoverTimer       = new System.Windows.Threading.DispatcherTimer
                            { Interval = TimeSpan.FromMilliseconds(350) };
        _hoverTimer.Tick += HoverTimer_Tick;
        Logger.Debug("缩略图预览服务已初始化。", Cat);

        var hwndSource = HwndSource.FromHwnd(_windowManager.HostHwnd);
        hwndSource?.AddHook(HostWndProc);
        Logger.Info("HostWndProc 钩子已安装。", Cat);

        // 共享同一个 CoreWebView2Environment，避免多进程开销
        await InitAllWebViewsAsync();

        Dispatcher.BeginInvoke(() => { UpdateHostArea(); UpdateDropZone(); });

        Logger.Info("MainWindow 已加载完毕。", Cat);
        Logger.Debug($"日志文件位置: {Logger.CurrentLogFilePath}", Cat);
    }

    private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        Logger.Separator("MainWindow Closing");

        _hotkey.Dispose();
        _hoverTimer?.Stop();
        _thumbnailPreview?.Dispose();

        _windowManager.ReleaseAll();
        _windowManager.Dispose();
        _dragDetection.Dispose();
        _winEventHook.Dispose();

        _overlay?.Close();
        _releaseOverlay?.Close();

        Logger.Info("MainWindow 清理完成。", Cat);
    }

    private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        Logger.Trace(
            $"MainWindow SizeChanged: {e.PreviousSize.Width:F0}×{e.PreviousSize.Height:F0} → " +
            $"{e.NewSize.Width:F0}×{e.NewSize.Height:F0}", Cat);
        Dispatcher.BeginInvoke(() => { UpdateHostArea(); UpdateDropZone(); });
    }

    private void HostPanel_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        Logger.Trace($"HostPanel SizeChanged: {e.NewSize.Width:F0}×{e.NewSize.Height:F0}", Cat);
        Dispatcher.BeginInvoke(UpdateHostArea);
    }

    private void SidebarBorder_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        Logger.Trace($"SidebarBorder SizeChanged: {e.NewSize.Width:F0}×{e.NewSize.Height:F0}", Cat);
        Dispatcher.BeginInvoke(UpdateDropZone);
    }

    // ── WebView2 初始化（共享同一 Environment） ───────────────────────

    private async Task InitAllWebViewsAsync()
    {
        Logger.Debug("InitAllWebViewsAsync 开始。", Cat);

        var cacheDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WindowPilot", "WebView2Cache");

        // 所有 3 个 WebView2 共用同一个 Environment，节省进程资源
        var env = await CoreWebView2Environment.CreateAsync(userDataFolder: cacheDir);

        var uiFolder = Path.Combine(AppContext.BaseDirectory, "UI");
        const string vhost = "windowpilot.local";

        // 初始化工具栏 WebView2
        await ToolbarWebView.EnsureCoreWebView2Async(env);
        ConfigureWebView(ToolbarWebView.CoreWebView2, vhost, uiFolder);
        _bridge.AttachToolbar(ToolbarWebView.CoreWebView2);
        ToolbarWebView.Source = new Uri($"https://{vhost}/toolbar.html");

        // 初始化侧边栏 WebView2
        await SidebarWebView.EnsureCoreWebView2Async(env);
        ConfigureWebView(SidebarWebView.CoreWebView2, vhost, uiFolder);
        _bridge.AttachSidebar(SidebarWebView.CoreWebView2);
        SidebarWebView.Source = new Uri($"https://{vhost}/sidebar.html");

        // 初始化状态栏 WebView2
        await StatusWebView.EnsureCoreWebView2Async(env);
        ConfigureWebView(StatusWebView.CoreWebView2, vhost, uiFolder);
        _bridge.AttachStatus(StatusWebView.CoreWebView2);
        StatusWebView.Source = new Uri($"https://{vhost}/statusbar.html");

        Logger.Info("3 个 WebView2 初始化完成。", Cat);
    }

    private static void ConfigureWebView(CoreWebView2 core, string vhost, string folder)
    {
        core.SetVirtualHostNameToFolderMapping(vhost, folder, CoreWebView2HostResourceAccessKind.Allow);
        core.Settings.AreDevToolsEnabled            = true;  // 发布时改为 false
        core.Settings.AreDefaultContextMenusEnabled = false;
        core.Settings.IsStatusBarEnabled            = false;
        core.Settings.IsZoomControlEnabled          = false;
    }

    // ── 侧边栏折叠/展开 ───────────────────────────────────────────────

    private void ToggleSidebar()
    {
        Logger.Debug($"ToggleSidebar: expanded={_sidebarExpanded} → {!_sidebarExpanded}", Cat);
        if (_sidebarExpanded)
        {
            _savedSidebarWidth  = SidebarColumn.ActualWidth;
            SidebarColumn.Width    = new GridLength(28);
            SidebarColumn.MinWidth = 0;
            _sidebarExpanded    = false;
        }
        else
        {
            SidebarColumn.Width    = new GridLength(_savedSidebarWidth);
            SidebarColumn.MinWidth = 28;
            _sidebarExpanded    = true;
        }
        _bridge.SetSidebarExpanded(_sidebarExpanded);
        _bridge.PushSidebarCollapsed(!_sidebarExpanded);
        Dispatcher.BeginInvoke(() => { UpdateHostArea(); UpdateDropZone(); });
    }

    // ── 布局模式切换 ──────────────────────────────────────────────────

    private void ApplyLayoutMode(LayoutService.LayoutMode mode)
    {
        Logger.Info($"ApplyLayoutMode: {mode}", Cat);
        _layout.CurrentMode = mode;
        _layout.ApplyLayout();
        _bridge.PushLayoutMode(mode);
        _bridge.PushStatus($"布局模式：{GetLayoutName(mode)}");
    }

    private static string GetLayoutName(LayoutService.LayoutMode mode) => mode switch
    {
        LayoutService.LayoutMode.Stacked   => "堆叠抽屉",
        LayoutService.LayoutMode.QuadSplit => "四分区",
        LayoutService.LayoutMode.LeftRight => "左右分屏",
        LayoutService.LayoutMode.TopBottom => "上下分屏",
        _                                  => mode.ToString()
    };

    // ── 宿主区域与放置区域（与原版完全一致） ─────────────────────────────

    private void UpdateHostArea()
    {
        if (!IsLoaded || HostPanel.ActualWidth <= 0) return;

        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero) return;

        var screenTL = HostPanel.PointToScreen(new Point(0, 0));
        var screenBR = HostPanel.PointToScreen(
            new Point(HostPanel.ActualWidth, HostPanel.ActualHeight));

        var clientTL = new POINT { X = (int)screenTL.X, Y = (int)screenTL.Y };
        NativeMethods.ScreenToClient(hwnd, ref clientTL);

        int w = (int)(screenBR.X - screenTL.X);
        int h = (int)(screenBR.Y - screenTL.Y);

        Logger.Debug($"UpdateHostArea: client=({clientTL.X},{clientTL.Y}) {w}×{h}px", Cat);

        _layout.SetHostArea(clientTL.X, clientTL.Y, w, h);
        _layout.ApplyLayout();
    }

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

    private void SyncManagedWindowsToDetection()
    {
        int before = _dragDetection.ManagedWindows.Count;
        _dragDetection.ManagedWindows.Clear();
        foreach (var w in _windowManager.ManagedWindows)
            _dragDetection.ManagedWindows.Add(w.Handle);

        Logger.Trace(
            $"SyncManagedWindowsToDetection: {before} → {_dragDetection.ManagedWindows.Count} 个句柄", Cat);
    }

    // ── 覆盖层 ────────────────────────────────────────────────────────

    private void ShowDropOverlay()
    {
        if (_overlay == null || !IsLoaded || SidebarBorder.ActualWidth <= 0) return;
        var tl   = SidebarBorder.PointToScreen(new Point(0, 0));
        var rect = new Rect(tl, new Size(SidebarBorder.ActualWidth, SidebarBorder.ActualHeight));
        _overlay.ShowAtRect(rect);
    }

    private void HideDropOverlay()   => _overlay?.HideOverlay();
    private void ShowReleaseOverlay()
    {
        if (_releaseOverlay == null || !IsLoaded || SidebarBorder.ActualWidth <= 0) return;
        var tl   = SidebarBorder.PointToScreen(new Point(0, 0));
        var rect = new Rect(tl, new Size(SidebarBorder.ActualWidth, SidebarBorder.ActualHeight));
        _releaseOverlay.ShowAtRect(rect);
    }
    private void HideReleaseOverlay() => _releaseOverlay?.HideOverlay();

    // ── 外部拖拽回调（与原版完全一致） ───────────────────────────────────

    private void OnWindowDroppedInZone(IntPtr hwnd)
    {
        Logger.Info($"OnWindowDroppedInZone  hwnd=0x{hwnd:X}", Cat);

        bool hasPre = _dragDetection.PreDragRects.TryGetValue(hwnd, out RECT preDragRect);
        Logger.Debug(hasPre
            ? $"  PreDragRect 捕获成功: ({preDragRect.Left},{preDragRect.Top})"
            : "  PreDragRect 未找到", Cat);

        Dispatcher.BeginInvoke(() =>
        {
            _windowManager.TryManageWindow(hwnd);
            _dragDetection.ManagedWindows.Add(hwnd);

            if (hasPre)
            {
                var window = _windowManager.FindByHandle(hwnd);
                if (window != null) window.OriginalRect = preDragRect;
            }
        });
    }

    private void OnWindowDraggedOutOfZone(IntPtr hwnd)
    {
        Logger.Info($"OnWindowDraggedOutOfZone  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            _windowManager.ReleaseWindow(hwnd);
            _dragDetection.ManagedWindows.Remove(hwnd);
        });
    }

    // ── 托管窗口拖拽回调（与原版完全一致） ───────────────────────────────

    private void OnManagedDragStarted(IntPtr hwnd)
    {
        Logger.Debug($"OnManagedDragStarted  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            ShowReleaseOverlay();
            _bridge.PushStatus("拖拽中 — 移至侧边栏松开可解除托管");
        });
    }

    private void OnManagedDragMoved(IntPtr hwnd, Point screenPt) { }

    private void OnManagedWindowDroppedOnSidebar(IntPtr hwnd)
    {
        Logger.Info($"OnManagedWindowDroppedOnSidebar  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            _windowManager.ReleaseWindow(hwnd);
            _dragDetection.ManagedWindows.Remove(hwnd);
            var mw = _windowManager.FindByHandle(hwnd);
            _bridge.PushStatus($"已释放：{mw?.Title ?? "窗口"}");
        });
    }

    private void OnManagedDragCancelled(IntPtr hwnd)
    {
        Logger.Debug($"OnManagedDragCancelled  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            if (_layout.CurrentMode == LayoutService.LayoutMode.Stacked)
            {
                ReEmbedAndRelayout(hwnd);
                return;
            }

            NativeMethods.GetCursorPos(out POINT cursorPt);
            var cursorPoint = new Point(cursorPt.X, cursorPt.Y);
            var hostHwnd    = new WindowInteropHelper(this).Handle;
            int targetSlot  = _layout.GetSlotIndexAtScreenPoint(cursorPoint, hostHwnd);

            var draggedWindow = _windowManager.FindByHandle(hwnd);
            if (draggedWindow == null) return;

            int originalSlot = draggedWindow.SlotIndex;

            if (targetSlot >= 0 && targetSlot != originalSlot)
            {
                var targetWindow = _layout.GetWindowAtSlotIndex(targetSlot);
                if (targetWindow != null && targetWindow.Handle != hwnd)
                {
                    if (!draggedWindow.IsEmbedded) _windowManager.ReEmbed(hwnd);
                    _windowManager.SwapOrder(hwnd, targetWindow.Handle);
                    _layout.ApplyLayout();
                    SyncManagedWindowsToDetection();
                    _bridge.PushStatus($"已互换位置：{draggedWindow.Title}  ⇄  {targetWindow.Title}");
                    _bridge.PushWindowList();
                    return;
                }
            }

            ReEmbedAndRelayout(hwnd);
            _bridge.PushStatus("拖拽已取消，窗口回到原位");
        });
    }

    private void OnManagedDragEnded(IntPtr hwnd)
    {
        Logger.Debug($"OnManagedDragEnded  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(HideReleaseOverlay);
    }

    private void ReEmbedAndRelayout(IntPtr hwnd)
    {
        var window = _windowManager.FindByHandle(hwnd);
        if (window != null && !window.IsEmbedded) _windowManager.ReEmbed(hwnd);
        _layout.ApplyLayout();
        SyncManagedWindowsToDetection();
    }

    // ── 集合变化 ──────────────────────────────────────────────────────

    private void OnManagedWindowsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        Logger.Trace(
            $"ManagedWindows CollectionChanged: action={e.Action}  " +
            $"count={_windowManager.ManagedWindows.Count}", Cat);

        Dispatcher.BeginInvoke(() =>
        {
            bool has = _windowManager.ManagedWindows.Count > 0;
            DropHintPanel.Visibility = has ? Visibility.Collapsed : Visibility.Visible;
            SyncManagedWindowsToDetection();
            _bridge.PushWindowList();
        });
    }

    // ── 热键动作（与原版完全一致） ────────────────────────────────────

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

    private void StartActiveWindowDrag()
    {
        Logger.Debug("StartActiveWindowDrag 热键触发", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            if (_layout.CurrentMode == LayoutService.LayoutMode.Stacked &&
                _windowManager.ManagedWindows.Count <= 1)
            {
                _bridge.PushStatus("堆叠模式下仅有一个窗口，无需拖拽");
                return;
            }

            var activeWindow = _windowManager.ManagedWindows.FirstOrDefault(w => w.IsActive)
                             ?? _windowManager.ManagedWindows.FirstOrDefault();

            if (activeWindow == null)
            {
                _bridge.PushStatus("没有可拖拽的托管窗口");
                return;
            }

            int originalSlot = activeWindow.SlotIndex;

            if (_windowManager.StartProgrammaticDrag(activeWindow.Handle))
            {
                _dragDetection.NotifyProgrammaticDragStarted(activeWindow.Handle, originalSlot);
                _bridge.PushStatus($"拖拽 \"{activeWindow.Title}\" — 移动到目标位置后松开鼠标");
            }
            else
            {
                Logger.Error($"StartProgrammaticDrag 失败: \"{activeWindow.Title}\"", Cat);
                _bridge.PushStatus($"无法发起拖拽：\"{activeWindow.Title}\"");
            }
        });
    }

    // ── 缩略图预览（与原版一致） ──────────────────────────────────────

    private void HoverTimer_Tick(object? sender, EventArgs e)
    {
        _hoverTimer?.Stop();
        if (_hoverTarget == null || _thumbnailPreview == null) return;

        var source      = PresentationSource.FromVisual(this);
        double dpiScale = source?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;

        Logger.Debug(
            $"HoverTimer 触发 → 显示预览: \"{_hoverTarget.Title}\"  dpiScale={dpiScale:F2}", Cat);

        _thumbnailPreview.Show(_hoverTarget, _hoverItemScreenTL, _sidebarScreenRight, dpiScale);
    }

    // ── WndProc 钩子（与原版完全一致） ───────────────────────────────

    private IntPtr HostWndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == NativeConstants.WM_PARENTNOTIFY)
        {
            int childMsg = wParam.ToInt32() & 0xFFFF;
            if (childMsg == NativeConstants.WM_LBUTTONDOWN)
            {
                int clientX = (short)(lParam.ToInt64() & 0xFFFF);
                int clientY = (short)((lParam.ToInt64() >> 16) & 0xFFFF);

                var target = FindManagedWindowAtClientPoint(clientX, clientY);
                if (target != null)
                {
                    Logger.Debug(
                        $"WM_PARENTNOTIFY click → FocusEmbedded: \"{target.Title}\"  " +
                        $"client=({clientX},{clientY})", Cat);
                    _windowManager.FocusEmbeddedWindow(target.Handle);

                    foreach (var w in _windowManager.ManagedWindows)
                        w.IsActive = w.Handle == target.Handle;

                    _bridge.PushWindowList();
                }
            }
        }
        return IntPtr.Zero;
    }

    private ManagedWindow? FindManagedWindowAtClientPoint(int clientX, int clientY)
    {
        var pt = new POINT { X = clientX, Y = clientY };
        NativeMethods.ClientToScreen(_windowManager.HostHwnd, ref pt);
        IntPtr clickedHwnd = NativeMethods.WindowFromPoint(pt);

        if (clickedHwnd == IntPtr.Zero) return null;

        IntPtr current = clickedHwnd;
        while (current != IntPtr.Zero && current != _windowManager.HostHwnd)
        {
            var managed = _windowManager.FindByHandle(current);
            if (managed != null) return managed;
            current = NativeMethods.GetParent(current);
        }
        return null;
    }
}