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

    // ── 服务 ──
    private readonly WinEventHookService   _winEventHook  = new();
    private readonly WindowManagerService  _windowManager;
    private readonly DragDetectionService  _dragDetection;
    private readonly LayoutService         _layout;
    private readonly HotkeyService         _hotkey        = new();

    // ── 覆盖层（拖拽外部窗口时显示在侧边栏上方） ──
    private DropZoneOverlay?       _overlay;

    // ── 覆盖层（拖拽已托管窗口时显示在侧边栏上） ──
    private ReleaseZoneOverlay? _releaseOverlay;

    // ── 侧边栏状态 ──
    private bool   _sidebarExpanded   = true;
    private double _savedSidebarWidth = 200;

    // ── 侧边栏条目拖拽状态 ──
    private Point _dragItemStartPoint;
    private bool  _isItemDragInProgress;

    // ────────────────────────── 构造函数 ──────────────────────────

    public MainWindow()
    {
        Logger.Separator("MainWindow Init");
        Logger.Debug("MainWindow 构造中…", Cat);

        InitializeComponent();

        _windowManager = new WindowManagerService(_winEventHook);
        _dragDetection = new DragDetectionService(_winEventHook);
        _layout        = new LayoutService(_windowManager);

        // 绑定窗口列表
        WindowListBox.ItemsSource = _windowManager.ManagedWindows;
        _windowManager.ManagedWindows.CollectionChanged += OnManagedWindowsChanged;

        // 管理器事件
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

        // ── 外部窗口拖入侧边栏 ──
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

        // ── 已托管窗口拖拽（互换 / 释放） ──
        _dragDetection.ManagedDragStarted            += OnManagedDragStarted;
        _dragDetection.ManagedDragMoved              += OnManagedDragMoved;
        _dragDetection.ManagedWindowDroppedOnSidebar += OnManagedWindowDroppedOnSidebar;
        _dragDetection.ManagedDragCancelled          += OnManagedDragCancelled;
        _dragDetection.ManagedDragEnded              += OnManagedDragEnded;

        Logger.Debug("所有事件订阅完毕。", Cat);
    }

    // ────────────────────────── 生命周期 ──────────────────────────

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        Logger.Separator("MainWindow Loaded");

        _windowManager.HostHwnd = new WindowInteropHelper(this).Handle;
        Logger.Info($"HostHwnd = 0x{_windowManager.HostHwnd:X}", Cat);

        _winEventHook.Start();

        // 注册全局热键
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

        _overlay        = new DropZoneOverlay();
        _releaseOverlay = new ReleaseZoneOverlay();
        Logger.Debug("覆盖层窗口已创建: DropZone / ReleaseZone", Cat);

        Dispatcher.BeginInvoke(() => { UpdateHostArea(); UpdateDropZone(); });

        Logger.Info("MainWindow 已加载完毕，程序就绪。", Cat);
        Logger.Debug($"日志文件位置: {Logger.CurrentLogFilePath}", Cat);

        SetStatus("就绪 — 拖拽窗口到侧边栏以接管，或按 Ctrl+Alt+G 抓取当前窗口 │ Ctrl+Alt+M 拖拽移动窗口");
    }

    private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        Logger.Separator("MainWindow Closing");
        Logger.Info("主窗口正在关闭，开始清理…", Cat);

        _hotkey.Dispose();
        Logger.Debug("HotkeyService 已释放。", Cat);

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

    // ────────────────────────── 宿主区域 / 放置区域 ──────────────────────────

    private void UpdateHostArea()
    {
        if (!IsLoaded || HostPanel.ActualWidth <= 0) return;

        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero) return;

        var screenTL = HostPanel.PointToScreen(new Point(0, 0));
        var screenBR = HostPanel.PointToScreen(new Point(HostPanel.ActualWidth, HostPanel.ActualHeight));

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

    // ────────────────────────── 覆盖层（外部窗口拖入） ──────────────────────────

    private void ShowDropOverlay()
    {
        if (_overlay == null || !IsLoaded || SidebarBorder.ActualWidth <= 0) return;
        var tl   = SidebarBorder.PointToScreen(new Point(0, 0));
        var rect = new Rect(tl, new Size(SidebarBorder.ActualWidth, SidebarBorder.ActualHeight));
        Logger.Trace($"ShowDropOverlay at ({rect.Left:F0},{rect.Top:F0})", Cat);
        _overlay.ShowAtRect(rect);
    }

    private void HideDropOverlay()
    {
        Logger.Trace("HideDropOverlay", Cat);
        _overlay?.HideOverlay();
    }

    // ────────────────────────── 覆盖层（已托管窗口拖拽） ──────────────────────────

    private void ShowReleaseOverlay()
    {
        if (_releaseOverlay == null || !IsLoaded || SidebarBorder.ActualWidth <= 0) return;
        var tl   = SidebarBorder.PointToScreen(new Point(0, 0));
        var rect = new Rect(tl, new Size(SidebarBorder.ActualWidth, SidebarBorder.ActualHeight));
        Logger.Trace($"ShowReleaseOverlay at ({rect.Left:F0},{rect.Top:F0})", Cat);
        _releaseOverlay.ShowAtRect(rect);
    }

    private void HideReleaseOverlay()
    {
        Logger.Trace("HideReleaseOverlay", Cat);
        _releaseOverlay?.HideOverlay();
    }

    // ────────────────────────── 外部拖拽检测回调 ──────────────────────────

    private void OnWindowDroppedInZone(IntPtr hwnd)
    {
        Logger.Info($"OnWindowDroppedInZone  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            _windowManager.TryManageWindow(hwnd);
            _dragDetection.ManagedWindows.Add(hwnd);
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

    // ════════════════════════════════════════════════════════════
    // 已托管窗口拖拽回调（互换位置 / 拖至侧边栏释放）
    // ════════════════════════════════════════════════════════════

    private void OnManagedDragStarted(IntPtr hwnd)
    {
        Logger.Debug($"OnManagedDragStarted  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            ShowReleaseOverlay();
            SetStatus("拖拽中 — 移至侧边栏松开可解除托管");
        });
    }

    private void OnManagedDragMoved(IntPtr hwnd, Point screenPt)
    {
        // 拖拽位置由 DragDetectionService 在 Trace 级别持续记录，此处无需额外处理
    }

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

    private void OnManagedDragCancelled(IntPtr hwnd)
    {
        Logger.Debug($"OnManagedDragCancelled  hwnd=0x{hwnd:X}", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            if (_layout.CurrentMode == LayoutService.LayoutMode.Stacked)
            {
                Logger.Debug("  堆叠模式下取消拖拽 → ReEmbedAndRelayout", Cat);
                ReEmbedAndRelayout(hwnd);
                return;
            }

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

            if (targetSlot >= 0 && targetSlot != originalSlot)
            {
                var targetWindow = _layout.GetWindowAtSlotIndex(targetSlot);
                if (targetWindow != null && targetWindow.Handle != hwnd)
                {
                    Logger.Info(
                        $"  互换: \"{draggedWindow.Title}\"[{originalSlot}] ⇄ " +
                        $"\"{targetWindow.Title}\"[{targetSlot}]", Cat);

                    if (!draggedWindow.IsEmbedded)
                        _windowManager.ReEmbed(hwnd);

                    _windowManager.SwapOrder(hwnd, targetWindow.Handle);
                    _layout.ApplyLayout();
                    SyncManagedWindowsToDetection();

                    SetStatus($"已互换位置：{draggedWindow.Title}  ⇄  {targetWindow.Title}");
                    return;
                }
            }

            Logger.Debug("  未命中有效目标 → ReEmbedAndRelayout（回到原位）", Cat);
            ReEmbedAndRelayout(hwnd);
            SetStatus("拖拽已取消，窗口回到原位");
        });
    }

    private void OnManagedDragEnded(IntPtr hwnd)
    {
        Logger.Debug($"OnManagedDragEnded  hwnd=0x{hwnd:X}  → 隐藏释放覆盖层", Cat);
        Dispatcher.BeginInvoke(HideReleaseOverlay);
    }

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

    // ════════════════════════════════════════════════════════════
    // 编程式发起活跃窗口拖拽（Ctrl+Alt+M）
    // ════════════════════════════════════════════════════════════

    private void StartActiveWindowDrag()
    {
        Logger.Debug("StartActiveWindowDrag 热键触发", Cat);
        Dispatcher.BeginInvoke(() =>
        {
            if (_layout.CurrentMode == LayoutService.LayoutMode.Stacked &&
                _windowManager.ManagedWindows.Count <= 1)
            {
                Logger.Warning("堆叠模式仅有一个窗口，无需拖拽。", Cat);
                SetStatus("堆叠模式下仅有一个窗口，无需拖拽");
                return;
            }

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

    // ════════════════════════════════════════════════════════════
    // 侧边栏条目拖拽：互换槽位 / 拖至释放区解除托管
    // ════════════════════════════════════════════════════════════

    private void WindowItemBorder_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        => _dragItemStartPoint = e.GetPosition(null);

    private void WindowItemBorder_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _isItemDragInProgress) return;

        var diff = e.GetPosition(null) - _dragItemStartPoint;
        if (Math.Abs(diff.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(diff.Y) < SystemParameters.MinimumVerticalDragDistance) return;

        if (sender is not Border border || border.DataContext is not ManagedWindow mw) return;

        Logger.Debug($"侧边栏条目拖拽开始: \"{mw.Title}\"  hwnd=0x{mw.Handle:X}", Cat);
        _isItemDragInProgress = true;

        SidebarReleaseZone.Visibility = Visibility.Visible;

        DragDrop.DoDragDrop(border,
            new DataObject(typeof(ManagedWindow), mw),
            DragDropEffects.Move);

        SidebarReleaseZone.Visibility = Visibility.Collapsed;
        _isItemDragInProgress = false;
        Logger.Trace("侧边栏条目拖拽结束（DoDragDrop 返回）。", Cat);
    }

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
        bool canSwap = dragged != null && dragged != target;

        if (canSwap)
        {
            e.Effects         = DragDropEffects.Move;
            border.Background = (Brush)FindResource("BgHover");
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private void WindowItemBorder_DragLeave(object sender, DragEventArgs e)
    {
        if (sender is Border border)
            border.ClearValue(Border.BackgroundProperty);
    }

    private void WindowItemBorder_Drop(object sender, DragEventArgs e)
    {
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

        string modeNote = _layout.CurrentMode == LayoutService.LayoutMode.Stacked
            ? "（堆叠模式：切换顺序已调整）" : string.Empty;
        SetStatus($"已交换位置：{dragged.Title}  ⇄  {target.Title}  {modeNote}".TrimEnd());

        e.Handled = true;
    }

    private void SidebarReleaseZone_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(typeof(ManagedWindow))
            ? DragDropEffects.Move
            : DragDropEffects.None;
        e.Handled = true;
    }

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

    // ────────────────────────── 工具栏按钮 ──────────────────────────

    private void LayoutMode_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn &&
            Enum.TryParse<LayoutService.LayoutMode>(btn.Tag?.ToString(), out var mode))
        {
            Logger.Debug($"LayoutMode 按钮点击: {mode}", Cat);
            SetLayoutMode(mode);
        }
    }

    private void SetLayoutMode(LayoutService.LayoutMode mode)
    {
        Logger.Info($"SetLayoutMode: {mode}", Cat);
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
        Logger.Debug("RefreshLayout 按钮点击", Cat);
        UpdateHostArea();
        SetStatus("布局已重排");
    }

    private void ReleaseAll_Click(object sender, RoutedEventArgs e)
    {
        Logger.Info("ReleaseAll 按钮点击", Cat);
        _windowManager.ReleaseAll();
        SetStatus("已释放所有窗口");
    }

    private void ToggleSidebar_Click(object sender, RoutedEventArgs e)
    {
        Logger.Debug($"ToggleSidebar: expanded={_sidebarExpanded} → {!_sidebarExpanded}", Cat);
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

    private void WindowItem_Click(object sender, MouseButtonEventArgs e)
    {
        if (_isItemDragInProgress) return;
        if (sender is FrameworkElement fe && fe.DataContext is ManagedWindow mw)
        {
            Logger.Debug($"WindowItem 点击切换: \"{mw.Title}\"  hwnd=0x{mw.Handle:X}", Cat);
            _layout.SwitchToWindow(mw);
        }
    }

    private void ReleaseButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is IntPtr hwnd)
        {
            Logger.Debug($"ReleaseButton 点击: hwnd=0x{hwnd:X}", Cat);
            _windowManager.ReleaseWindow(hwnd);
        }
        e.Handled = true;
    }

    // ────────────────────────── 热键动作 ──────────────────────────

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

    // ────────────────────────── UI 辅助 ──────────────────────────

    private void OnManagedWindowsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        Logger.Trace(
            $"ManagedWindows CollectionChanged: action={e.Action}  " +
            $"count={_windowManager.ManagedWindows.Count}", Cat);

        Dispatcher.BeginInvoke(() =>
        {
            bool has = _windowManager.ManagedWindows.Count > 0;
            WindowListBox.Visibility   = has ? Visibility.Visible   : Visibility.Collapsed;
            EmptyStatePanel.Visibility = has ? Visibility.Collapsed : Visibility.Visible;
            DropHintPanel.Visibility   = has ? Visibility.Collapsed : Visibility.Visible;
            WindowCountText.Text = has
                ? $"{_windowManager.ManagedWindows.Count} 个窗口"
                : "0 个窗口";

            SyncManagedWindowsToDetection();
        });
    }

    private void SetStatus(string msg)
    {
        Logger.Trace($"SetStatus: {msg}", Cat);
        StatusText.Text = msg;
    }
}