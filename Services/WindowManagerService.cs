using System.Collections.ObjectModel;
using System.Windows;
using WindowPilot.Diagnostics;
using WindowPilot.Models;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 窗口管理核心：通过 SetParent 将外部窗口嵌入到宿主窗口内部
/// </summary>
public class WindowManagerService : IDisposable
{
    private const string Cat = "WindowManager";

    private readonly WinEventHookService _winEventHook;

    /// <summary>宿主窗口句柄（我们的主窗口）</summary>
    public IntPtr HostHwnd { get; set; }

    /// <summary>当前被管理的窗口列表</summary>
    public ObservableCollection<ManagedWindow> ManagedWindows { get; } = new();

    public event Action<IntPtr, string>? ManageFailed;
    public event Action<ManagedWindow>? WindowManaged;
    public event Action<ManagedWindow>? WindowReleased;
    public event Action? LayoutChanged;

    // ── 统计 ──
    private int _embedCount;
    private int _releaseCount;
    private int _repositionCount;

    public WindowManagerService(WinEventHookService winEventHook)
    {
        Logger.Debug("WindowManagerService 构造中…", Cat);
        _winEventHook = winEventHook;
        _winEventHook.WindowDestroyed    += OnWindowDestroyed;
        _winEventHook.WindowTitleChanged += OnWindowTitleChanged;
        Logger.Debug("已订阅 WindowDestroyed / WindowTitleChanged 事件。", Cat);
    }

    // ── 查询 ──────────────────────────────────────────────

    public ManagedWindow? FindByHandle(IntPtr hwnd)
        => ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);

    public bool IsManaged(IntPtr hwnd)
        => ManagedWindows.Any(w => w.Handle == hwnd);

    // ── 嵌入 ──────────────────────────────────────────────

    /// <summary>
    /// 尝试嵌入一个窗口到宿主内部
    /// </summary>
    public bool TryManageWindow(IntPtr hwnd)
    {
        Logger.Debug($"TryManageWindow  hwnd=0x{hwnd:X}", Cat);

        if (HostHwnd == IntPtr.Zero)
        {
            Logger.Error("TryManageWindow 失败：宿主窗口未就绪（HostHwnd == Zero）。", Cat);
            ManageFailed?.Invoke(hwnd, "宿主窗口未就绪");
            return false;
        }

        if (ManagedWindows.Any(w => w.Handle == hwnd))
        {
            Logger.Warning($"TryManageWindow: 0x{hwnd:X} 已在管理列表中，跳过。", Cat);
            return true;
        }

        if (hwnd == HostHwnd)
        {
            Logger.Warning("TryManageWindow: 尝试管理宿主窗口自身，已拒绝。", Cat);
            return false;
        }

        var window = new ManagedWindow(hwnd);
        Logger.Debug($"  目标窗口: \"{window.Title}\"  进程: {window.ProcessName}  PID: {window.ProcessId}", Cat);

        if (!window.CanWeControl())
        {
            Logger.Error(
                $"TryManageWindow 失败: \"{window.Title}\" 不可控制（可能需要更高权限）。", Cat);
            ManageFailed?.Invoke(hwnd, $"无法控制 \"{window.Title}\"：可能以管理员权限运行。");
            return false;
        }

        window.SaveOriginalState();
        Logger.Trace(
            $"  原始状态已保存: parent=0x{window.OriginalParent:X}  " +
            $"style=0x{window.OriginalStyle:X}  exStyle=0x{window.OriginalExStyle:X}  " +
            $"rect=({window.OriginalRect.Left},{window.OriginalRect.Top}," +
            $"{window.OriginalRect.Width}×{window.OriginalRect.Height})  " +
            $"wasMaximized={window.WasMaximized}", Cat);

        if (NativeMethods.IsZoomed(hwnd))
        {
            Logger.Debug($"  窗口处于最大化状态，先还原。", Cat);
            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_RESTORE);
        }
        if (NativeMethods.IsIconic(hwnd))
        {
            Logger.Debug($"  窗口处于最小化状态，先还原。", Cat);
            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_RESTORE);
        }

        if (!EmbedWindow(window))
        {
            Logger.Error(
                $"TryManageWindow 失败: 嵌入 \"{window.Title}\" (0x{hwnd:X}) 失败。", Cat);
            ManageFailed?.Invoke(hwnd, $"嵌入 \"{window.Title}\" 失败：该窗口可能不支持嵌入。");
            return false;
        }

        int count = Interlocked.Increment(ref _embedCount);
        window.IsManaged  = true;
        window.IsEmbedded = true;
        ManagedWindows.Add(window);
        WindowManaged?.Invoke(window);
        LayoutChanged?.Invoke();

        Logger.Info(
            $"[#{count}] 窗口接管成功: \"{window.Title}\"  hwnd=0x{hwnd:X}  " +
            $"当前管理数量: {ManagedWindows.Count}", Cat);
        return true;
    }

    private bool EmbedWindow(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;
        Logger.Trace($"EmbedWindow hwnd=0x{hwnd:X} → 修改样式并 SetParent", Cat);
        try
        {
            long style = window.OriginalStyle;
            style &= ~(long)(
                NativeConstants.WS_CAPTION     |
                NativeConstants.WS_THICKFRAME  |
                NativeConstants.WS_SYSMENU     |
                NativeConstants.WS_MINIMIZEBOX |
                NativeConstants.WS_MAXIMIZEBOX |
                NativeConstants.WS_POPUP       |
                NativeConstants.WS_BORDER      |
                NativeConstants.WS_DLGFRAME);
            style |= (long)NativeConstants.WS_CHILD;
            style |= (long)NativeConstants.WS_VISIBLE;
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE, style);
            Logger.Trace($"  新 Style=0x{style:X}", Cat);

            long exStyle = window.OriginalExStyle;
            exStyle &= ~(long)(NativeConstants.WS_EX_APPWINDOW | NativeConstants.WS_EX_TOOLWINDOW);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_EXSTYLE, exStyle);
            Logger.Trace($"  新 ExStyle=0x{exStyle:X}", Cat);

            var result = NativeMethods.SetParent(hwnd, HostHwnd);
            if (result == IntPtr.Zero)
            {
                Logger.Error($"  SetParent 返回 Zero（失败）。", Cat);
                return false;
            }
            Logger.Trace($"  SetParent 成功，旧父窗口=0x{result:X}", Cat);

            NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeConstants.SWP_NOMOVE | NativeConstants.SWP_NOSIZE |
                NativeConstants.SWP_NOZORDER | NativeConstants.SWP_FRAMECHANGED);

            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_SHOW);
            Logger.Trace($"  EmbedWindow 完成。", Cat);
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"EmbedWindow 异常", ex, Cat);
            return false;
        }
    }

    // ── 临时脱嵌 / 重新嵌入（用于拖拽互换和释放） ─────────

    /// <summary>
    /// 临时将窗口从宿主中脱离，使其成为顶层窗口以便自由拖拽。
    /// </summary>
    public bool TemporaryUnembed(IntPtr hwnd)
    {
        var window = FindByHandle(hwnd);
        if (window == null || !window.IsEmbedded)
        {
            Logger.Warning(
                $"TemporaryUnembed: hwnd=0x{hwnd:X} 未找到或未嵌入，跳过。", Cat);
            return false;
        }

        Logger.Debug($"TemporaryUnembed: \"{window.Title}\"  hwnd=0x{hwnd:X}", Cat);
        try
        {
            NativeMethods.GetWindowRect(hwnd, out RECT currentRect);
            Logger.Trace(
                $"  当前屏幕 Rect: ({currentRect.Left},{currentRect.Top}) " +
                $"{currentRect.Width}×{currentRect.Height}", Cat);

            NativeMethods.SetParent(hwnd, IntPtr.Zero);

            long style = NativeMethods.GetWindowLongSafe(hwnd, NativeConstants.GWL_STYLE);
            style &= ~(long)NativeConstants.WS_CHILD;
            style |= (long)(NativeConstants.WS_POPUP | NativeConstants.WS_CAPTION
                            | NativeConstants.WS_SYSMENU | NativeConstants.WS_VISIBLE);
            style &= ~(long)(NativeConstants.WS_THICKFRAME
                             | NativeConstants.WS_MINIMIZEBOX
                             | NativeConstants.WS_MAXIMIZEBOX);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE, style);

            NativeMethods.SetWindowPos(hwnd, NativeConstants.HWND_TOPMOST,
                currentRect.Left, currentRect.Top,
                currentRect.Width, currentRect.Height,
                NativeConstants.SWP_FRAMECHANGED | NativeConstants.SWP_SHOWWINDOW);

            window.IsEmbedded = false;
            Logger.Info($"TemporaryUnembed 成功: \"{window.Title}\"  hwnd=0x{hwnd:X}", Cat);
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error("TemporaryUnembed 异常", ex, Cat);
            return false;
        }
    }

    /// <summary>
    /// 将临时脱嵌的窗口重新嵌入宿主
    /// </summary>
    public bool ReEmbed(IntPtr hwnd)
    {
        var window = FindByHandle(hwnd);
        if (window == null || window.IsEmbedded)
        {
            Logger.Warning(
                $"ReEmbed: hwnd=0x{hwnd:X} 未找到或已嵌入，跳过。", Cat);
            return false;
        }

        Logger.Debug($"ReEmbed: \"{window.Title}\"  hwnd=0x{hwnd:X}", Cat);
        try
        {
            long style = NativeMethods.GetWindowLongSafe(hwnd, NativeConstants.GWL_STYLE);
            style &= ~(long)(NativeConstants.WS_CAPTION | NativeConstants.WS_THICKFRAME
                             | NativeConstants.WS_SYSMENU | NativeConstants.WS_MINIMIZEBOX
                             | NativeConstants.WS_MAXIMIZEBOX | NativeConstants.WS_POPUP
                             | NativeConstants.WS_BORDER | NativeConstants.WS_DLGFRAME);
            style |= (long)(NativeConstants.WS_CHILD | NativeConstants.WS_VISIBLE);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE, style);

            long exStyle = NativeMethods.GetWindowLongSafe(hwnd, NativeConstants.GWL_EXSTYLE);
            exStyle &= ~(long)(NativeConstants.WS_EX_APPWINDOW | NativeConstants.WS_EX_TOOLWINDOW
                               | NativeConstants.WS_EX_TOPMOST);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_EXSTYLE, exStyle);

            NativeMethods.SetParent(hwnd, HostHwnd);

            NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeConstants.SWP_NOMOVE | NativeConstants.SWP_NOSIZE
                | NativeConstants.SWP_NOZORDER | NativeConstants.SWP_FRAMECHANGED);

            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_SHOW);

            window.IsEmbedded = true;
            Logger.Info($"ReEmbed 成功: \"{window.Title}\"  hwnd=0x{hwnd:X}", Cat);
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error("ReEmbed 异常", ex, Cat);
            return false;
        }
    }

    /// <summary>
    /// 编程式发起窗口拖拽：临时脱嵌窗口并启动系统 Move 操作。
    /// </summary>
    public bool StartProgrammaticDrag(IntPtr hwnd)
    {
        Logger.Debug($"StartProgrammaticDrag  hwnd=0x{hwnd:X}", Cat);

        if (!TemporaryUnembed(hwnd))
        {
            Logger.Error($"StartProgrammaticDrag: TemporaryUnembed 失败，放弃。", Cat);
            return false;
        }

        // PostMessage SC_MOVE 进入系统拖拽循环
        bool posted = NativeMethods.PostMessage(hwnd,
            (uint)NativeConstants.WM_SYSCOMMAND,
            (IntPtr)(NativeConstants.SC_MOVE | 0x0002),
            IntPtr.Zero);

        Logger.Debug(
            $"PostMessage(WM_SYSCOMMAND, SC_MOVE) → posted={posted}  hwnd=0x{hwnd:X}", Cat);
        return true;
    }

    // ── 释放 ──────────────────────────────────────────────

    public void ReleaseWindow(IntPtr hwnd)
    {
        var window = ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);
        if (window == null)
        {
            Logger.Warning($"ReleaseWindow: hwnd=0x{hwnd:X} 不在管理列表中，跳过。", Cat);
            return;
        }

        Logger.Info(
            $"释放窗口: \"{window.Title}\"  hwnd=0x{hwnd:X}  isEmbedded={window.IsEmbedded}", Cat);

        if (!window.IsEmbedded)
        {
            Logger.Debug("  窗口处于临时脱嵌状态，直接还原原始状态。", Cat);
            RestoreOriginalState(window);
        }
        else
        {
            UnembedWindow(window);
        }

        window.IsManaged  = false;
        window.IsEmbedded = false;
        window.SlotIndex  = -1;
        ManagedWindows.Remove(window);

        int count = Interlocked.Increment(ref _releaseCount);
        Logger.Info(
            $"[#Release {count}] 窗口已释放: \"{window.Title}\"  " +
            $"剩余管理数量: {ManagedWindows.Count}", Cat);

        WindowReleased?.Invoke(window);
        LayoutChanged?.Invoke();
    }

    private void UnembedWindow(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;
        Logger.Trace($"UnembedWindow hwnd=0x{hwnd:X}", Cat);
        try
        {
            NativeMethods.SetParent(hwnd, IntPtr.Zero);
            RestoreOriginalState(window);
        }
        catch (Exception ex)
        {
            Logger.Error("UnembedWindow 异常", ex, Cat);
        }
    }

    private void RestoreOriginalState(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;
        Logger.Trace(
            $"RestoreOriginalState hwnd=0x{hwnd:X}  " +
            $"style=0x{window.OriginalStyle:X}  exStyle=0x{window.OriginalExStyle:X}  " +
            $"rect=({window.OriginalRect.Left},{window.OriginalRect.Top}," +
            $"{window.OriginalRect.Width}×{window.OriginalRect.Height})", Cat);
        try
        {
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE,   window.OriginalStyle);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_EXSTYLE, window.OriginalExStyle);

            NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeConstants.SWP_NOMOVE | NativeConstants.SWP_NOSIZE |
                NativeConstants.SWP_NOZORDER | NativeConstants.SWP_FRAMECHANGED);

            var r = window.OriginalRect;
            NativeMethods.MoveWindow(hwnd, r.Left, r.Top, r.Width, r.Height, true);

            if (window.WasMaximized)
            {
                Logger.Trace($"  还原最大化状态。", Cat);
                NativeMethods.ShowWindow(hwnd, 3);
            }

            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_SHOW);
            Logger.Trace($"  RestoreOriginalState 完成。", Cat);
        }
        catch (Exception ex)
        {
            Logger.Error("RestoreOriginalState 异常", ex, Cat);
        }
    }

    public void ReleaseAll()
    {
        Logger.Info($"ReleaseAll: 即将释放 {ManagedWindows.Count} 个窗口。", Cat);
        foreach (var w in ManagedWindows.ToList())
            ReleaseWindow(w.Handle);
        Logger.Info("ReleaseAll 完成。", Cat);
    }

    // ── 位置 / 激活 ──────────────────────────────────────

    public void RepositionEmbedded(IntPtr hwnd, int x, int y, int width, int height)
    {
        int count = Interlocked.Increment(ref _repositionCount);
        // 非常密集，只写 Trace
        Logger.Trace(
            $"RepositionEmbedded #{count}  hwnd=0x{hwnd:X}  " +
            $"→ ({x},{y}) {width}×{height}", Cat);
        NativeMethods.MoveWindow(hwnd, x, y, width, height, true);
    }

    public void ActivateWindow(IntPtr hwnd)
    {
        Logger.Debug($"ActivateWindow hwnd=0x{hwnd:X}", Cat);
        NativeMethods.BringWindowToTop(hwnd);
        NativeMethods.SetFocus(hwnd);
        foreach (var w in ManagedWindows)
            w.IsActive = w.Handle == hwnd;
    }

    public void ShowOnly(IntPtr hwnd)
    {
        Logger.Debug($"ShowOnly hwnd=0x{hwnd:X}", Cat);
        foreach (var w in ManagedWindows)
        {
            if (w.Handle == hwnd)
            {
                NativeMethods.ShowWindow(w.Handle, NativeConstants.SW_SHOW);
                NativeMethods.BringWindowToTop(w.Handle);
                w.IsActive = true;
            }
            else
            {
                NativeMethods.ShowWindow(w.Handle, 0);
                w.IsActive = false;
            }
        }
    }

    public void ShowAll()
    {
        Logger.Trace($"ShowAll: {ManagedWindows.Count} 个窗口", Cat);
        foreach (var w in ManagedWindows)
            NativeMethods.ShowWindow(w.Handle, NativeConstants.SW_SHOW);
    }

    // ── 交换 ──────────────────────────────────────────────

    /// <summary>
    /// 交换两个窗口在 ManagedWindows 列表中的位置。
    /// </summary>
    public void SwapOrder(IntPtr hwndA, IntPtr hwndB)
    {
        var a = ManagedWindows.FirstOrDefault(w => w.Handle == hwndA);
        var b = ManagedWindows.FirstOrDefault(w => w.Handle == hwndB);

        if (a == null || b == null || a == b)
        {
            Logger.Warning(
                $"SwapOrder: 无法交换  a=0x{hwndA:X} ({a?.Title ?? "null"})  " +
                $"b=0x{hwndB:X} ({b?.Title ?? "null"})", Cat);
            return;
        }

        int idxA = ManagedWindows.IndexOf(a);
        int idxB = ManagedWindows.IndexOf(b);
        Logger.Info(
            $"SwapOrder: \"{a.Title}\"[{idxA}] ⇄ \"{b.Title}\"[{idxB}]", Cat);

        if (idxA > idxB)
        {
            (idxA, idxB) = (idxB, idxA);
            (a, b) = (b, a);
        }

        ManagedWindows.RemoveAt(idxB);
        ManagedWindows.RemoveAt(idxA);
        ManagedWindows.Insert(idxA, b);
        ManagedWindows.Insert(idxB, a);

        Logger.Debug(
            $"  交换完成。新顺序: {string.Join(", ", ManagedWindows.Select(w => $"\"{w.Title}\""))}", Cat);
    }

    // ── 事件处理 ──────────────────────────────────────────

    private void OnWindowDestroyed(IntPtr hwnd)
    {
        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            var window = ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);
            if (window != null)
            {
                Logger.Warning(
                    $"托管窗口被销毁: \"{window.Title}\"  hwnd=0x{hwnd:X}", Cat);
                window.IsManaged  = false;
                window.IsEmbedded = false;
                ManagedWindows.Remove(window);
                LayoutChanged?.Invoke();
            }
        });
    }

    private void OnWindowTitleChanged(IntPtr hwnd)
    {
        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            var w = ManagedWindows.FirstOrDefault(x => x.Handle == hwnd);
            if (w != null)
            {
                string oldTitle = w.Title;
                w.RefreshTitle();
                if (oldTitle != w.Title)
                    Logger.Trace($"窗口标题变更: 0x{hwnd:X}  \"{oldTitle}\" → \"{w.Title}\"", Cat);
            }
        });
    }

    public void Dispose()
    {
        Logger.Debug("WindowManagerService.Dispose()", Cat);
        Logger.Debug(
            $"统计 — 嵌入: {_embedCount}次  释放: {_releaseCount}次  " +
            $"Reposition: {_repositionCount}次", Cat);
        ReleaseAll();
        _winEventHook.WindowDestroyed    -= OnWindowDestroyed;
        _winEventHook.WindowTitleChanged -= OnWindowTitleChanged;
        GC.SuppressFinalize(this);
    }
}