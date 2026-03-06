using System.Collections.ObjectModel;
using System.Windows;
using WindowPilot.Models;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 窗口管理核心：通过 SetParent 将外部窗口嵌入到宿主窗口内部
/// </summary>
public class WindowManagerService : IDisposable
{
    private readonly WinEventHookService _winEventHook;

    /// <summary>宿主窗口句柄（我们的主窗口）</summary>
    public IntPtr HostHwnd { get; set; }

    /// <summary>当前被管理的窗口列表</summary>
    public ObservableCollection<ManagedWindow> ManagedWindows { get; } = new();

    public event Action<IntPtr, string>? ManageFailed;
    public event Action<ManagedWindow>? WindowManaged;
    public event Action<ManagedWindow>? WindowReleased;
    public event Action? LayoutChanged;

    public WindowManagerService(WinEventHookService winEventHook)
    {
        _winEventHook = winEventHook;
        _winEventHook.WindowDestroyed    += OnWindowDestroyed;
        _winEventHook.WindowTitleChanged += OnWindowTitleChanged;
    }

    // ── 嵌入 ──

    /// <summary>
    /// 尝试嵌入一个窗口到宿主内部
    /// </summary>
    public bool TryManageWindow(IntPtr hwnd)
    {
        if (HostHwnd == IntPtr.Zero)
        {
            ManageFailed?.Invoke(hwnd, "宿主窗口未就绪");
            return false;
        }

        if (ManagedWindows.Any(w => w.Handle == hwnd))
            return true;

        if (hwnd == HostHwnd)
            return false;

        var window = new ManagedWindow(hwnd);

        if (!window.CanWeControl())
        {
            ManageFailed?.Invoke(hwnd, $"无法控制 \"{window.Title}\"：可能以管理员权限运行。");
            return false;
        }

        window.SaveOriginalState();

        if (NativeMethods.IsZoomed(hwnd))
            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_RESTORE);
        if (NativeMethods.IsIconic(hwnd))
            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_RESTORE);

        if (!EmbedWindow(window))
        {
            ManageFailed?.Invoke(hwnd, $"嵌入 \"{window.Title}\" 失败：该窗口可能不支持嵌入。");
            return false;
        }

        window.IsManaged  = true;
        window.IsEmbedded = true;
        ManagedWindows.Add(window);
        WindowManaged?.Invoke(window);
        LayoutChanged?.Invoke();
        return true;
    }

    private bool EmbedWindow(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;
        try
        {
            long style = window.OriginalStyle;
            style &= ~(long)(
                NativeConstants.WS_CAPTION    |
                NativeConstants.WS_THICKFRAME |
                NativeConstants.WS_SYSMENU    |
                NativeConstants.WS_MINIMIZEBOX|
                NativeConstants.WS_MAXIMIZEBOX|
                NativeConstants.WS_POPUP      |
                NativeConstants.WS_BORDER     |
                NativeConstants.WS_DLGFRAME);
            style |= (long)NativeConstants.WS_CHILD;
            style |= (long)NativeConstants.WS_VISIBLE;
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE, style);

            long exStyle = window.OriginalExStyle;
            exStyle &= ~(long)(NativeConstants.WS_EX_APPWINDOW | NativeConstants.WS_EX_TOOLWINDOW);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_EXSTYLE, exStyle);

            if (NativeMethods.SetParent(hwnd, HostHwnd) == IntPtr.Zero)
                return false;

            NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeConstants.SWP_NOMOVE | NativeConstants.SWP_NOSIZE |
                NativeConstants.SWP_NOZORDER | NativeConstants.SWP_FRAMECHANGED);

            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_SHOW);
            return true;
        }
        catch { return false; }
    }

    // ── 释放 ──

    public void ReleaseWindow(IntPtr hwnd)
    {
        var window = ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);
        if (window == null) return;

        UnembedWindow(window);

        window.IsManaged  = false;
        window.IsEmbedded = false;
        window.SlotIndex  = -1;
        ManagedWindows.Remove(window);

        WindowReleased?.Invoke(window);
        LayoutChanged?.Invoke();
    }

    private void UnembedWindow(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;
        try
        {
            NativeMethods.SetParent(hwnd, IntPtr.Zero);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE,   window.OriginalStyle);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_EXSTYLE, window.OriginalExStyle);

            NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeConstants.SWP_NOMOVE | NativeConstants.SWP_NOSIZE |
                NativeConstants.SWP_NOZORDER | NativeConstants.SWP_FRAMECHANGED);

            var r = window.OriginalRect;
            NativeMethods.MoveWindow(hwnd, r.Left, r.Top, r.Width, r.Height, true);

            if (window.WasMaximized)
                NativeMethods.ShowWindow(hwnd, 3);

            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_SHOW);
        }
        catch { }
    }

    public void ReleaseAll()
    {
        foreach (var w in ManagedWindows.ToList())
            ReleaseWindow(w.Handle);
    }

    // ── 位置 / 激活 ──

    public void RepositionEmbedded(IntPtr hwnd, int x, int y, int width, int height)
        => NativeMethods.MoveWindow(hwnd, x, y, width, height, true);

    public void ActivateWindow(IntPtr hwnd)
    {
        NativeMethods.BringWindowToTop(hwnd);
        NativeMethods.SetFocus(hwnd);
        foreach (var w in ManagedWindows)
            w.IsActive = w.Handle == hwnd;
    }

    public void ShowOnly(IntPtr hwnd)
    {
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
        foreach (var w in ManagedWindows)
            NativeMethods.ShowWindow(w.Handle, NativeConstants.SW_SHOW);
    }

    // ── ✨ 新功能：交换两个托管窗口在列表中的顺序 ──

    /// <summary>
    /// 交换两个窗口在 ManagedWindows 列表中的位置。
    /// 列表顺序决定了布局槽位和堆叠切换顺序。
    /// </summary>
    public void SwapOrder(IntPtr hwndA, IntPtr hwndB)
    {
        var a = ManagedWindows.FirstOrDefault(w => w.Handle == hwndA);
        var b = ManagedWindows.FirstOrDefault(w => w.Handle == hwndB);
        if (a == null || b == null || a == b) return;

        int idxA = ManagedWindows.IndexOf(a);
        int idxB = ManagedWindows.IndexOf(b);

        // 确保低索引在前，简化逻辑
        if (idxA > idxB)
        {
            (idxA, idxB) = (idxB, idxA);
            (a, b) = (b, a);
        }

        // 先移除高索引再移除低索引，避免索引偏移
        ManagedWindows.RemoveAt(idxB);
        ManagedWindows.RemoveAt(idxA);
        ManagedWindows.Insert(idxA, b);
        ManagedWindows.Insert(idxB, a);
    }

    // ── 事件处理 ──

    private void OnWindowDestroyed(IntPtr hwnd)
    {
        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            var window = ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);
            if (window != null)
            {
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
            ManagedWindows.FirstOrDefault(w => w.Handle == hwnd)?.RefreshTitle());
    }

    public void Dispose()
    {
        ReleaseAll();
        _winEventHook.WindowDestroyed    -= OnWindowDestroyed;
        _winEventHook.WindowTitleChanged -= OnWindowTitleChanged;
        GC.SuppressFinalize(this);
    }
}