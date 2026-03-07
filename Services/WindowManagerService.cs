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

    // ── 查询 ──

    /// <summary>
    /// 根据句柄查找 ManagedWindow 实例
    /// </summary>
    public ManagedWindow? FindByHandle(IntPtr hwnd)
        => ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);

    /// <summary>
    /// 判断窗口是否已被管理
    /// </summary>
    public bool IsManaged(IntPtr hwnd)
        => ManagedWindows.Any(w => w.Handle == hwnd);

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

    // ── 临时脱嵌 / 重新嵌入（用于拖拽互换和释放） ──

    /// <summary>
    /// 临时将窗口从宿主中脱离，使其成为顶层窗口以便自由拖拽。
    /// 窗口仍然保留在 ManagedWindows 列表中。
    /// </summary>
    public bool TemporaryUnembed(IntPtr hwnd)
    {
        var window = FindByHandle(hwnd);
        if (window == null || !window.IsEmbedded) return false;

        try
        {
            // 先获取窗口当前在宿主客户区内的位置
            NativeMethods.GetWindowRect(hwnd, out RECT currentRect);

            // 脱离宿主
            NativeMethods.SetParent(hwnd, IntPtr.Zero);

            // 恢复为可拖拽的顶层窗口样式（保留精简外观）
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
            return true;
        }
        catch { return false; }
    }

    /// <summary>
    /// 将临时脱嵌的窗口重新嵌入宿主
    /// </summary>
    public bool ReEmbed(IntPtr hwnd)
    {
        var window = FindByHandle(hwnd);
        if (window == null || window.IsEmbedded) return false;

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

            // TemporaryUnembed 使用了 HWND_TOPMOST；在 SetParent 之前先用 HWND_NOTOPMOST
            // 清除置顶 Z 序，否则 ExStyle 修改无法完全消除置顶效果
            NativeMethods.SetWindowPos(hwnd, NativeConstants.HWND_NOTOPMOST, 0, 0, 0, 0,
                NativeConstants.SWP_NOMOVE | NativeConstants.SWP_NOSIZE | NativeConstants.SWP_NOACTIVATE);

            // 在 SetParent 之前先将屏幕坐标转换为宿主客户区坐标，
            // 避免 SetParent 后位置值被误解为客户区坐标而产生偏移
            NativeMethods.GetWindowRect(hwnd, out RECT screenRect);
            var clientTL = new POINT { X = screenRect.Left, Y = screenRect.Top };
            NativeMethods.ScreenToClient(HostHwnd, ref clientTL);

            NativeMethods.SetParent(hwnd, HostHwnd);

            // 用换算后的正确客户区坐标定位，并触发 WM_NCCALCSIZE 更新非客户区
            NativeMethods.SetWindowPos(hwnd, IntPtr.Zero,
                clientTL.X, clientTL.Y, screenRect.Width, screenRect.Height,
                NativeConstants.SWP_NOZORDER | NativeConstants.SWP_FRAMECHANGED);

            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_SHOW);

            window.IsEmbedded = true;
            return true;
        }
        catch { return false; }
    }

    /// <summary>
    /// 编程式发起窗口拖拽：临时脱嵌窗口并启动系统 Move 操作。
    /// 调用后窗口进入模态拖拽循环，直到用户松开鼠标。
    /// </summary>
    public bool StartProgrammaticDrag(IntPtr hwnd)
    {
        if (!TemporaryUnembed(hwnd))
            return false;

        // 使用 PostMessage 而非 SendMessage，让消息进入目标窗口的消息队列，
        // 这样不会阻塞我们的线程。
        // SC_MOVE | 0x02 表示通过鼠标移动（系统会自动捕获鼠标位置）
        NativeMethods.PostMessage(hwnd,
            (uint)NativeConstants.WM_SYSCOMMAND,
            (IntPtr)(NativeConstants.SC_MOVE | 0x0002),
            IntPtr.Zero);

        return true;
    }

    // ── 释放 ──

    public void ReleaseWindow(IntPtr hwnd)
    {
        var window = ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);
        if (window == null) return;

        // 如果窗口处于临时脱嵌状态，直接还原即可
        if (!window.IsEmbedded)
        {
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

        WindowReleased?.Invoke(window);
        LayoutChanged?.Invoke();
    }

    private void UnembedWindow(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;
        try
        {
            NativeMethods.SetParent(hwnd, IntPtr.Zero);
            RestoreOriginalState(window);
        }
        catch { }
    }

    /// <summary>
    /// 还原窗口的原始样式和位置
    /// </summary>
    private void RestoreOriginalState(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;
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
    {
        NativeMethods.MoveWindow(hwnd, x, y, width, height, true);
        // 强制触发 WM_NCCALCSIZE，确保窗口重新计算非客户区与命中测试区域，
        // 避免移动后标题栏判定位置停留在旧坐标
        NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
            NativeConstants.SWP_NOMOVE | NativeConstants.SWP_NOSIZE |
            NativeConstants.SWP_NOZORDER | NativeConstants.SWP_FRAMECHANGED);
    }

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

    // ── ✨ 交换两个托管窗口在列表中的顺序 ──

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