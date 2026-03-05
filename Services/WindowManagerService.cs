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
        _winEventHook.WindowDestroyed += OnWindowDestroyed;
        _winEventHook.WindowTitleChanged += OnWindowTitleChanged;
    }

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

        // 防止重复
        if (ManagedWindows.Any(w => w.Handle == hwnd))
            return true;

        // 不能嵌入自己
        if (hwnd == HostHwnd)
            return false;

        var window = new ManagedWindow(hwnd);

        // 权限检查
        if (!window.CanWeControl())
        {
            ManageFailed?.Invoke(hwnd, $"无法控制 \"{window.Title}\"：可能以管理员权限运行。");
            return false;
        }

        // 保存原始状态
        window.SaveOriginalState();

        // 如果最大化/最小化，先还原
        if (NativeMethods.IsZoomed(hwnd))
            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_RESTORE);
        if (NativeMethods.IsIconic(hwnd))
            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_RESTORE);

        // ── 执行嵌入 ──
        if (!EmbedWindow(window))
        {
            ManageFailed?.Invoke(hwnd, $"嵌入 \"{window.Title}\" 失败：该窗口可能不支持嵌入。");
            return false;
        }

        window.IsManaged = true;
        window.IsEmbedded = true;
        ManagedWindows.Add(window);
        WindowManaged?.Invoke(window);
        LayoutChanged?.Invoke();
        return true;
    }

    /// <summary>
    /// 执行 SetParent 嵌入
    /// </summary>
    private bool EmbedWindow(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;

        try
        {
            // 1) 修改窗口样式：移除标题栏、边框、弹出等，添加子窗口标志
            long style = window.OriginalStyle;
            style &= ~(long)(
                NativeConstants.WS_CAPTION |
                NativeConstants.WS_THICKFRAME |
                NativeConstants.WS_SYSMENU |
                NativeConstants.WS_MINIMIZEBOX |
                NativeConstants.WS_MAXIMIZEBOX |
                NativeConstants.WS_POPUP |
                NativeConstants.WS_BORDER |
                NativeConstants.WS_DLGFRAME
            );
            style |= (long)NativeConstants.WS_CHILD;
            style |= (long)NativeConstants.WS_VISIBLE;

            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE, style);

            // 2) 移除扩展样式中的弹窗/工具窗口标志
            long exStyle = window.OriginalExStyle;
            exStyle &= ~(long)(
                NativeConstants.WS_EX_APPWINDOW |
                NativeConstants.WS_EX_TOOLWINDOW
            );
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_EXSTYLE, exStyle);

            // 3) SetParent 嵌入
            IntPtr result = NativeMethods.SetParent(hwnd, HostHwnd);
            if (result == IntPtr.Zero)
                return false;

            // 4) 触发样式重绘
            NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeConstants.SWP_NOMOVE | NativeConstants.SWP_NOSIZE |
                NativeConstants.SWP_NOZORDER | NativeConstants.SWP_FRAMECHANGED);

            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_SHOW);

            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 释放窗口：还原原始父窗口和样式
    /// </summary>
    public void ReleaseWindow(IntPtr hwnd)
    {
        var window = ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);
        if (window == null) return;

        UnembedWindow(window);

        window.IsManaged = false;
        window.IsEmbedded = false;
        window.SlotIndex = -1;
        ManagedWindows.Remove(window);

        WindowReleased?.Invoke(window);
        LayoutChanged?.Invoke();
    }

    /// <summary>
    /// 执行反嵌入：还原窗口到桌面
    /// </summary>
    private void UnembedWindow(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;

        try
        {
            // 脱离父窗口
            NativeMethods.SetParent(hwnd, IntPtr.Zero);

            // 还原原始窗口样式
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE, window.OriginalStyle);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_EXSTYLE, window.OriginalExStyle);

            // 触发样式重绘
            NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeConstants.SWP_NOMOVE | NativeConstants.SWP_NOSIZE |
                NativeConstants.SWP_NOZORDER | NativeConstants.SWP_FRAMECHANGED);

            // 还原原始位置
            var r = window.OriginalRect;
            NativeMethods.MoveWindow(hwnd, r.Left, r.Top, r.Width, r.Height, true);

            // 如果原来是最大化的，还原最大化
            if (window.WasMaximized)
                NativeMethods.ShowWindow(hwnd, 3);

            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_SHOW);
        }
        catch { }
    }

    /// <summary>
    /// 调整嵌入窗口在宿主内的位置和大小（像素坐标，相对于宿主客户区）
    /// </summary>
    public void RepositionEmbedded(IntPtr hwnd, int x, int y, int width, int height)
    {
        NativeMethods.MoveWindow(hwnd, x, y, width, height, true);
    }

    /// <summary>
    /// 激活（置顶）指定的嵌入窗口
    /// </summary>
    public void ActivateWindow(IntPtr hwnd)
    {
        // 将该窗口带到最顶层
        NativeMethods.BringWindowToTop(hwnd);
        NativeMethods.SetFocus(hwnd);

        foreach (var w in ManagedWindows)
            w.IsActive = w.Handle == hwnd;
    }

    /// <summary>
    /// 显示指定嵌入窗口，隐藏其他（堆叠模式用）
    /// </summary>
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

    /// <summary>
    /// 显示所有嵌入窗口（分区模式用）
    /// </summary>
    public void ShowAll()
    {
        foreach (var w in ManagedWindows)
            NativeMethods.ShowWindow(w.Handle, NativeConstants.SW_SHOW);
    }

    public void ReleaseAll()
    {
        foreach (var w in ManagedWindows.ToList())
            ReleaseWindow(w.Handle);
    }

    // ── 事件处理 ──

    private void OnWindowDestroyed(IntPtr hwnd)
    {
        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            var window = ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);
            if (window != null)
            {
                window.IsManaged = false;
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
            var window = ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);
            window?.RefreshTitle();
        });
    }

    public void Dispose()
    {
        ReleaseAll();
        _winEventHook.WindowDestroyed -= OnWindowDestroyed;
        _winEventHook.WindowTitleChanged -= OnWindowTitleChanged;
        GC.SuppressFinalize(this);
    }
}
