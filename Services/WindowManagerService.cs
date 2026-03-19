using System.Collections.ObjectModel;
using System.Windows;
using WindowPilot.Diagnostics;
using WindowPilot.Models;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 窗口管理核心，通过 <c>SetParent</c> 将外部窗口嵌入宿主窗口。
/// 负责窗口的嵌入、释放、重定位、激活和编程式拖拽。
/// </summary>
public class WindowManagerService : IDisposable
{
    private const string Cat = "WindowManager";

    private readonly WinEventHookService _winEventHook;

    // 宿主窗口句柄（本程序的主窗口），所有子窗口都以此为父
    public IntPtr HostHwnd { get; set; }

    // 当前托管窗口列表，绑定到侧边栏 ListBox
    public ObservableCollection<ManagedWindow> ManagedWindows { get; } = new();

    public event Action<IntPtr, string>? ManageFailed;    // 嵌入失败，附带原因描述
    public event Action<ManagedWindow>?  WindowManaged;   // 窗口成功纳入托管
    public event Action<ManagedWindow>?  WindowReleased;  // 窗口已从托管中释放
    public event Action?                 LayoutChanged;   // 列表结构变化，需要重新布局

    // 统计数据
    private int _embedCount;      // 历史嵌入次数
    private int _releaseCount;    // 历史释放次数
    private int _repositionCount; // 历史重定位次数（高频）

    /// <summary>
    /// 构造窗口管理服务，订阅窗口销毁和标题变化事件。
    /// </summary>
    /// <param name="winEventHook">WinEvent 钩子服务，提供 WindowDestroyed 和 WindowTitleChanged 事件。</param>
    public WindowManagerService(WinEventHookService winEventHook)
    {
        Logger.Debug("WindowManagerService 构造中…", Cat);
        _winEventHook = winEventHook;
        _winEventHook.WindowDestroyed    += OnWindowDestroyed;
        _winEventHook.WindowTitleChanged += OnWindowTitleChanged;
        Logger.Debug("已订阅 WindowDestroyed / WindowTitleChanged 事件。", Cat);
    }

    // 查询

    /// <summary>
    /// 按句柄在托管列表中查找对应的 <see cref="ManagedWindow"/>。
    /// </summary>
    /// <param name="hwnd">目标窗口句柄。</param>
    /// <returns>找到时返回对应对象，否则返回 null。</returns>
    public ManagedWindow? FindByHandle(IntPtr hwnd)
        => ManagedWindows.FirstOrDefault(w => w.Handle == hwnd);

    /// <summary>
    /// 判断指定句柄是否已在托管列表中。
    /// </summary>
    /// <param name="hwnd">要检查的窗口句柄。</param>
    public bool IsManaged(IntPtr hwnd)
        => ManagedWindows.Any(w => w.Handle == hwnd);

    // 嵌入

    /// <summary>
    /// 尝试将外部窗口嵌入宿主，纳入托管。
    /// 会自动保存原始状态、还原最大化/最小化、修改样式并重设父窗口。
    /// </summary>
    /// <param name="hwnd">要嵌入的外部窗口句柄。</param>
    /// <returns>嵌入成功（包括已在列表中的情况）返回 true，失败返回 false。</returns>
    public bool TryManageWindow(IntPtr hwnd)
    {
        Logger.Debug($"TryManageWindow  hwnd=0x{hwnd:X}", Cat);

        if (HostHwnd == IntPtr.Zero)
        {
            Logger.Error("TryManageWindow 失败：宿主窗口未就绪（HostHwnd == Zero）。", Cat);
            ManageFailed?.Invoke(hwnd, "宿主窗口未就绪");
            return false;
        }

        // 已在列表中则视为成功（幂等）
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

        // 快照原始状态，供释放时还原使用
        window.SaveOriginalState();
        Logger.Trace(
            $"  原始状态已保存: parent=0x{window.OriginalParent:X}  " +
            $"style=0x{window.OriginalStyle:X}  exStyle=0x{window.OriginalExStyle:X}  " +
            $"rect=({window.OriginalRect.Left},{window.OriginalRect.Top}," +
            $"{window.OriginalRect.Width}×{window.OriginalRect.Height})  " +
            $"wasMaximized={window.WasMaximized}", Cat);

        // 嵌入前必须先还原窗口状态，最大化/最小化窗口无法正常设置父窗口
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

    /// <summary>
    /// 修改窗口样式并调用 SetParent 将其嵌入宿主。
    /// </summary>
    /// <param name="window">要嵌入的 <see cref="ManagedWindow"/>，原始样式从其属性读取。</param>
    /// <returns>嵌入成功返回 true，SetParent 失败或发生异常时返回 false。</returns>
    private bool EmbedWindow(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;
        Logger.Trace($"EmbedWindow hwnd=0x{hwnd:X} → 修改样式并 SetParent", Cat);
        try
        {
            // 移除所有标准窗口装饰（标题栏、边框、系统菜单、最大/最小化按钮）
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
            // 添加 WS_CHILD 使其成为宿主的子窗口
            style |= (long)NativeConstants.WS_CHILD;
            style |= (long)NativeConstants.WS_VISIBLE;
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE, style);
            Logger.Trace($"  新 Style=0x{style:X}", Cat);

            // 移除 AppWindow/ToolWindow 标志，防止嵌入后仍出现在任务栏
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

            // SWP_FRAMECHANGED 通知系统重新计算非客户区，使样式变更生效
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

    // 临时脱嵌与重新嵌入

    /// <summary>
    /// 将嵌入的窗口临时从宿主中脱离，使其成为可自由拖拽的顶层窗口。
    /// 用于编程式拖拽和互换位置前的准备步骤。
    /// </summary>
    /// <param name="hwnd">要临时脱嵌的托管窗口句柄。</param>
    /// <returns>脱嵌成功返回 true，窗口未找到/未嵌入或发生异常时返回 false。</returns>
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
            // 先记录当前屏幕坐标，脱嵌后恢复到相同位置
            NativeMethods.GetWindowRect(hwnd, out RECT currentRect);
            Logger.Trace(
                $"  当前屏幕 Rect: ({currentRect.Left},{currentRect.Top}) " +
                $"{currentRect.Width}×{currentRect.Height}", Cat);

            // 解除父子关系，变为顶层窗口
            NativeMethods.SetParent(hwnd, IntPtr.Zero);

            // 恢复弹出式窗口样式，使其有基础的可拖拽外观
            long style = NativeMethods.GetWindowLongSafe(hwnd, NativeConstants.GWL_STYLE);
            style &= ~(long)NativeConstants.WS_CHILD;
            style |= (long)(NativeConstants.WS_POPUP | NativeConstants.WS_CAPTION
                            | NativeConstants.WS_SYSMENU | NativeConstants.WS_VISIBLE);
            // 不恢复调整大小边框，保持简洁外观
            style &= ~(long)(NativeConstants.WS_THICKFRAME
                             | NativeConstants.WS_MINIMIZEBOX
                             | NativeConstants.WS_MAXIMIZEBOX);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE, style);

            // 设为置顶并移到原来的屏幕位置
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
    /// 将临时脱嵌的窗口重新嵌入宿主，恢复子窗口状态。
    /// </summary>
    /// <param name="hwnd">要重新嵌入的窗口句柄，必须是已临时脱嵌（IsEmbedded == false）的托管窗口。</param>
    /// <returns>重嵌成功返回 true，窗口未找到/已嵌入或发生异常时返回 false。</returns>
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
            // 移除所有顶层窗口装饰，还原为子窗口样式
            long style = NativeMethods.GetWindowLongSafe(hwnd, NativeConstants.GWL_STYLE);
            style &= ~(long)(NativeConstants.WS_CAPTION | NativeConstants.WS_THICKFRAME
                             | NativeConstants.WS_SYSMENU | NativeConstants.WS_MINIMIZEBOX
                             | NativeConstants.WS_MAXIMIZEBOX | NativeConstants.WS_POPUP
                             | NativeConstants.WS_BORDER | NativeConstants.WS_DLGFRAME);
            style |= (long)(NativeConstants.WS_CHILD | NativeConstants.WS_VISIBLE);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE, style);

            // 移除 TOPMOST 标志，子窗口不应保持置顶
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
    /// 编程式发起窗口拖拽：先临时脱嵌，再通过 WM_SYSCOMMAND + SC_MOVE 进入系统拖拽循环。
    /// </summary>
    /// <param name="hwnd">要发起拖拽的托管窗口句柄。</param>
    /// <returns>成功进入拖拽循环返回 true，TemporaryUnembed 失败时返回 false。</returns>
    public bool StartProgrammaticDrag(IntPtr hwnd)
    {
        Logger.Debug($"StartProgrammaticDrag  hwnd=0x{hwnd:X}", Cat);

        if (!TemporaryUnembed(hwnd))
        {
            Logger.Error($"StartProgrammaticDrag: TemporaryUnembed 失败，放弃。", Cat);
            return false;
        }

        // SC_MOVE | 0x0002 进入系统 Move 模式，等效于用户点击标题栏拖拽
        bool posted = NativeMethods.PostMessage(hwnd,
            (uint)NativeConstants.WM_SYSCOMMAND,
            (IntPtr)(NativeConstants.SC_MOVE | 0x0002),
            IntPtr.Zero);

        Logger.Debug(
            $"PostMessage(WM_SYSCOMMAND, SC_MOVE) → posted={posted}  hwnd=0x{hwnd:X}", Cat);
        return true;
    }

    // 释放

    /// <summary>
    /// 将指定窗口从托管中释放，还原其原始样式、位置和父窗口关系。
    /// </summary>
    /// <param name="hwnd">要释放的托管窗口句柄。</param>
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
            // 临时脱嵌状态直接还原原始状态，无需再次 SetParent
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

    /// <summary>
    /// 解除父子关系并还原窗口原始状态。
    /// </summary>
    /// <param name="window">要解除嵌入的托管窗口对象。</param>
    private void UnembedWindow(ManagedWindow window)
    {
        IntPtr hwnd = window.Handle;
        Logger.Trace($"UnembedWindow hwnd=0x{hwnd:X}", Cat);
        try
        {
            // 先断开父子关系，再还原样式和位置
            NativeMethods.SetParent(hwnd, IntPtr.Zero);
            RestoreOriginalState(window);
        }
        catch (Exception ex)
        {
            Logger.Error("UnembedWindow 异常", ex, Cat);
        }
    }

    /// <summary>
    /// 将窗口的样式、扩展样式和屏幕位置还原为嵌入前的快照值。
    /// </summary>
    /// <param name="window">保存了原始状态快照的 <see cref="ManagedWindow"/>。</param>
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
            // 还原样式，然后强制刷新非客户区
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_STYLE,   window.OriginalStyle);
            NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_EXSTYLE, window.OriginalExStyle);

            NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeConstants.SWP_NOMOVE | NativeConstants.SWP_NOSIZE |
                NativeConstants.SWP_NOZORDER | NativeConstants.SWP_FRAMECHANGED);

            // 还原到嵌入前的屏幕位置
            var r = window.OriginalRect;
            NativeMethods.MoveWindow(hwnd, r.Left, r.Top, r.Width, r.Height, true);

            if (window.WasMaximized)
            {
                // 原来是最大化状态则恢复最大化
                Logger.Trace($"  还原最大化状态。", Cat);
                NativeMethods.ShowWindow(hwnd, 3); // SW_MAXIMIZE = 3
            }

            NativeMethods.ShowWindow(hwnd, NativeConstants.SW_SHOW);
            Logger.Trace($"  RestoreOriginalState 完成。", Cat);
        }
        catch (Exception ex)
        {
            Logger.Error("RestoreOriginalState 异常", ex, Cat);
        }
    }

    // 释放所有托管窗口
    public void ReleaseAll()
    {
        Logger.Info($"ReleaseAll: 即将释放 {ManagedWindows.Count} 个窗口。", Cat);
        foreach (var w in ManagedWindows.ToList())
            ReleaseWindow(w.Handle);
        Logger.Info("ReleaseAll 完成。", Cat);
    }

    // 位置与激活

    /// <summary>
    /// 将嵌入窗口移动到指定的客户区矩形，对应布局槽位的坐标。
    /// </summary>
    /// <param name="hwnd">要重定位的嵌入窗口句柄。</param>
    /// <param name="x">目标区域左上角 X（相对于宿主客户区，像素）。</param>
    /// <param name="y">目标区域左上角 Y（相对于宿主客户区，像素）。</param>
    /// <param name="width">目标宽度（像素）。</param>
    /// <param name="height">目标高度（像素）。</param>
    public void RepositionEmbedded(IntPtr hwnd, int x, int y, int width, int height)
    {
        int count = Interlocked.Increment(ref _repositionCount);
        Logger.Trace(
            $"RepositionEmbedded #{count}  hwnd=0x{hwnd:X}  " +
            $"→ ({x},{y}) {width}×{height}", Cat);
        NativeMethods.MoveWindow(hwnd, x, y, width, height, true);
    }

    /// <summary>
    /// 将指定窗口提到前台并更新 IsActive 标志。
    /// 使用 <see cref="FocusEmbeddedWindow"/> 确保跨进程焦点生效。
    /// </summary>
    /// <param name="hwnd">要激活的托管窗口句柄。</param>
    public void ActivateWindow(IntPtr hwnd)
    {
        Logger.Debug($"ActivateWindow hwnd=0x{hwnd:X}", Cat);
        FocusEmbeddedWindow(hwnd);                          // ← 替换原来的 BringWindowToTop + SetFocus
        foreach (var w in ManagedWindows)
            w.IsActive = w.Handle == hwnd;
    }
    
    /// <summary>
    /// 将键盘焦点安全地转移到嵌入窗口。
    /// 跨进程嵌入时，必须先通过 <see cref="NativeMethods.AttachThreadInput"/> 临时
    /// 合并两个线程的输入队列，否则 <see cref="NativeMethods.SetFocus"/> 会被静默忽略。
    /// </summary>
    /// <param name="hwnd">要获得焦点的嵌入窗口句柄。</param>
    public void FocusEmbeddedWindow(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero) return;
 
        uint targetTid  = NativeMethods.GetWindowThreadProcessId(hwnd, out _);
        uint selfTid    = NativeMethods.GetCurrentThreadId();
        bool attached   = targetTid != selfTid &&
                          NativeMethods.AttachThreadInput(selfTid, targetTid, true);
 
        Logger.Trace(
            $"FocusEmbeddedWindow hwnd=0x{hwnd:X}  targetTid={targetTid}  " +
            $"selfTid={selfTid}  attached={attached}", Cat);
 
        NativeMethods.BringWindowToTop(hwnd);
        NativeMethods.SetFocus(hwnd);
 
        if (attached)
            NativeMethods.AttachThreadInput(selfTid, targetTid, false);
    }

    /// <summary>
    /// 只显示指定窗口（堆叠模式用）。
    /// </summary>
    /// <param name="hwnd">要置顶显示的托管窗口句柄。</param>
    public void ShowOnly(IntPtr hwnd)
    {
        Logger.Debug($"ShowOnly hwnd=0x{hwnd:X}", Cat);
        foreach (var w in ManagedWindows)
        {
            NativeMethods.ShowWindow(w.Handle, NativeConstants.SW_SHOW);
            w.IsActive = w.Handle == hwnd;
        }
        // 将活跃窗口提到最顶层，视觉上遮盖其他窗口
        NativeMethods.BringWindowToTop(hwnd);
        // 确保键盘焦点也跟随
        FocusEmbeddedWindow(hwnd);
    }

    // 显示所有托管窗口，用于切换到多窗口布局模式
    public void ShowAll()
    {
        Logger.Trace($"ShowAll: {ManagedWindows.Count} 个窗口", Cat);
        foreach (var w in ManagedWindows)
            NativeMethods.ShowWindow(w.Handle, NativeConstants.SW_SHOW);
    }

    // 交换

    /// <summary>
    /// 交换两个窗口在 <see cref="ManagedWindows"/> 列表中的位置，触发布局重排。
    /// </summary>
    /// <param name="hwndA">第一个窗口的句柄。</param>
    /// <param name="hwndB">第二个窗口的句柄。</param>
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

        // 确保 idxA < idxB，先移除高索引再移除低索引，避免索引错位
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

    // 事件处理

    /// <summary>
    /// 托管窗口被系统销毁时，从列表中移除并触发布局更新。
    /// </summary>
    /// <param name="hwnd">被销毁的窗口句柄。</param>
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

    /// <summary>
    /// 托管窗口标题发生变化时刷新 Title 属性，驱动侧边栏 UI 更新。
    /// </summary>
    /// <param name="hwnd">标题发生变化的窗口句柄。</param>
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

    // 释放所有资源，先释放所有窗口再取消事件订阅
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