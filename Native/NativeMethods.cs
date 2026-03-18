using System.Runtime.InteropServices;
using System.Text;

namespace WindowPilot.Native;

/// <summary>
/// Win32 API P/Invoke 声明及安全封装。所有直接调用系统 API 的入口均在此处集中定义。
/// </summary>
public static class NativeMethods
{
    // 委托类型声明
    public delegate bool    EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    public delegate IntPtr  LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
    public delegate void    WinEventDelegate(
        IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

    // 窗口枚举
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsZoomed(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetParent(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetShellWindow();

    [DllImport("user32.dll")]
    public static extern IntPtr GetDesktopWindow();

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    // 窗口样式读写
    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    // 64 位版本，使用 IntPtr 支持 64 位指针宽度
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
    public static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
    public static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    /// <summary>
    /// 读取窗口的 Long 值，自动适配 32/64 位进程。
    /// </summary>
    /// <param name="hWnd">目标窗口句柄。</param>
    /// <param name="nIndex">要读取的值的索引，使用 <see cref="NativeConstants"/> 中的 GWL_* 常量。</param>
    /// <returns>读取到的样式值，失败时返回 0。</returns>
    public static long GetWindowLongSafe(IntPtr hWnd, int nIndex)
    {
        // 根据指针大小选择对应 API
        if (IntPtr.Size == 8)
            return GetWindowLongPtr64(hWnd, nIndex).ToInt64();
        return GetWindowLong(hWnd, nIndex);
    }

    // 窗口位置与尺寸
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

    // 坐标转换
    /// <summary>
    /// 将屏幕坐标转换为指定窗口客户区的本地坐标。
    /// </summary>
    /// <param name="hWnd">参考窗口句柄，定义客户区坐标系原点。</param>
    /// <param name="lpPoint">输入屏幕坐标，转换后的结果也写入此结构。</param>
    [DllImport("user32.dll")]
    public static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

    /// <summary>
    /// 将指定窗口客户区坐标转换为屏幕坐标。
    /// </summary>
    /// <param name="hWnd">参考窗口句柄，定义客户区坐标系原点。</param>
    /// <param name="lpPoint">输入客户区坐标，转换后的结果也写入此结构。</param>
    [DllImport("user32.dll")]
    public static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

    // 消息发送
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, ref MINMAXINFO lParam);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    // 进程信息
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("kernel32.dll")]
    public static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

    [DllImport("kernel32.dll")]
    public static extern bool CloseHandle(IntPtr hObject);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    public static extern bool QueryFullProcessImageName(IntPtr hProcess, uint dwFlags,
        StringBuilder lpExeName, ref uint lpdwSize);

    // 钩子安装与卸载
    [DllImport("user32.dll")]
    public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn,
        IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    public static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr GetModuleHandle(string? lpModuleName);

    // WinEvent 钩子
    [DllImport("user32.dll")]
    public static extern IntPtr SetWinEventHook(
        uint eventMin, uint eventMax, IntPtr hmodWinEventProc,
        WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

    // 获取指定父窗口的 Z-order 最顶层子窗口（截图前保存原始顶层，截图后还原用）
    [DllImport("user32.dll")]
    public static extern IntPtr GetTopWindow(IntPtr hWnd);

    
    [DllImport("user32.dll")]
    public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

    // 全局热键注册
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    // DWM 缩略图
    [DllImport("dwmapi.dll")]
    public static extern int DwmRegisterThumbnail(IntPtr hwndDestination, IntPtr hwndSource, out IntPtr phThumbnailId);

    [DllImport("dwmapi.dll")]
    public static extern int DwmUnregisterThumbnail(IntPtr hThumbnailId);

    [DllImport("dwmapi.dll")]
    public static extern int DwmUpdateThumbnailProperties(IntPtr hThumbnailId, ref DWM_THUMBNAIL_PROPERTIES ptnProperties);

    [DllImport("dwmapi.dll")]
    public static extern int DwmQueryThumbnailSourceSize(IntPtr hThumbnailId, out SIZE pSize);

    // 显示器信息
    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

    public const uint MONITOR_DEFAULTTONEAREST = 2;

    // 鼠标光标
    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);

    // 点命中测试
    [DllImport("user32.dll")]
    public static extern IntPtr WindowFromPoint(POINT point);

    [DllImport("user32.dll")]
    public static extern IntPtr ChildWindowFromPointEx(IntPtr hWndParent, POINT pt, uint uFlags);

    // 父窗口重设（窗口嵌入核心 API）
    [DllImport("user32.dll")]
    public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
 

    /// <summary>
    /// 写入窗口的 Long 值，自动适配 32/64 位进程。
    /// </summary>
    /// <param name="hWnd">目标窗口句柄。</param>
    /// <param name="nIndex">要写入的值的索引，使用 <see cref="NativeConstants"/> 中的 GWL_* 常量。</param>
    /// <param name="dwNewLong">要设置的新值。</param>
    public static void SetWindowLongPtrSafe(IntPtr hWnd, int nIndex, long dwNewLong)
    {
        if (IntPtr.Size == 8)
            SetWindowLongPtr64(hWnd, nIndex, new IntPtr(dwNewLong));
        else
            SetWindowLong(hWnd, nIndex, (int)dwNewLong);
    }

    // DPI 查询
    [DllImport("user32.dll")]
    public static extern uint GetDpiForWindow(IntPtr hwnd);

    // 键盘状态
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);

    // 杂项
    [DllImport("user32.dll")]
    public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    [DllImport("kernel32.dll")]
    public static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    public static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, bool bErase);

    [DllImport("user32.dll")]
    public static extern bool UpdateWindow(IntPtr hWnd);

    // 图标提取
    [DllImport("user32.dll")]
    public static extern IntPtr GetClassLongPtr(IntPtr hWnd, int nIndex);

    public const int GCLP_HICON   = -14; // 类的大图标
    public const int GCLP_HICONSM = -34; // 类的小图标

    public const uint WM_GETICON  = 0x007F;
    public const int  ICON_SMALL  = 0;
    public const int  ICON_BIG    = 1;
    public const int  ICON_SMALL2 = 2; // 任务栏使用的内部小图标

    [DllImport("user32.dll")]
    public static extern bool DestroyIcon(IntPtr hIcon);

    // ReleaseCapture 用于配合 SC_MOVE 编程式发起窗口拖拽
    [DllImport("user32.dll")]
    public static extern bool ReleaseCapture();

    // ── GDI 截图（PrintWindow 回退方案）──

    /// <summary>
    /// 获取指定窗口（或屏幕）的设备上下文句柄。
    /// 传 IntPtr.Zero 获取整个屏幕的 DC。
    /// </summary>
    [DllImport("user32.dll")]
    public static extern IntPtr GetDC(IntPtr hwnd);

    /// <summary>
    /// 释放 GetDC 获取的 DC。
    /// </summary>
    [DllImport("user32.dll")]
    public static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

    /// <summary>
    /// 创建与指定 DC 兼容的内存 DC，用于离屏绘制。
    /// </summary>
    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    /// <summary>
    /// 创建与指定 DC 兼容的位图，尺寸由 nWidth/nHeight 指定（物理像素）。
    /// </summary>
    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

    /// <summary>
    /// 将 GDI 对象（如 HBITMAP）选入 DC，返回被替换的旧对象句柄。
    /// </summary>
    [DllImport("gdi32.dll")]
    public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

    /// <summary>
    /// 删除 GDI 对象（HBITMAP、HPEN、HBRUSH 等），释放其占用的资源。
    /// </summary>
    [DllImport("gdi32.dll")]
    public static extern bool DeleteObject(IntPtr hObject);

    /// <summary>
    /// 删除由 CreateCompatibleDC 创建的内存 DC，释放资源。
    /// </summary>
    [DllImport("gdi32.dll")]
    public static extern bool DeleteDC(IntPtr hdc);

    /// <summary>
    /// 将窗口内容绘制到指定 DC。
    /// <paramref name="nFlags"/> 传 <see cref="PW_RENDERFULLCONTENT"/> 可捕获 DirectX/GPU 加速内容。
    /// </summary>
    [DllImport("user32.dll")]
    public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

    /// <summary>PrintWindow 标志：渲染完整客户区内容，包含 DirectX/GPU 合成层。</summary>
    public const uint PW_RENDERFULLCONTENT = 0x00000002;
}