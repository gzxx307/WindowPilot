namespace WindowPilot.Native;

/// <summary>
/// Win32 API 常量定义，集中管理所有 P/Invoke 调用中使用的数值。
/// </summary>
public static class NativeConstants
{
    // 窗口样式索引
    public const int GWL_STYLE   = -16; // 窗口基础样式
    public const int GWL_EXSTYLE = -20; // 窗口扩展样式

    // 基础窗口样式标志
    public const uint WS_VISIBLE    = 0x10000000;
    public const uint WS_MINIMIZE   = 0x20000000;
    public const uint WS_MAXIMIZE   = 0x01000000;
    public const uint WS_CAPTION    = 0x00C00000; // 标题栏（含边框）
    public const uint WS_THICKFRAME = 0x00040000; // 可调整大小的边框
    public const uint WS_POPUP      = 0x80000000;
    public const uint WS_CHILD      = 0x40000000;
    public const uint WS_DISABLED   = 0x08000000;

    // 扩展窗口样式标志
    public const uint WS_EX_TOOLWINDOW = 0x00000080; // 工具窗口，不显示在任务栏
    public const uint WS_EX_APPWINDOW  = 0x00040000; // 强制显示在任务栏
    public const uint WS_EX_NOACTIVATE = 0x08000000; // 激活时不获取焦点
    public const uint WS_EX_TOPMOST    = 0x00000008;
    public const uint WS_EX_TRANSPARENT = 0x00000020;
    public const uint WS_EX_LAYERED    = 0x00080000;

    // 嵌入窗口时需要移除的标准窗口样式组合
    public const uint WS_OVERLAPPEDWINDOW = WS_CAPTION | WS_THICKFRAME
        | 0x00020000 | 0x00010000 | 0x00010000;

    public const uint WS_CLIPCHILDREN = 0x02000000;
    public const uint WS_CLIPSIBLINGS = 0x04000000;
    public const uint WS_BORDER       = 0x00800000;
    public const uint WS_DLGFRAME     = 0x00400000;
    public const uint WS_SYSMENU      = 0x00080000;
    public const uint WS_MINIMIZEBOX  = 0x00020000;
    public const uint WS_MAXIMIZEBOX  = 0x00010000;

    // SetWindowPos 位置标志
    public const uint SWP_NOSIZE       = 0x0001; // 不改变尺寸
    public const uint SWP_NOMOVE       = 0x0002; // 不改变位置
    public const uint SWP_NOZORDER     = 0x0004; // 不改变 Z 序
    public const uint SWP_NOACTIVATE   = 0x0010; // 不激活窗口
    public const uint SWP_SHOWWINDOW   = 0x0040;
    public const uint SWP_FRAMECHANGED = 0x0020; // 强制重新计算非客户区，样式变更后必须使用

    // ShowWindow 命令
    public const int SW_RESTORE  = 9;
    public const int SW_SHOW     = 5;
    public const int SW_MINIMIZE = 6;
    public const int SW_HIDE     = 0;

    // WinEvent 钩子标志
    public const uint WINEVENT_OUTOFCONTEXT   = 0x0000; // 事件在调用线程上下文中处理
    public const uint WINEVENT_SKIPOWNPROCESS = 0x0002; // 忽略本进程自身产生的事件

    // WinEvent 事件类型
    public const uint EVENT_SYSTEM_MOVESIZESTART  = 0x000A; // 窗口开始移动或调整大小
    public const uint EVENT_SYSTEM_MOVESIZEEND    = 0x000B; // 窗口完成移动或调整大小
    public const uint EVENT_SYSTEM_FOREGROUND     = 0x0003; // 前景窗口发生变化
    public const uint EVENT_SYSTEM_MINIMIZESTART  = 0x0016;
    public const uint EVENT_SYSTEM_MINIMIZEEND    = 0x0017;
    public const uint EVENT_OBJECT_DESTROY        = 0x8001; // 窗口对象被销毁
    public const uint EVENT_OBJECT_LOCATIONCHANGE = 0x800B; // 窗口位置或大小发生变化
    public const uint EVENT_OBJECT_NAMECHANGE     = 0x800C; // 窗口标题发生变化

    public const int OBJID_WINDOW = 0; // 事件目标为窗口本身

    // 低级鼠标钩子类型与消息
    public const int WH_MOUSE_LL  = 14;
    public const int WM_MOUSEMOVE    = 0x0200;
    public const int WM_LBUTTONUP   = 0x0202;
    public const int WM_LBUTTONDOWN = 0x0201;

    // 窗口消息
    public const int WM_GETMINMAXINFO = 0x0024;
    public const int WM_HOTKEY        = 0x0312;
    public const int WM_SYSCOMMAND    = 0x0112;
    public const int WM_KEYDOWN       = 0x0100;

    // SC_MOVE 用于编程式发起窗口拖拽
    public const int SC_MOVE = 0xF010;

    // 虚拟键码
    public const int VK_MENU  = 0x12; // Alt 键
    public const int VK_LEFT  = 0x25;
    public const int VK_UP    = 0x26;
    public const int VK_RIGHT = 0x27;
    public const int VK_DOWN  = 0x28;

    // 非客户区命中测试结果
    public const int HTCAPTION = 2; // 命中标题栏区域

    // DWM 缩略图属性标志
    public const int DWM_TNP_RECTDESTINATION      = 0x00000001;
    public const int DWM_TNP_RECTSOURCE           = 0x00000002;
    public const int DWM_TNP_OPACITY              = 0x00000004;
    public const int DWM_TNP_VISIBLE              = 0x00000008;
    public const int DWM_TNP_SOURCECLIENTAREAONLY = 0x00000010;

    // 进程访问权限
    public const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;

    // HWND 特殊值
    public static readonly IntPtr HWND_TOP      = IntPtr.Zero;
    public static readonly IntPtr HWND_TOPMOST  = new(-1);
    public static readonly IntPtr HWND_NOTOPMOST = new(-2);
}