namespace WindowPilot.Native;

/// <summary>
/// Win32 常量定义
/// </summary>
public static class NativeConstants
{
    // ── Styles ──
    public const int GWL_STYLE = -16;
    public const int GWL_EXSTYLE = -20;

    public const uint WS_VISIBLE     = 0x10000000;
    public const uint WS_MINIMIZE    = 0x20000000;
    public const uint WS_MAXIMIZE    = 0x01000000;
    public const uint WS_CAPTION     = 0x00C00000;
    public const uint WS_THICKFRAME  = 0x00040000;
    public const uint WS_POPUP       = 0x80000000;
    public const uint WS_CHILD       = 0x40000000;
    public const uint WS_DISABLED    = 0x08000000;

    public const uint WS_EX_TOOLWINDOW   = 0x00000080;
    public const uint WS_EX_APPWINDOW    = 0x00040000;
    public const uint WS_EX_NOACTIVATE   = 0x08000000;
    public const uint WS_EX_TOPMOST      = 0x00000008;
    public const uint WS_EX_TRANSPARENT  = 0x00000020;
    public const uint WS_EX_LAYERED      = 0x00080000;

    // 用于嵌入时修改窗口样式
    public const uint WS_OVERLAPPEDWINDOW = WS_CAPTION | WS_THICKFRAME
        | 0x00020000 /*SYSMENU*/ | 0x00010000 /*MINIMIZEBOX*/ | 0x00010000 /*MAXIMIZEBOX*/;
    public const uint WS_CLIPCHILDREN  = 0x02000000;
    public const uint WS_CLIPSIBLINGS  = 0x04000000;
    public const uint WS_BORDER        = 0x00800000;
    public const uint WS_DLGFRAME      = 0x00400000;
    public const uint WS_SYSMENU       = 0x00080000;
    public const uint WS_MINIMIZEBOX   = 0x00020000;
    public const uint WS_MAXIMIZEBOX   = 0x00010000;

    // ── SetWindowPos Flags ──
    public const uint SWP_NOSIZE       = 0x0001;
    public const uint SWP_NOMOVE       = 0x0002;
    public const uint SWP_NOZORDER     = 0x0004;
    public const uint SWP_NOACTIVATE   = 0x0010;
    public const uint SWP_SHOWWINDOW   = 0x0040;
    public const uint SWP_FRAMECHANGED = 0x0020;

    // ── ShowWindow Commands ──
    public const int SW_RESTORE  = 9;
    public const int SW_SHOW     = 5;
    public const int SW_MINIMIZE = 6;

    // ── WinEvent Constants ──
    public const uint WINEVENT_OUTOFCONTEXT   = 0x0000;
    public const uint WINEVENT_SKIPOWNPROCESS = 0x0002;
    public const uint EVENT_SYSTEM_MOVESIZESTART = 0x000A;
    public const uint EVENT_SYSTEM_MOVESIZEEND   = 0x000B;
    public const uint EVENT_SYSTEM_FOREGROUND    = 0x0003;
    public const uint EVENT_SYSTEM_MINIMIZESTART = 0x0016;
    public const uint EVENT_SYSTEM_MINIMIZEEND   = 0x0017;
    public const uint EVENT_OBJECT_DESTROY       = 0x8001;
    public const uint EVENT_OBJECT_LOCATIONCHANGE = 0x800B;
    public const uint EVENT_OBJECT_NAMECHANGE     = 0x800C;

    public const int OBJID_WINDOW = 0;

    // ── Mouse Hook ──
    public const int WH_MOUSE_LL = 14;
    public const int WM_MOUSEMOVE   = 0x0200;
    public const int WM_LBUTTONUP   = 0x0202;
    public const int WM_LBUTTONDOWN = 0x0201;

    // ── Window Messages ──
    public const int WM_GETMINMAXINFO = 0x0024;
    public const int WM_HOTKEY        = 0x0312;

    // ── DWM Thumbnail ──
    public const int DWM_TNP_RECTDESTINATION = 0x00000001;
    public const int DWM_TNP_RECTSOURCE      = 0x00000002;
    public const int DWM_TNP_OPACITY         = 0x00000004;
    public const int DWM_TNP_VISIBLE         = 0x00000008;
    public const int DWM_TNP_SOURCECLIENTAREAONLY = 0x00000010;

    // ── Process Access ──
    public const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;

    // ── HWND Special Values ──
    public static readonly IntPtr HWND_TOP    = IntPtr.Zero;
    public static readonly IntPtr HWND_TOPMOST = new(-1);
    public static readonly IntPtr HWND_NOTOPMOST = new(-2);
}
