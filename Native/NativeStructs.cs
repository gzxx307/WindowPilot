using System.Runtime.InteropServices;

namespace WindowPilot.Native;

// Win32 原生结构体定义，与 P/Invoke 调用的内存布局保持一致

[StructLayout(LayoutKind.Sequential)]
public struct RECT
{
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;

    // 派生属性，从坐标计算宽高
    public int Width  => Right - Left;
    public int Height => Bottom - Top;

    /// <summary>
    /// 使用四个坐标值构造矩形。
    /// </summary>
    /// <param name="left">左边界坐标。</param>
    /// <param name="top">上边界坐标。</param>
    /// <param name="right">右边界坐标。</param>
    /// <param name="bottom">下边界坐标。</param>
    public RECT(int left, int top, int right, int bottom)
    {
        Left = left; Top = top; Right = right; Bottom = bottom;
    }
}

[StructLayout(LayoutKind.Sequential)]
public struct POINT
{
    public int X;
    public int Y;
}

// WM_GETMINMAXINFO 消息附带的窗口最大最小尺寸信息
[StructLayout(LayoutKind.Sequential)]
public struct MINMAXINFO
{
    public POINT ptReserved;     // 系统保留
    public POINT ptMaxSize;      // 最大化时的窗口尺寸
    public POINT ptMaxPosition;  // 最大化时的窗口位置
    public POINT ptMinTrackSize; // 拖拽缩放的最小尺寸限制
    public POINT ptMaxTrackSize; // 拖拽缩放的最大尺寸限制
}

// 低级鼠标钩子回调中传入的鼠标事件数据
[StructLayout(LayoutKind.Sequential)]
public struct MSLLHOOKSTRUCT
{
    public POINT pt;          // 鼠标光标的屏幕坐标
    public uint mouseData;    // 附加数据，滚轮消息时表示滚动量
    public uint flags;
    public uint time;         // 事件发生的系统时间戳
    public IntPtr dwExtraInfo;
}

// DWM 缩略图的属性配置，通过 DwmUpdateThumbnailProperties 设置
[StructLayout(LayoutKind.Sequential)]
public struct DWM_THUMBNAIL_PROPERTIES
{
    public int  dwFlags;                // 指定哪些字段有效的标志位组合
    public RECT rcDestination;          // 缩略图在目标窗口中的绘制区域
    public RECT rcSource;               // 从源窗口截取的区域
    public byte opacity;                // 缩略图透明度，0 全透明，255 不透明
    [MarshalAs(UnmanagedType.Bool)]
    public bool fVisible;               // 是否显示缩略图
    [MarshalAs(UnmanagedType.Bool)]
    public bool fSourceClientAreaOnly;  // 为 true 时只渲染客户区，不含标题栏
}

[StructLayout(LayoutKind.Sequential)]
public struct SIZE
{
    public int cx; // 宽度
    public int cy; // 高度
}

// GetWindowPlacement / SetWindowPlacement 使用的窗口布局信息
[StructLayout(LayoutKind.Sequential)]
public struct WINDOWPLACEMENT
{
    public int   length;          // 结构体大小，调用前必须设置
    public int   flags;
    public int   showCmd;         // 窗口的显示状态，对应 SW_* 常量
    public POINT ptMinPosition;   // 最小化时的窗口左上角坐标
    public POINT ptMaxPosition;   // 最大化时的窗口左上角坐标
    public RECT  rcNormalPosition; // 正常状态下的窗口矩形
}