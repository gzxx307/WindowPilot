using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using WindowPilot.Diagnostics;
using WindowPilot.Native;

namespace WindowPilot.Models;

/// <summary>
/// 表示一个已被 WindowPilot 托管的外部窗口，封装其原始状态、当前属性和图标。
/// 实现 <see cref="INotifyPropertyChanged"/> 以支持 WPF 数据绑定。
/// </summary>
public class ManagedWindow : INotifyPropertyChanged
{
    private const string Cat = "ManagedWindow";

    // 只读标识属性，构造后不可变
    public IntPtr Handle        { get; }  // 窗口句柄，全局唯一标识
    public uint   ProcessId     { get; }  // 所属进程 ID
    public string ProcessName   { get; }  // 进程名（不含路径和扩展名）
    public string ExecutablePath { get; } // 可执行文件完整路径

    // 嵌入前保存的原始状态，释放时用于完整还原
    public IntPtr OriginalParent  { get; set; } // 原始父窗口句柄
    public long   OriginalStyle   { get; set; } // 原始 GWL_STYLE
    public long   OriginalExStyle { get; set; } // 原始 GWL_EXSTYLE
    public RECT   OriginalRect    { get; set; } // 原始屏幕坐标矩形
    public bool   WasMaximized    { get; set; } // 嵌入前是否处于最大化状态

    // 可绑定属性

    private string _title = string.Empty;
    public string Title
    {
        get => _title;
        set { _title = value; OnPropertyChanged(); }
    }

    private ImageSource? _icon;
    public ImageSource? Icon
    {
        get => _icon;
        set { _icon = value; OnPropertyChanged(); }
    }

    private bool _isManaged;
    // 是否在托管列表中
    public bool IsManaged
    {
        get => _isManaged;
        set
        {
            if (_isManaged == value) return;
            Logger.Trace($"[0x{Handle:X}] IsManaged: {_isManaged} → {value}", Cat);
            _isManaged = value;
            OnPropertyChanged();
        }
    }

    private bool _isEmbedded;
    // 是否已通过 SetParent 嵌入宿主，false 表示临时脱嵌或尚未嵌入
    public bool IsEmbedded
    {
        get => _isEmbedded;
        set
        {
            if (_isEmbedded == value) return;
            Logger.Trace($"[0x{Handle:X}] IsEmbedded: {_isEmbedded} → {value}", Cat);
            _isEmbedded = value;
            OnPropertyChanged();
        }
    }

    private int _slotIndex = -1;
    // 当前所在的布局槽位索引，-1 表示尚未分配
    public int SlotIndex
    {
        get => _slotIndex;
        set
        {
            if (_slotIndex == value) return;
            Logger.Trace($"[0x{Handle:X}] SlotIndex: {_slotIndex} → {value}", Cat);
            _slotIndex = value;
            OnPropertyChanged();
        }
    }

    private bool _isActive;
    // 是否是当前激活（前景）的托管窗口，用于侧边栏高亮显示
    public bool IsActive
    {
        get => _isActive;
        set
        {
            if (_isActive == value) return;
            Logger.Trace($"[0x{Handle:X}] IsActive: {_isActive} → {value}  \"{_title}\"", Cat);
            _isActive = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    /// 构造托管窗口对象，读取进程信息、标题和图标。
    /// </summary>
    /// <param name="handle">目标窗口的系统句柄。</param>
    public ManagedWindow(IntPtr handle)
    {
        Handle = handle;
        // 获取进程 ID，用于后续读取进程名和路径
        NativeMethods.GetWindowThreadProcessId(handle, out uint pid);
        ProcessId = pid;

        try
        {
            var proc    = Process.GetProcessById((int)pid);
            ProcessName = proc.ProcessName;
        }
        catch
        {
            // 进程可能在查询间隙退出，回退到 Unknown
            ProcessName = "Unknown";
        }

        ExecutablePath = GetProcessPath(pid);
        RefreshTitle();
        RefreshIcon();

        Logger.Debug(
            $"ManagedWindow 构造完成: hwnd=0x{handle:X}  " +
            $"\"{_title}\"  进程={ProcessName}  PID={pid}  " +
            $"exe=\"{ExecutablePath}\"", Cat);
    }

    // 在嵌入前调用，快照当前窗口的完整样式和位置，以备释放时还原
    public void SaveOriginalState()
    {
        OriginalParent  = NativeMethods.GetParent(Handle);
        OriginalStyle   = NativeMethods.GetWindowLongSafe(Handle, NativeConstants.GWL_STYLE);
        OriginalExStyle = NativeMethods.GetWindowLongSafe(Handle, NativeConstants.GWL_EXSTYLE);
        NativeMethods.GetWindowRect(Handle, out RECT rect);
        OriginalRect    = rect;
        // 记录最大化状态，释放后需要复原
        WasMaximized    = NativeMethods.IsZoomed(Handle);

        Logger.Trace(
            $"SaveOriginalState [0x{Handle:X}]: " +
            $"parent=0x{OriginalParent:X}  style=0x{OriginalStyle:X}  " +
            $"exStyle=0x{OriginalExStyle:X}  " +
            $"rect=({rect.Left},{rect.Top},{rect.Width}×{rect.Height})  " +
            $"wasMaximized={WasMaximized}", Cat);
    }

    // 重新从系统读取窗口标题，标题变化时通知绑定方
    public void RefreshTitle()
    {
        int length = NativeMethods.GetWindowTextLength(Handle);
        string newTitle;
        if (length > 0)
        {
            var sb = new StringBuilder(length + 1);
            NativeMethods.GetWindowText(Handle, sb, sb.Capacity);
            newTitle = sb.ToString();
        }
        else
        {
            // 窗口无标题时回退到进程名
            newTitle = ProcessName;
        }

        if (_title != newTitle)
        {
            Logger.Trace($"[0x{Handle:X}] RefreshTitle: \"{_title}\" → \"{newTitle}\"", Cat);
            Title = newTitle;
        }
    }

    // 按优先顺序尝试多种方式获取窗口图标，从小图标到大图标逐级回退
    public void RefreshIcon()
    {
        Logger.Trace($"[0x{Handle:X}] RefreshIcon 开始", Cat);
        try
        {
            // 依次尝试：任务栏小图标 → 标准小图标 → 窗口类图标 → 大图标
            IntPtr hIcon = NativeMethods.SendMessage(Handle, NativeMethods.WM_GETICON,
                (IntPtr)NativeMethods.ICON_SMALL2, IntPtr.Zero);
            if (hIcon == IntPtr.Zero)
                hIcon = NativeMethods.SendMessage(Handle, NativeMethods.WM_GETICON,
                    (IntPtr)NativeMethods.ICON_SMALL, IntPtr.Zero);
            if (hIcon == IntPtr.Zero)
                hIcon = NativeMethods.GetClassLongPtr(Handle, NativeMethods.GCLP_HICONSM);
            if (hIcon == IntPtr.Zero)
                hIcon = NativeMethods.SendMessage(Handle, NativeMethods.WM_GETICON,
                    (IntPtr)NativeMethods.ICON_BIG, IntPtr.Zero);

            if (hIcon != IntPtr.Zero)
            {
                // 转换为 WPF 可用的 BitmapSource 并冻结，允许跨线程访问
                Icon = Imaging.CreateBitmapSourceFromHIcon(
                    hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                Icon.Freeze();
                Logger.Trace($"[0x{Handle:X}] RefreshIcon 成功 hIcon=0x{hIcon:X}", Cat);
            }
            else
            {
                Logger.Trace($"[0x{Handle:X}] RefreshIcon: 未能获取图标。", Cat);
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"[0x{Handle:X}] RefreshIcon 异常: {ex.Message}", Cat);
        }
    }

    // 检查窗口是否仍然存在且可见
    public bool IsStillValid() => NativeMethods.IsWindowVisible(Handle);

    // 检查是否有足够权限操控该窗口
    public bool CanWeControl()
    {
        try
        {
            long style = NativeMethods.GetWindowLongSafe(Handle, NativeConstants.GWL_STYLE);
            // style 为 0 通常表示访问被拒绝（如管理员进程窗口）
            bool result = style != 0 && NativeMethods.GetWindowRect(Handle, out _);
            Logger.Trace($"[0x{Handle:X}] CanWeControl = {result}", Cat);
            return result;
        }
        catch (Exception ex)
        {
            Logger.Warning($"[0x{Handle:X}] CanWeControl 异常: {ex.Message}", Cat);
            return false;
        }
    }

    /// <summary>
    /// 通过 QueryFullProcessImageName 获取进程的完整可执行文件路径。
    /// </summary>
    /// <param name="pid">目标进程 ID。</param>
    /// <returns>完整路径字符串，获取失败时返回空字符串。</returns>
    private static string GetProcessPath(uint pid)
    {
        // PROCESS_QUERY_LIMITED_INFORMATION 是读取路径所需的最低权限
        IntPtr hProc = NativeMethods.OpenProcess(
            NativeConstants.PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
        if (hProc == IntPtr.Zero) return string.Empty;
        try
        {
            var sb    = new StringBuilder(1024);
            uint size = 1024;
            return NativeMethods.QueryFullProcessImageName(hProc, 0, sb, ref size)
                ? sb.ToString() : string.Empty;
        }
        finally
        {
            // 无论成功与否都必须关闭句柄，防止句柄泄漏
            NativeMethods.CloseHandle(hProc);
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>
    /// 触发属性变更通知，驱动 WPF 绑定更新。
    /// </summary>
    /// <param name="name">属性名，由编译器通过 <see cref="CallerMemberNameAttribute"/> 自动填充。</param>
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    public override string ToString()    => $"[{Handle}] {Title} ({ProcessName})";
    public override int    GetHashCode() => Handle.GetHashCode();

    /// <summary>
    /// 以句柄为唯一标识进行相等比较。
    /// </summary>
    /// <param name="obj">要比较的另一个对象。</param>
    public override bool   Equals(object? obj) => obj is ManagedWindow other && Handle == other.Handle;
}