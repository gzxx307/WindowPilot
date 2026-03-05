using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using WindowPilot.Native;

namespace WindowPilot.Models;

/// <summary>
/// 表示一个被嵌入管理的窗口
/// </summary>
public class ManagedWindow : INotifyPropertyChanged
{
    public IntPtr Handle { get; }
    public uint ProcessId { get; }
    public string ProcessName { get; }
    public string ExecutablePath { get; }

    // ── 嵌入前保存的原始状态（释放时还原） ──
    public IntPtr OriginalParent { get; set; }
    public long OriginalStyle { get; set; }
    public long OriginalExStyle { get; set; }
    public RECT OriginalRect { get; set; }
    public bool WasMaximized { get; set; }

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
    public bool IsManaged
    {
        get => _isManaged;
        set { _isManaged = value; OnPropertyChanged(); }
    }

    private bool _isEmbedded;
    /// <summary>
    /// 是否已通过 SetParent 嵌入
    /// </summary>
    public bool IsEmbedded
    {
        get => _isEmbedded;
        set { _isEmbedded = value; OnPropertyChanged(); }
    }

    private int _slotIndex = -1;
    public int SlotIndex
    {
        get => _slotIndex;
        set { _slotIndex = value; OnPropertyChanged(); }
    }

    private bool _isActive;
    public bool IsActive
    {
        get => _isActive;
        set { _isActive = value; OnPropertyChanged(); }
    }

    public ManagedWindow(IntPtr handle)
    {
        Handle = handle;
        NativeMethods.GetWindowThreadProcessId(handle, out uint pid);
        ProcessId = pid;

        try
        {
            var proc = Process.GetProcessById((int)pid);
            ProcessName = proc.ProcessName;
        }
        catch
        {
            ProcessName = "Unknown";
        }

        ExecutablePath = GetProcessPath(pid);
        RefreshTitle();
        RefreshIcon();
    }

    /// <summary>
    /// 保存嵌入前的原始状态
    /// </summary>
    public void SaveOriginalState()
    {
        OriginalParent = NativeMethods.GetParent(Handle);
        OriginalStyle = NativeMethods.GetWindowLongSafe(Handle, NativeConstants.GWL_STYLE);
        OriginalExStyle = NativeMethods.GetWindowLongSafe(Handle, NativeConstants.GWL_EXSTYLE);
        NativeMethods.GetWindowRect(Handle, out RECT rect);
        OriginalRect = rect;
        WasMaximized = NativeMethods.IsZoomed(Handle);
    }

    public void RefreshTitle()
    {
        int length = NativeMethods.GetWindowTextLength(Handle);
        if (length > 0)
        {
            var sb = new StringBuilder(length + 1);
            NativeMethods.GetWindowText(Handle, sb, sb.Capacity);
            Title = sb.ToString();
        }
        else
        {
            Title = ProcessName;
        }
    }

    public void RefreshIcon()
    {
        try
        {
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
                Icon = Imaging.CreateBitmapSourceFromHIcon(
                    hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                Icon.Freeze();
            }
        }
        catch { }
    }

    public bool IsStillValid() => NativeMethods.IsWindowVisible(Handle);

    public bool CanWeControl()
    {
        try
        {
            long style = NativeMethods.GetWindowLongSafe(Handle, NativeConstants.GWL_STYLE);
            return style != 0 && NativeMethods.GetWindowRect(Handle, out _);
        }
        catch { return false; }
    }

    private static string GetProcessPath(uint pid)
    {
        IntPtr hProc = NativeMethods.OpenProcess(
            NativeConstants.PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
        if (hProc == IntPtr.Zero) return string.Empty;
        try
        {
            var sb = new StringBuilder(1024);
            uint size = 1024;
            return NativeMethods.QueryFullProcessImageName(hProc, 0, sb, ref size)
                ? sb.ToString() : string.Empty;
        }
        finally { NativeMethods.CloseHandle(hProc); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    public override string ToString() => $"[{Handle}] {Title} ({ProcessName})";
    public override int GetHashCode() => Handle.GetHashCode();
    public override bool Equals(object? obj) => obj is ManagedWindow other && Handle == other.Handle;
}
