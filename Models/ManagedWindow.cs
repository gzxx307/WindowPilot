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
/// 表示一个被嵌入管理的窗口
/// </summary>
public class ManagedWindow : INotifyPropertyChanged
{
    private const string Cat = "ManagedWindow";

    public IntPtr Handle       { get; }
    public uint   ProcessId    { get; }
    public string ProcessName  { get; }
    public string ExecutablePath { get; }

    // ── 嵌入前保存的原始状态（释放时还原） ──
    public IntPtr OriginalParent  { get; set; }
    public long   OriginalStyle   { get; set; }
    public long   OriginalExStyle { get; set; }
    public RECT   OriginalRect    { get; set; }
    public bool   WasMaximized    { get; set; }

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
        set
        {
            if (_isManaged == value) return;
            Logger.Trace($"[0x{Handle:X}] IsManaged: {_isManaged} → {value}", Cat);
            _isManaged = value;
            OnPropertyChanged();
        }
    }

    private bool _isEmbedded;
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

        Logger.Debug(
            $"ManagedWindow 构造完成: hwnd=0x{handle:X}  " +
            $"\"{_title}\"  进程={ProcessName}  PID={pid}  " +
            $"exe=\"{ExecutablePath}\"", Cat);
    }

    /// <summary>
    /// 保存嵌入前的原始状态
    /// </summary>
    public void SaveOriginalState()
    {
        OriginalParent  = NativeMethods.GetParent(Handle);
        OriginalStyle   = NativeMethods.GetWindowLongSafe(Handle, NativeConstants.GWL_STYLE);
        OriginalExStyle = NativeMethods.GetWindowLongSafe(Handle, NativeConstants.GWL_EXSTYLE);
        NativeMethods.GetWindowRect(Handle, out RECT rect);
        OriginalRect    = rect;
        WasMaximized    = NativeMethods.IsZoomed(Handle);

        Logger.Trace(
            $"SaveOriginalState [0x{Handle:X}]: " +
            $"parent=0x{OriginalParent:X}  style=0x{OriginalStyle:X}  " +
            $"exStyle=0x{OriginalExStyle:X}  " +
            $"rect=({rect.Left},{rect.Top},{rect.Width}×{rect.Height})  " +
            $"wasMaximized={WasMaximized}", Cat);
    }

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
            newTitle = ProcessName;
        }

        if (_title != newTitle)
        {
            Logger.Trace($"[0x{Handle:X}] RefreshTitle: \"{_title}\" → \"{newTitle}\"", Cat);
            Title = newTitle;
        }
    }

    public void RefreshIcon()
    {
        Logger.Trace($"[0x{Handle:X}] RefreshIcon 开始", Cat);
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

    public bool IsStillValid() => NativeMethods.IsWindowVisible(Handle);

    public bool CanWeControl()
    {
        try
        {
            long style = NativeMethods.GetWindowLongSafe(Handle, NativeConstants.GWL_STYLE);
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

    private static string GetProcessPath(uint pid)
    {
        IntPtr hProc = NativeMethods.OpenProcess(
            NativeConstants.PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
        if (hProc == IntPtr.Zero) return string.Empty;
        try
        {
            var sb   = new StringBuilder(1024);
            uint size = 1024;
            return NativeMethods.QueryFullProcessImageName(hProc, 0, sb, ref size)
                ? sb.ToString() : string.Empty;
        }
        finally { NativeMethods.CloseHandle(hProc); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    public override string ToString()    => $"[{Handle}] {Title} ({ProcessName})";
    public override int    GetHashCode() => Handle.GetHashCode();
    public override bool   Equals(object? obj) => obj is ManagedWindow other && Handle == other.Handle;
}