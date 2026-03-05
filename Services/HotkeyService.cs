using System.Windows;
using System.Windows.Interop;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 全局热键管理
/// </summary>
public class HotkeyService : IDisposable
{
    private IntPtr _hwnd;
    private HwndSource? _source;
    private readonly Dictionary<int, Action> _hotkeyActions = new();
    private int _nextId = 9000;

    // 修饰键
    [Flags]
    public enum Modifiers : uint
    {
        None = 0x0000,
        Alt = 0x0001,
        Ctrl = 0x0002,
        Shift = 0x0004,
        Win = 0x0008,
    }

    /// <summary>
    /// 绑定到窗口以接收热键消息
    /// </summary>
    public void Attach(Window window)
    {
        var helper = new WindowInteropHelper(window);
        _hwnd = helper.Handle;
        _source = HwndSource.FromHwnd(_hwnd);
        _source?.AddHook(WndProc);
    }

    /// <summary>
    /// 注册一个全局热键
    /// </summary>
    public int Register(Modifiers modifiers, uint virtualKey, Action callback)
    {
        int id = _nextId++;
        if (NativeMethods.RegisterHotKey(_hwnd, id, (uint)modifiers, virtualKey))
        {
            _hotkeyActions[id] = callback;
            return id;
        }
        // 注册失败
        return -1;
    }

    /// <summary>
    /// 注销一个热键
    /// </summary>
    public void Unregister(int id)
    {
        NativeMethods.UnregisterHotKey(_hwnd, id);
        _hotkeyActions.Remove(id);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == NativeConstants.WM_HOTKEY)
        {
            int id = wParam.ToInt32();
            if (_hotkeyActions.TryGetValue(id, out var action))
            {
                action.Invoke();
                handled = true;
            }
        }
        return IntPtr.Zero;
    }

    public void Dispose()
    {
        foreach (var id in _hotkeyActions.Keys.ToList())
        {
            NativeMethods.UnregisterHotKey(_hwnd, id);
        }
        _hotkeyActions.Clear();
        _source?.RemoveHook(WndProc);
        GC.SuppressFinalize(this);
    }
}
