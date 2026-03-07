using System.Windows;
using System.Windows.Interop;
using WindowPilot.Diagnostics;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 全局热键管理
/// </summary>
public class HotkeyService : IDisposable
{
    private const string Cat = "HotkeyService";

    private IntPtr _hwnd;
    private HwndSource? _source;
    private readonly Dictionary<int, Action>  _hotkeyActions = new();
    private readonly Dictionary<int, string>  _hotkeyNames   = new(); // 仅用于日志
    private int _nextId = 9000;
    private int _invokeCount;

    // 修饰键
    [Flags]
    public enum Modifiers : uint
    {
        None  = 0x0000,
        Alt   = 0x0001,
        Ctrl  = 0x0002,
        Shift = 0x0004,
        Win   = 0x0008,
    }

    /// <summary>
    /// 绑定到窗口以接收热键消息
    /// </summary>
    public void Attach(Window window)
    {
        var helper = new WindowInteropHelper(window);
        _hwnd   = helper.Handle;
        _source = HwndSource.FromHwnd(_hwnd);
        _source?.AddHook(WndProc);
        Logger.Info($"HotkeyService 已绑定到窗口 hwnd=0x{_hwnd:X}", Cat);
    }

    /// <summary>
    /// 注册一个全局热键
    /// </summary>
    public int Register(Modifiers modifiers, uint virtualKey, Action callback,
        string description = "")
    {
        int id = _nextId++;
        string name = string.IsNullOrEmpty(description)
            ? $"Mod={modifiers} VK=0x{virtualKey:X2}"
            : description;

        Logger.Debug($"注册热键 id={id}  [{name}]", Cat);

        if (NativeMethods.RegisterHotKey(_hwnd, id, (uint)modifiers, virtualKey))
        {
            _hotkeyActions[id] = callback;
            _hotkeyNames[id]   = name;
            Logger.Info($"  ✓ 热键注册成功: [{name}] id={id}", Cat);
            return id;
        }

        Logger.Error($"  ✗ 热键注册失败: [{name}]  （可能与其他程序冲突）", Cat);
        return -1;
    }

    /// <summary>
    /// 注销一个热键
    /// </summary>
    public void Unregister(int id)
    {
        string name = _hotkeyNames.TryGetValue(id, out var n) ? n : id.ToString();
        Logger.Debug($"注销热键 id={id}  [{name}]", Cat);
        NativeMethods.UnregisterHotKey(_hwnd, id);
        _hotkeyActions.Remove(id);
        _hotkeyNames.Remove(id);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == NativeConstants.WM_HOTKEY)
        {
            int id = wParam.ToInt32();
            if (_hotkeyActions.TryGetValue(id, out var action))
            {
                int count = Interlocked.Increment(ref _invokeCount);
                string name = _hotkeyNames.TryGetValue(id, out var n) ? n : id.ToString();
                Logger.Info($"热键触发 #{count}  id={id}  [{name}]", Cat);
                action.Invoke();
                handled = true;
            }
            else
            {
                Logger.Warning($"WM_HOTKEY: 未找到 id={id} 对应的回调。", Cat);
            }
        }
        return IntPtr.Zero;
    }

    public void Dispose()
    {
        Logger.Debug($"HotkeyService.Dispose()  共注销 {_hotkeyActions.Count} 个热键", Cat);
        foreach (var id in _hotkeyActions.Keys.ToList())
        {
            string name = _hotkeyNames.TryGetValue(id, out var n) ? n : id.ToString();
            Logger.Debug($"  注销热键 id={id}  [{name}]", Cat);
            NativeMethods.UnregisterHotKey(_hwnd, id);
        }
        _hotkeyActions.Clear();
        _hotkeyNames.Clear();
        _source?.RemoveHook(WndProc);
        Logger.Debug($"统计 — 热键触发总次数: {_invokeCount}", Cat);
        GC.SuppressFinalize(this);
    }
}