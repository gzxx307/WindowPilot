using WindowPilot.Diagnostics;
using WindowPilot.Models;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 布局引擎：在宿主窗口内部计算嵌入窗口的位置
/// 所有坐标都是相对于宿主客户区的像素坐标
/// </summary>
public class LayoutService
{
    private const string Cat = "LayoutService";

    public enum LayoutMode
    {
        Stacked,
        QuadSplit,
        LeftRight,
        TopBottom,
    }

    public class LayoutSlot
    {
        public int    Index  { get; set; }
        public string Name   { get; set; } = "";
        public int    X      { get; set; }
        public int    Y      { get; set; }
        public int    Width  { get; set; }
        public int    Height { get; set; }
    }

    private readonly WindowManagerService _windowManager;
    private LayoutMode _currentMode = LayoutMode.Stacked;
    private int _applyCount;

    public LayoutMode CurrentMode
    {
        get => _currentMode;
        set
        {
            if (_currentMode == value) return;
            Logger.Info($"布局模式切换: {_currentMode} → {value}", Cat);
            _currentMode = value;
            RecalculateSlots();
        }
    }

    public List<LayoutSlot> Slots { get; private set; } = new();

    /// <summary>宿主区域的像素矩形（相对于宿主客户区左上角）</summary>
    public RECT HostArea { get; private set; }

    /// <summary>窗口间距（像素）</summary>
    public int Gap { get; set; } = 2;

    public LayoutService(WindowManagerService windowManager)
    {
        Logger.Debug("LayoutService 构造完成。", Cat);
        _windowManager = windowManager;
    }

    /// <summary>
    /// 设置宿主区域大小（像素，相对于宿主）
    /// </summary>
    public void SetHostArea(int x, int y, int width, int height)
    {
        var oldArea = HostArea;
        HostArea = new RECT(x, y, x + width, y + height);

        Logger.Debug(
            $"SetHostArea: ({x},{y}) {width}×{height}  " +
            $"[之前: ({oldArea.Left},{oldArea.Top}) {oldArea.Width}×{oldArea.Height}]", Cat);

        RecalculateSlots();
    }

    /// <summary>
    /// 重新计算布局槽位
    /// </summary>
    public void RecalculateSlots()
    {
        int x = HostArea.Left;
        int y = HostArea.Top;
        int w = HostArea.Width;
        int h = HostArea.Height;
        int g = Gap;

        if (w <= 0 || h <= 0)
        {
            Logger.Warning($"RecalculateSlots: 宿主区域无效 ({w}×{h})，跳过。", Cat);
            return;
        }

        Slots.Clear();

        switch (_currentMode)
        {
            case LayoutMode.Stacked:
                Slots.Add(new LayoutSlot { Index = 0, Name = "堆叠",
                    X = x, Y = y, Width = w, Height = h });
                break;

            case LayoutMode.QuadSplit:
                int halfW = (w - g) / 2;
                int halfH = (h - g) / 2;
                Slots.Add(new LayoutSlot { Index = 0, Name = "左上",
                    X = x, Y = y, Width = halfW, Height = halfH });
                Slots.Add(new LayoutSlot { Index = 1, Name = "右上",
                    X = x + halfW + g, Y = y, Width = w - halfW - g, Height = halfH });
                Slots.Add(new LayoutSlot { Index = 2, Name = "左下",
                    X = x, Y = y + halfH + g, Width = halfW, Height = h - halfH - g });
                Slots.Add(new LayoutSlot { Index = 3, Name = "右下",
                    X = x + halfW + g, Y = y + halfH + g, Width = w - halfW - g, Height = h - halfH - g });
                break;

            case LayoutMode.LeftRight:
                int lrHalf = (w - g) / 2;
                Slots.Add(new LayoutSlot { Index = 0, Name = "左",
                    X = x, Y = y, Width = lrHalf, Height = h });
                Slots.Add(new LayoutSlot { Index = 1, Name = "右",
                    X = x + lrHalf + g, Y = y, Width = w - lrHalf - g, Height = h });
                break;

            case LayoutMode.TopBottom:
                int tbHalf = (h - g) / 2;
                Slots.Add(new LayoutSlot { Index = 0, Name = "上",
                    X = x, Y = y, Width = w, Height = tbHalf });
                Slots.Add(new LayoutSlot { Index = 1, Name = "下",
                    X = x, Y = y + tbHalf + g, Width = w, Height = h - tbHalf - g });
                break;
        }

        Logger.Debug(
            $"RecalculateSlots [{_currentMode}]: 生成 {Slots.Count} 个槽位", Cat);
        foreach (var slot in Slots)
            Logger.Trace(
                $"  Slot[{slot.Index}] {slot.Name,-4}  " +
                $"({slot.X},{slot.Y}) {slot.Width}×{slot.Height}", Cat);
    }

    /// <summary>
    /// 应用布局：将嵌入窗口移动到对应槽位
    /// </summary>
    public void ApplyLayout()
    {
        var windows = _windowManager.ManagedWindows.ToList();
        if (windows.Count == 0 || Slots.Count == 0)
        {
            Logger.Trace(
                $"ApplyLayout 跳过: windows={windows.Count}  slots={Slots.Count}", Cat);
            return;
        }

        int applyId = Interlocked.Increment(ref _applyCount);
        Logger.Debug(
            $"ApplyLayout #{applyId}  mode={_currentMode}  " +
            $"windows={windows.Count}  slots={Slots.Count}", Cat);

        if (_currentMode == LayoutMode.Stacked)
        {
            var slot = Slots[0];
            foreach (var win in windows)
            {
                win.SlotIndex = 0;
                _windowManager.RepositionEmbedded(win.Handle, slot.X, slot.Y, slot.Width, slot.Height);
            }

            var active = windows.FirstOrDefault(w => w.IsActive) ?? windows[0];
            Logger.Debug($"  堆叠模式 ShowOnly: \"{active.Title}\"", Cat);
            _windowManager.ShowOnly(active.Handle);
        }
        else
        {
            _windowManager.ShowAll();
            for (int i = 0; i < windows.Count && i < Slots.Count; i++)
            {
                var win  = windows[i];
                var slot = Slots[i];
                win.SlotIndex = slot.Index;
                Logger.Trace(
                    $"  [{slot.Name}] \"{win.Title}\" → ({slot.X},{slot.Y}) {slot.Width}×{slot.Height}", Cat);
                _windowManager.RepositionEmbedded(win.Handle, slot.X, slot.Y, slot.Width, slot.Height);
            }

            for (int i = Slots.Count; i < windows.Count; i++)
            {
                Logger.Debug(
                    $"  槽位不足，隐藏溢出窗口: \"{windows[i].Title}\"  index={i}", Cat);
                NativeMethods.ShowWindow(windows[i].Handle, 0);
            }
        }

        Logger.Trace($"ApplyLayout #{applyId} 完成。", Cat);
    }

    // ── 槽位查询 ──────────────────────────────────────────

    public int GetSlotIndexAtScreenPoint(System.Windows.Point screenPt, IntPtr hostHwnd)
    {
        if (hostHwnd == IntPtr.Zero || Slots.Count == 0)
            return -1;

        var clientPt = new POINT { X = (int)screenPt.X, Y = (int)screenPt.Y };
        NativeMethods.ScreenToClient(hostHwnd, ref clientPt);

        int result = GetSlotIndexAtClientPoint(clientPt.X, clientPt.Y);
        Logger.Trace(
            $"GetSlotIndexAtScreenPoint  screen=({screenPt.X:F0},{screenPt.Y:F0})  " +
            $"client=({clientPt.X},{clientPt.Y})  → slot={result}", Cat);
        return result;
    }

    public int GetSlotIndexAtClientPoint(int clientX, int clientY)
    {
        for (int i = 0; i < Slots.Count; i++)
        {
            var slot = Slots[i];
            if (clientX >= slot.X && clientX < slot.X + slot.Width &&
                clientY >= slot.Y && clientY < slot.Y + slot.Height)
                return i;
        }
        return -1;
    }

    public ManagedWindow? GetWindowAtSlotIndex(int slotIndex)
    {
        var windows = _windowManager.ManagedWindows;
        if (slotIndex < 0 || slotIndex >= windows.Count)
            return null;
        return windows[slotIndex];
    }

    public System.Windows.Rect GetSlotScreenRect(int slotIndex, IntPtr hostHwnd)
    {
        if (slotIndex < 0 || slotIndex >= Slots.Count || hostHwnd == IntPtr.Zero)
            return System.Windows.Rect.Empty;

        var slot = Slots[slotIndex];
        var tlPt = new POINT { X = slot.X, Y = slot.Y };
        NativeMethods.ClientToScreen(hostHwnd, ref tlPt);
        var brPt = new POINT { X = slot.X + slot.Width, Y = slot.Y + slot.Height };
        NativeMethods.ClientToScreen(hostHwnd, ref brPt);

        var result = new System.Windows.Rect(
            tlPt.X, tlPt.Y, brPt.X - tlPt.X, brPt.Y - tlPt.Y);

        Logger.Trace(
            $"GetSlotScreenRect slot={slotIndex}  " +
            $"→ ({result.Left:F0},{result.Top:F0}) {result.Width:F0}×{result.Height:F0}", Cat);
        return result;
    }

    // ── 切换 ──────────────────────────────────────────────

    public void SwitchToWindow(ManagedWindow window)
    {
        Logger.Debug($"SwitchToWindow: \"{window.Title}\"  mode={_currentMode}", Cat);
        if (_currentMode == LayoutMode.Stacked)
        {
            var slot = Slots.FirstOrDefault();
            if (slot != null)
                _windowManager.RepositionEmbedded(window.Handle, slot.X, slot.Y, slot.Width, slot.Height);
            _windowManager.ShowOnly(window.Handle);
        }
        else
        {
            _windowManager.ActivateWindow(window.Handle);
        }
    }

    public void SwitchToNext()
    {
        var windows = _windowManager.ManagedWindows;
        if (windows.Count <= 1) return;
        var active = windows.FirstOrDefault(w => w.IsActive);
        int idx    = active != null ? windows.IndexOf(active) : -1;
        int next   = (idx + 1) % windows.Count;
        Logger.Debug($"SwitchToNext: [{idx}] → [{next}]  \"{windows[next].Title}\"", Cat);
        SwitchToWindow(windows[next]);
    }

    public void SwitchToPrevious()
    {
        var windows = _windowManager.ManagedWindows;
        if (windows.Count <= 1) return;
        var active = windows.FirstOrDefault(w => w.IsActive);
        int idx    = active != null ? windows.IndexOf(active) : 0;
        int prev   = (idx - 1 + windows.Count) % windows.Count;
        Logger.Debug($"SwitchToPrevious: [{idx}] → [{prev}]  \"{windows[prev].Title}\"", Cat);
        SwitchToWindow(windows[prev]);
    }
}