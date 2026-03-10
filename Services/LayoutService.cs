using WindowPilot.Diagnostics;
using WindowPilot.Models;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 布局引擎，在宿主窗口的客户区内计算并应用嵌入窗口的位置。
/// 所有坐标均为相对于宿主客户区左上角的像素坐标。
/// </summary>
public class LayoutService
{
    private const string Cat = "LayoutService";

    // 可用布局模式
    public enum LayoutMode
    {
        Stacked,    // 所有窗口叠放，一次只显示一个
        QuadSplit,  // 四等分
        LeftRight,  // 左右各半
        TopBottom,  // 上下各半
    }

    // 表示布局中单个位置区域的元数据
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
    private int        _applyCount;  // ApplyLayout 调用计数，用于日志追踪

    public LayoutMode CurrentMode
    {
        get => _currentMode;
        set
        {
            if (_currentMode == value) return;
            Logger.Info($"布局模式切换: {_currentMode} → {value}", Cat);
            _currentMode = value;
            // 模式变化后立即重算槽位
            RecalculateSlots();
        }
    }

    // 当前所有槽位的计算结果
    public List<LayoutSlot> Slots { get; private set; } = new();

    // 宿主区域矩形（相对于宿主客户区，像素单位）
    public RECT HostArea { get; private set; }

    // 相邻槽位之间的间隙（像素）
    public int Gap { get; set; } = 2;

    /// <summary>
    /// 构造布局服务，绑定到窗口管理器以读取托管窗口列表。
    /// </summary>
    /// <param name="windowManager">窗口管理服务实例，提供 ManagedWindows 列表和重排接口。</param>
    public LayoutService(WindowManagerService windowManager)
    {
        Logger.Debug("LayoutService 构造完成。", Cat);
        _windowManager = windowManager;
    }

    /// <summary>
    /// 更新宿主区域大小，触发槽位重算。在宿主 Panel 尺寸变化时调用。
    /// </summary>
    /// <param name="x">区域左上角相对于宿主客户区的 X 坐标（像素）。</param>
    /// <param name="y">区域左上角相对于宿主客户区的 Y 坐标（像素）。</param>
    /// <param name="width">区域宽度（像素）。</param>
    /// <param name="height">区域高度（像素）。</param>
    public void SetHostArea(int x, int y, int width, int height)
    {
        var oldArea = HostArea;
        HostArea = new RECT(x, y, x + width, y + height);

        Logger.Debug(
            $"SetHostArea: ({x},{y}) {width}×{height}  " +
            $"[之前: ({oldArea.Left},{oldArea.Top}) {oldArea.Width}×{oldArea.Height}]", Cat);

        RecalculateSlots();
    }

    // 根据当前模式和宿主区域重新计算所有槽位的坐标与尺寸
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
                // 堆叠模式只有一个全区域槽位
                Slots.Add(new LayoutSlot { Index = 0, Name = "堆叠",
                    X = x, Y = y, Width = w, Height = h });
                break;

            case LayoutMode.QuadSplit:
                // 四分区：宽高各减去一个间隙后对半分
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

    // 将托管窗口逐一移动到对应槽位，堆叠模式只显示活跃窗口
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
            // 所有窗口都占满同一个槽位
            foreach (var win in windows)
            {
                win.SlotIndex = 0;
                _windowManager.RepositionEmbedded(win.Handle, slot.X, slot.Y, slot.Width, slot.Height);
            }

            // 只显示激活窗口，其余隐藏
            var active = windows.FirstOrDefault(w => w.IsActive) ?? windows[0];
            Logger.Debug($"  堆叠模式 ShowOnly: \"{active.Title}\"", Cat);
            _windowManager.ShowOnly(active.Handle);
        }
        else
        {
            // 非堆叠模式先全部显示，再按索引分配槽位
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

            // 槽位数量不足时隐藏溢出窗口
            for (int i = Slots.Count; i < windows.Count; i++)
            {
                Logger.Debug(
                    $"  槽位不足，隐藏溢出窗口: \"{windows[i].Title}\"  index={i}", Cat);
                NativeMethods.ShowWindow(windows[i].Handle, 0);
            }
        }

        Logger.Trace($"ApplyLayout #{applyId} 完成。", Cat);
    }

    // 槽位查询

    /// <summary>
    /// 将屏幕坐标转换为槽位索引，用于拖拽落点命中检测。
    /// </summary>
    /// <param name="screenPt">屏幕物理像素坐标。</param>
    /// <param name="hostHwnd">宿主窗口句柄，用于坐标系转换。</param>
    /// <returns>命中的槽位索引，未命中任何槽位时返回 -1。</returns>
    public int GetSlotIndexAtScreenPoint(System.Windows.Point screenPt, IntPtr hostHwnd)
    {
        if (hostHwnd == IntPtr.Zero || Slots.Count == 0)
            return -1;

        // 将屏幕坐标转换为宿主客户区坐标
        var clientPt = new POINT { X = (int)screenPt.X, Y = (int)screenPt.Y };
        NativeMethods.ScreenToClient(hostHwnd, ref clientPt);

        int result = GetSlotIndexAtClientPoint(clientPt.X, clientPt.Y);
        Logger.Trace(
            $"GetSlotIndexAtScreenPoint  screen=({screenPt.X:F0},{screenPt.Y:F0})  " +
            $"client=({clientPt.X},{clientPt.Y})  → slot={result}", Cat);
        return result;
    }

    /// <summary>
    /// 在宿主客户区坐标中命中测试，返回覆盖该点的槽位索引。
    /// </summary>
    /// <param name="clientX">相对于宿主客户区的 X 坐标。</param>
    /// <param name="clientY">相对于宿主客户区的 Y 坐标。</param>
    /// <returns>命中的槽位索引，未命中时返回 -1。</returns>
    public int GetSlotIndexAtClientPoint(int clientX, int clientY)
    {
        for (int i = 0; i < Slots.Count; i++)
        {
            var slot = Slots[i];
            // 判断点是否在槽位矩形内（左闭右开区间）
            if (clientX >= slot.X && clientX < slot.X + slot.Width &&
                clientY >= slot.Y && clientY < slot.Y + slot.Height)
                return i;
        }
        return -1;
    }

    /// <summary>
    /// 获取指定槽位索引上当前分配的托管窗口。
    /// </summary>
    /// <param name="slotIndex">槽位索引。</param>
    /// <returns>该槽位的 <see cref="ManagedWindow"/>，索引越界时返回 null。</returns>
    public ManagedWindow? GetWindowAtSlotIndex(int slotIndex)
    {
        var windows = _windowManager.ManagedWindows;
        if (slotIndex < 0 || slotIndex >= windows.Count)
            return null;
        return windows[slotIndex];
    }

    /// <summary>
    /// 计算指定槽位在屏幕坐标系中的矩形，用于显示覆盖层或命中检测。
    /// </summary>
    /// <param name="slotIndex">目标槽位索引。</param>
    /// <param name="hostHwnd">宿主窗口句柄，用于客户区到屏幕坐标的转换。</param>
    /// <returns>槽位的屏幕坐标矩形，参数无效时返回 <see cref="System.Windows.Rect.Empty"/>。</returns>
    public System.Windows.Rect GetSlotScreenRect(int slotIndex, IntPtr hostHwnd)
    {
        if (slotIndex < 0 || slotIndex >= Slots.Count || hostHwnd == IntPtr.Zero)
            return System.Windows.Rect.Empty;

        var slot = Slots[slotIndex];
        // 分别转换左上角和右下角，得到完整的屏幕矩形
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

    // 切换控制

    /// <summary>
    /// 切换到指定托管窗口，堆叠模式下仅显示该窗口，其他模式下激活它。
    /// </summary>
    /// <param name="window">要切换到的目标 <see cref="ManagedWindow"/>。</param>
    public void SwitchToWindow(ManagedWindow window)
    {
        Logger.Debug($"SwitchToWindow: \"{window.Title}\"  mode={_currentMode}", Cat);
        if (_currentMode == LayoutMode.Stacked)
        {
            // 先将窗口移至槽位区域，再 ShowOnly 使其可见
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

    // 切换到列表中的下一个托管窗口（循环）
    public void SwitchToNext()
    {
        var windows = _windowManager.ManagedWindows;
        if (windows.Count <= 1) return;
        var active = windows.FirstOrDefault(w => w.IsActive);
        int idx    = active != null ? windows.IndexOf(active) : -1;
        int next   = (idx + 1) % windows.Count; // 末尾回到头部
        Logger.Debug($"SwitchToNext: [{idx}] → [{next}]  \"{windows[next].Title}\"", Cat);
        SwitchToWindow(windows[next]);
    }

    // 切换到列表中的上一个托管窗口（循环）
    public void SwitchToPrevious()
    {
        var windows = _windowManager.ManagedWindows;
        if (windows.Count <= 1) return;
        var active = windows.FirstOrDefault(w => w.IsActive);
        int idx    = active != null ? windows.IndexOf(active) : 0;
        int prev   = (idx - 1 + windows.Count) % windows.Count; // 头部回到末尾
        Logger.Debug($"SwitchToPrevious: [{idx}] → [{prev}]  \"{windows[prev].Title}\"", Cat);
        SwitchToWindow(windows[prev]);
    }
}