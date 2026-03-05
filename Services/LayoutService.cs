using WindowPilot.Models;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 布局引擎：在宿主窗口内部计算嵌入窗口的位置
/// 所有坐标都是相对于宿主客户区的像素坐标
/// </summary>
public class LayoutService
{
    public enum LayoutMode
    {
        Stacked,      //堆叠
        QuadSplit,    //四分区
        LeftRight,    //左右分屏
        TopBottom,    //上下分屏
    }

    public class LayoutSlot
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }

    private readonly WindowManagerService _windowManager;
    private LayoutMode _currentMode = LayoutMode.Stacked;

    public LayoutMode CurrentMode
    {
        get => _currentMode;
        set { _currentMode = value; RecalculateSlots(); }
    }

    public List<LayoutSlot> Slots { get; private set; } = new();

    /// <summary>
    /// 宿主区域的像素矩形（相对于宿主客户区左上角）
    /// </summary>
    public RECT HostArea { get; private set; }

    /// <summary>
    /// 窗口间距（像素）
    /// </summary>
    public int Gap { get; set; } = 2;

    public LayoutService(WindowManagerService windowManager)
    {
        _windowManager = windowManager;
    }

    /// <summary>
    /// 设置宿主区域大小（像素，相对于宿主）
    /// 每次宿主窗口大小变化时调用
    /// </summary>
    public void SetHostArea(int x, int y, int width, int height)
    {
        HostArea = new RECT(x, y, x + width, y + height);
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

        if (w <= 0 || h <= 0) return;

        Slots.Clear();

        switch (_currentMode)
        {
            case LayoutMode.Stacked:
                Slots.Add(new LayoutSlot { Index = 0, Name = "堆叠", X = x, Y = y, Width = w, Height = h });
                break;

            case LayoutMode.QuadSplit:
                int halfW = (w - g) / 2;
                int halfH = (h - g) / 2;
                Slots.Add(new LayoutSlot { Index = 0, Name = "左上", X = x, Y = y, Width = halfW, Height = halfH });
                Slots.Add(new LayoutSlot { Index = 1, Name = "右上", X = x + halfW + g, Y = y, Width = w - halfW - g, Height = halfH });
                Slots.Add(new LayoutSlot { Index = 2, Name = "左下", X = x, Y = y + halfH + g, Width = halfW, Height = h - halfH - g });
                Slots.Add(new LayoutSlot { Index = 3, Name = "右下", X = x + halfW + g, Y = y + halfH + g, Width = w - halfW - g, Height = h - halfH - g });
                break;

            case LayoutMode.LeftRight:
                int lrHalf = (w - g) / 2;
                Slots.Add(new LayoutSlot { Index = 0, Name = "左", X = x, Y = y, Width = lrHalf, Height = h });
                Slots.Add(new LayoutSlot { Index = 1, Name = "右", X = x + lrHalf + g, Y = y, Width = w - lrHalf - g, Height = h });
                break;

            case LayoutMode.TopBottom:
                int tbHalf = (h - g) / 2;
                Slots.Add(new LayoutSlot { Index = 0, Name = "上", X = x, Y = y, Width = w, Height = tbHalf });
                Slots.Add(new LayoutSlot { Index = 1, Name = "下", X = x, Y = y + tbHalf + g, Width = w, Height = h - tbHalf - g });
                break;
        }
    }

    /// <summary>
    /// 应用布局：将嵌入窗口移动到对应槽位
    /// </summary>
    public void ApplyLayout()
    {
        var windows = _windowManager.ManagedWindows.ToList();
        if (windows.Count == 0 || Slots.Count == 0) return;

        if (_currentMode == LayoutMode.Stacked)
        {
            // 堆叠模式
            var slot = Slots[0];
            foreach (var win in windows)
            {
                win.SlotIndex = 0;
                _windowManager.RepositionEmbedded(win.Handle, slot.X, slot.Y, slot.Width, slot.Height);
            }

            // 激活第一个活跃窗口，顶部窗口
            var active = windows.FirstOrDefault(w => w.IsActive) ?? windows[0];
            _windowManager.ShowOnly(active.Handle);
        }
        else
        {
            // 分区模式
            _windowManager.ShowAll();
            for (int i = 0; i < windows.Count && i < Slots.Count; i++)
            {
                var win = windows[i];
                var slot = Slots[i];
                win.SlotIndex = slot.Index;
                _windowManager.RepositionEmbedded(win.Handle, slot.X, slot.Y, slot.Width, slot.Height);
            }

            // 超出槽位数的窗口隐藏
            for (int i = Slots.Count; i < windows.Count; i++)
            {
                NativeMethods.ShowWindow(windows[i].Handle, 0);
            }
        }
    }

    /// <summary>
    /// 堆叠模式切换到指定窗口
    /// </summary>
    public void SwitchToWindow(ManagedWindow window)
    {
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

    /// <summary>
    /// 切换到下一个窗口
    /// </summary>
    public void SwitchToNext()
    {
        var windows = _windowManager.ManagedWindows;
        if (windows.Count <= 1) return;
        var active = windows.FirstOrDefault(w => w.IsActive);
        int idx = active != null ? windows.IndexOf(active) : -1;
        int next = (idx + 1) % windows.Count;
        SwitchToWindow(windows[next]);
    }

    /// <summary>
    /// 切换到上一个窗口
    /// </summary>
    public void SwitchToPrevious()
    {
        var windows = _windowManager.ManagedWindows;
        if (windows.Count <= 1) return;
        var active = windows.FirstOrDefault(w => w.IsActive);
        int idx = active != null ? windows.IndexOf(active) : 0;
        int prev = (idx - 1 + windows.Count) % windows.Count;
        SwitchToWindow(windows[prev]);
    }
}
