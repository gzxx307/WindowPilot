using System.Text;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 枚举桌面上所有可纳入管理的顶层窗口，并提供过滤逻辑。
/// </summary>
public static class WindowEnumerator
{
    // 明确排除的窗口类名，这些均为系统 Shell 组件，不应被托管
    private static readonly HashSet<string> ExcludedClasses = new(StringComparer.OrdinalIgnoreCase)
    {
        "Progman",                    // 桌面图标宿主
        "Shell_TrayWnd",              // 主任务栏
        "Shell_SecondaryTrayWnd",     // 多显示器辅助任务栏
        "DV2ControlHost",             // 开始菜单
        "Windows.UI.Core.CoreWindow", // UWP 核心窗口
        "WorkerW",                    // 桌面背景
        "SHELLDLL_DefView",           // 桌面文件视图
        "EdgeUiInputTopWndClass",
        "NativeHWNDHost",
        "Button",
    };

    /// <summary>
    /// 枚举所有当前符合管理条件的顶层窗口。
    /// </summary>
    /// <param name="excludeOwnerHwnd">需要排除的窗口句柄，通常传入宿主窗口自身。</param>
    /// <returns>可管理的窗口句柄列表。</returns>
    public static List<IntPtr> GetManageableWindows(IntPtr? excludeOwnerHwnd = null)
    {
        var windows     = new List<IntPtr>();
        var shellHwnd   = NativeMethods.GetShellWindow();

        // EnumWindows 按 Z 序遍历所有顶层窗口
        NativeMethods.EnumWindows((hWnd, _) =>
        {
            if (ShouldInclude(hWnd, shellHwnd, excludeOwnerHwnd))
                windows.Add(hWnd);
            return true; // 返回 true 继续枚举
        }, IntPtr.Zero);

        return windows;
    }

    /// <summary>
    /// 判断指定窗口是否满足被纳入管理的条件。
    /// </summary>
    /// <param name="hWnd">待判断的窗口句柄。</param>
    /// <param name="shellHwnd">Shell 窗口句柄缓存，为 null 时内部自动获取。</param>
    /// <param name="excludeOwnerHwnd">需要排除的特定窗口句柄，通常为宿主自身。</param>
    /// <returns>满足所有条件时返回 true，任一条件不满足时返回 false。</returns>
    public static bool ShouldInclude(IntPtr hWnd, IntPtr? shellHwnd = null, IntPtr? excludeOwnerHwnd = null)
    {
        shellHwnd ??= NativeMethods.GetShellWindow();

        // Shell 窗口是桌面本身，不可托管
        if (hWnd == shellHwnd) return false;

        // 排除宿主窗口自身
        if (excludeOwnerHwnd.HasValue && hWnd == excludeOwnerHwnd.Value) return false;

        // 不可见窗口无意义，不纳入
        if (!NativeMethods.IsWindowVisible(hWnd)) return false;

        // 无标题的窗口通常是内部辅助窗口，跳过
        if (NativeMethods.GetWindowTextLength(hWnd) == 0) return false;

        long style   = NativeMethods.GetWindowLongSafe(hWnd, NativeConstants.GWL_STYLE);
        long exStyle = NativeMethods.GetWindowLongSafe(hWnd, NativeConstants.GWL_EXSTYLE);

        // 子窗口不是顶层窗口，不纳入管理
        if ((style & NativeConstants.WS_CHILD) != 0) return false;

        // 禁用状态的窗口无法接受用户输入，无需托管
        if ((style & NativeConstants.WS_DISABLED) != 0) return false;

        // 工具窗口通常是辅助浮窗，除非同时带有 AppWindow 标志表明它想显示在任务栏
        if ((exStyle & NativeConstants.WS_EX_TOOLWINDOW) != 0 &&
            (exStyle & NativeConstants.WS_EX_APPWINDOW) == 0)
            return false;

        // 不激活窗口是系统 UI 层元素，不纳入
        if ((exStyle & NativeConstants.WS_EX_NOACTIVATE) != 0) return false;

        // 有所有者窗口的非 AppWindow 是从属弹窗，不属于独立顶层窗口
        IntPtr owner = NativeMethods.GetWindow(hWnd, 4); // GW_OWNER = 4
        if (owner != IntPtr.Zero && (exStyle & NativeConstants.WS_EX_APPWINDOW) == 0)
            return false;

        // 对照排除类名列表
        var classNameSb = new StringBuilder(256);
        NativeMethods.GetClassName(hWnd, classNameSb, classNameSb.Capacity);
        string className = classNameSb.ToString();
        if (ExcludedClasses.Contains(className)) return false;

        return true;
    }

    /// <summary>
    /// 获取指定窗口的标题文字。
    /// </summary>
    /// <param name="hWnd">目标窗口句柄。</param>
    /// <returns>窗口标题字符串，无标题时返回空字符串。</returns>
    public static string GetWindowTitle(IntPtr hWnd)
    {
        int length = NativeMethods.GetWindowTextLength(hWnd);
        if (length == 0) return string.Empty;
        var sb = new StringBuilder(length + 1);
        NativeMethods.GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    /// <summary>
    /// 获取指定窗口的注册类名。
    /// </summary>
    /// <param name="hWnd">目标窗口句柄。</param>
    /// <returns>类名字符串。</returns>
    public static string GetWindowClassName(IntPtr hWnd)
    {
        var sb = new StringBuilder(256);
        NativeMethods.GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }
}