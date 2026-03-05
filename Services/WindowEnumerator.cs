using System.Text;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 枚举桌面上所有可管理的窗口
/// </summary>
public static class WindowEnumerator
{
    private static readonly HashSet<string> ExcludedClasses = new(StringComparer.OrdinalIgnoreCase)
    {
        "Progman",
        "Shell_TrayWnd",
        "Shell_SecondaryTrayWnd",
        "DV2ControlHost",
        "Windows.UI.Core.CoreWindow",
        "WorkerW",
        "SHELLDLL_DefView",
        "EdgeUiInputTopWndClass",
        "NativeHWNDHost",
        "Button",
    };

    /// <summary>
    /// 获取所有可管理的顶层窗口句柄
    /// </summary>
    public static List<IntPtr> GetManageableWindows(IntPtr? excludeOwnerHwnd = null)
    {
        var windows = new List<IntPtr>();
        var shellHwnd = NativeMethods.GetShellWindow();

        NativeMethods.EnumWindows((hWnd, _) =>
        {
            if (ShouldInclude(hWnd, shellHwnd, excludeOwnerHwnd))
                windows.Add(hWnd);
            return true;
        }, IntPtr.Zero);

        return windows;
    }

    /// <summary>
    /// 判断一个窗口是否应该被纳入管理候选
    /// </summary>
    public static bool ShouldInclude(IntPtr hWnd, IntPtr? shellHwnd = null, IntPtr? excludeOwnerHwnd = null)
    {
        shellHwnd ??= NativeMethods.GetShellWindow();

        // 排除 shell 窗口
        if (hWnd == shellHwnd) return false;

        // 排除自己
        if (excludeOwnerHwnd.HasValue && hWnd == excludeOwnerHwnd.Value) return false;

        // 必须可见
        if (!NativeMethods.IsWindowVisible(hWnd)) return false;

        // 必须有标题
        if (NativeMethods.GetWindowTextLength(hWnd) == 0) return false;

        // 获取样式
        long style = NativeMethods.GetWindowLongSafe(hWnd, NativeConstants.GWL_STYLE);
        long exStyle = NativeMethods.GetWindowLongSafe(hWnd, NativeConstants.GWL_EXSTYLE);

        // 排除子窗口
        if ((style & NativeConstants.WS_CHILD) != 0) return false;

        // 排除禁用窗口
        if ((style & NativeConstants.WS_DISABLED) != 0) return false;

        // 排除工具窗口（除非同时有 AppWindow 样式）
        if ((exStyle & NativeConstants.WS_EX_TOOLWINDOW) != 0 &&
            (exStyle & NativeConstants.WS_EX_APPWINDOW) == 0)
            return false;

        // 排除无激活窗口
        if ((exStyle & NativeConstants.WS_EX_NOACTIVATE) != 0) return false;

        // 必须有所有者窗口为空（顶层窗口）或自身是 AppWindow
        IntPtr owner = NativeMethods.GetWindow(hWnd, 4);
        if (owner != IntPtr.Zero && (exStyle & NativeConstants.WS_EX_APPWINDOW) == 0)
            return false;

        // 排除特定类名
        var classNameSb = new StringBuilder(256);
        NativeMethods.GetClassName(hWnd, classNameSb, classNameSb.Capacity);
        string className = classNameSb.ToString();
        if (ExcludedClasses.Contains(className)) return false;

        return true;
    }

    /// <summary>
    /// 获取窗口标题
    /// </summary>
    public static string GetWindowTitle(IntPtr hWnd)
    {
        int length = NativeMethods.GetWindowTextLength(hWnd);
        if (length == 0) return string.Empty;
        var sb = new StringBuilder(length + 1);
        NativeMethods.GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    /// <summary>
    /// 获取窗口类名
    /// </summary>
    public static string GetWindowClassName(IntPtr hWnd)
    {
        var sb = new StringBuilder(256);
        NativeMethods.GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }
}
