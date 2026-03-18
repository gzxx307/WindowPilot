using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using WindowPilot.Native;

namespace WindowPilot.Controls;

/// <summary>
/// 缩略图预览浮窗。
/// 鼠标悬停侧边栏条目时显示在侧边栏右侧，实时（DWM）或静态（PrintWindow）展示窗口内容。
/// 设置了 WS_EX_NOACTIVATE，确保显示和交互时不抢夺输入焦点。
/// </summary>
public partial class ThumbnailPreviewWindow : Window
{
    // 标题区高度（WPF 逻辑像素），与 XAML 中 RowDefinition Height="36" 保持一致
    private const double HeaderHeightDip = 36.0;

    public ThumbnailPreviewWindow()
    {
        InitializeComponent();
    }

    /// <summary>
    /// HWND 就绪后追加 WS_EX_NOACTIVATE，防止浮窗在显示或鼠标点击时抢夺焦点。
    /// </summary>
    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);

        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero) return;

        long exStyle = NativeMethods.GetWindowLongSafe(hwnd, NativeConstants.GWL_EXSTYLE);
        exStyle |= NativeConstants.WS_EX_NOACTIVATE;
        NativeMethods.SetWindowLongPtrSafe(hwnd, NativeConstants.GWL_EXSTYLE, exStyle);
    }

    /// <summary>
    /// 更新标题区显示的窗口信息，在 Show() 之前调用。
    /// </summary>
    public void SetWindowInfo(string title, string processName, ImageSource? icon)
    {
        TitleTextBlock.Text   = title;
        ProcessTextBlock.Text = processName;
        IconImage.Source      = icon;
    }

    /// <summary>
    /// 显示 PrintWindow 截图作为预览内容（DWM 注册失败时的回退方案）。
    /// </summary>
    /// <param name="bitmap">截图位图，为 null 时显示占位错误提示。</param>
    public void SetFallbackImage(BitmapSource? bitmap)
    {
        if (bitmap != null)
        {
            FallbackImage.Source     = bitmap;
            FallbackImage.Visibility = Visibility.Visible;
            FallbackPanel.Visibility = Visibility.Collapsed;
        }
        else
        {
            FallbackImage.Visibility = Visibility.Collapsed;
            FallbackPanel.Visibility = Visibility.Visible;
        }
    }

    /// <summary>
    /// 切换为 DWM 实时预览模式（隐藏截图和占位面板，让 DWM 合成层直接覆盖）。
    /// </summary>
    public void SetDwmMode()
    {
        FallbackImage.Visibility = Visibility.Collapsed;
        FallbackPanel.Visibility = Visibility.Collapsed;
    }

    /// <summary>
    /// 计算 DWM 缩略图应渲染到的目标矩形（物理像素，相对于本窗口客户区左上角）。
    /// 渲染区域为标题区之下的全部空间。
    /// </summary>
    public RECT GetThumbnailDestRect()
    {
        var hwnd = new WindowInteropHelper(this).Handle;

        uint   dpi   = hwnd != IntPtr.Zero ? NativeMethods.GetDpiForWindow(hwnd) : 96u;
        double scale = dpi > 0 ? dpi / 96.0 : 1.0;

        int physW      = (int)(ActualWidth  * scale);
        int physH      = (int)(ActualHeight * scale);
        int physHeader = (int)(HeaderHeightDip * scale);

        // rcDestination：左上角 = (0, physHeader)，右下角 = (physW, physH)
        return new RECT(0, physHeader, physW, physH);
    }
}