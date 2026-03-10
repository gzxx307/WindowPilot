using System.Windows;
using System.Windows.Media.Animation;

namespace WindowPilot.Controls;

/// <summary>
/// 蓝色半透明覆盖层，在外部窗口拖入侧边栏时显示，提示用户可以放置。
/// 通过淡入/淡出动画显示和隐藏，覆盖位置跟随侧边栏矩形。
/// </summary>
public partial class DropZoneOverlay : Window
{
    // 构造函数，初始化 XAML 组件
    public DropZoneOverlay()
    {
        InitializeComponent();
    }

    /// <summary>
    /// 将覆盖层定位到指定屏幕矩形并以淡入动画显示。
    /// </summary>
    /// <param name="screenRect">覆盖层应覆盖的区域，使用屏幕物理像素坐标。</param>
    public void ShowAtRect(Rect screenRect)
    {
        // 将窗口定位到与侧边栏完全重合的位置
        Left   = screenRect.Left;
        Top    = screenRect.Top;
        Width  = screenRect.Width;
        Height = screenRect.Height;

        if (!IsVisible) Show();

        // 播放 180ms 淡入动画，从完全透明渐变到不透明
        OverlayBorder.BeginAnimation(
            UIElement.OpacityProperty,
            new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(180))));
    }

    // 播放淡出动画，动画结束后隐藏窗口
    public void HideOverlay()
    {
        if (!IsVisible) return;

        var anim = new DoubleAnimation(
            OverlayBorder.Opacity, 0,
            new Duration(TimeSpan.FromMilliseconds(200)));
        // 动画完成后隐藏窗口，释放视觉资源
        anim.Completed += (_, _) => { if (IsVisible) Hide(); };
        OverlayBorder.BeginAnimation(UIElement.OpacityProperty, anim);
    }
}