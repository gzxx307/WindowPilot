using System.Windows;
using System.Windows.Media.Animation;

namespace WindowPilot.Controls;

/// <summary>
/// 红色半透明覆盖层，在托管窗口被拖拽时显示于侧边栏上方，
/// 提示用户将窗口放置于此可解除托管。结构与 <see cref="DropZoneOverlay"/> 一致，颜色为红色。
/// </summary>
public partial class ReleaseZoneOverlay : Window
{
    // 构造函数，初始化 XAML 组件
    public ReleaseZoneOverlay()
    {
        InitializeComponent();
    }

    /// <summary>
    /// 将覆盖层定位到指定屏幕矩形并以淡入动画显示。
    /// </summary>
    /// <param name="screenRect">覆盖层应覆盖的区域，使用屏幕物理像素坐标。</param>
    public void ShowAtRect(Rect screenRect)
    {
        Left   = screenRect.Left;
        Top    = screenRect.Top;
        Width  = screenRect.Width;
        Height = screenRect.Height;

        if (!IsVisible) Show();

        // 播放 180ms 淡入动画
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
        // 动画完成后隐藏窗口
        anim.Completed += (_, _) => { if (IsVisible) Hide(); };
        OverlayBorder.BeginAnimation(UIElement.OpacityProperty, anim);
    }
}