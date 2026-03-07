using System.Windows;
using System.Windows.Media.Animation;

namespace WindowPilot.Controls;

/// <summary>
/// 蓝色覆盖层：当用户拖拽已托管窗口经过另一个窗口的槽位时显示，
/// 表示"松开可互换位置"。
/// </summary>
public partial class SwapIndicatorOverlay : Window
{
    public SwapIndicatorOverlay()
    {
        InitializeComponent();
    }

    /// <summary>
    /// 在指定的屏幕矩形处显示覆盖层，并播放淡入动画
    /// </summary>
    public void ShowAtRect(Rect screenRect)
    {
        Left   = screenRect.Left;
        Top    = screenRect.Top;
        Width  = screenRect.Width;
        Height = screenRect.Height;

        if (!IsVisible) Show();

        OverlayBorder.BeginAnimation(
            UIElement.OpacityProperty,
            new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150))));
    }

    /// <summary>
    /// 淡出后隐藏覆盖层
    /// </summary>
    public void HideOverlay()
    {
        if (!IsVisible) return;

        var anim = new DoubleAnimation(
            OverlayBorder.Opacity, 0,
            new Duration(TimeSpan.FromMilliseconds(150)));
        anim.Completed += (_, _) => { if (IsVisible) Hide(); };
        OverlayBorder.BeginAnimation(UIElement.OpacityProperty, anim);
    }
}
