using System.Windows;
using System.Windows.Media.Animation;

namespace WindowPilot.Controls;

public partial class DropZoneOverlay : Window
{
    public DropZoneOverlay()
    {
        InitializeComponent();
    }

    /// <summary>
    /// 在指定的屏幕矩形（物理像素）处显示覆盖层，并播放淡入动画
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
            new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(180))));
    }

    /// <summary>
    /// 淡出后隐藏覆盖层
    /// </summary>
    public void HideOverlay()
    {
        if (!IsVisible) return;

        var anim = new DoubleAnimation(
            OverlayBorder.Opacity, 0,
            new Duration(TimeSpan.FromMilliseconds(200)));
        anim.Completed += (_, _) => { if (IsVisible) Hide(); };
        OverlayBorder.BeginAnimation(UIElement.OpacityProperty, anim);
    }
}