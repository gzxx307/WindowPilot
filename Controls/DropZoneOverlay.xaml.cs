using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace WindowPilot.Controls;

public partial class DropZoneOverlay : Window
{
    private readonly Storyboard _fadeIn;
    private readonly Storyboard _fadeOut;

    public DropZoneOverlay()
    {
        InitializeComponent();

        _fadeIn = new Storyboard();
        var fadeInAnim = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200));
        Storyboard.SetTarget(fadeInAnim, OverlayBorder);
        Storyboard.SetTargetProperty(fadeInAnim, new PropertyPath("Opacity"));
        _fadeIn.Children.Add(fadeInAnim);

        _fadeOut = new Storyboard();
        var fadeOutAnim = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(150));
        Storyboard.SetTarget(fadeOutAnim, OverlayBorder);
        Storyboard.SetTargetProperty(fadeOutAnim, new PropertyPath("Opacity"));
        _fadeOut.Children.Add(fadeOutAnim);
        _fadeOut.Completed += (_, _) => Hide();
    }

    /// <summary>
    /// 设置覆盖区域。
    /// 
    /// <paramref name="physicalRect"/> 物理像素坐标
    /// <paramref name="dpiScaleX"/> / <paramref name="dpiScaleY"/> DPI 缩放比
    /// 
    /// </summary>
    public void SetBounds(Rect physicalRect, double dpiScaleX, double dpiScaleY)
    {
        Left = physicalRect.X / dpiScaleX;
        Top = physicalRect.Y / dpiScaleY;
        Width = physicalRect.Width / dpiScaleX;
        Height = physicalRect.Height / dpiScaleY;
    }

    public void ShowOverlay()
    {
        if (!IsVisible) Show();
        _fadeIn.Begin(this);
    }

    public void HideOverlay()
    {
        if (IsVisible) _fadeOut.Begin(this);
    }

    public void UpdateHighlight(bool isMouseInZone)
    {
        OverlayBorder.Background = isMouseInZone
            ? (Brush)FindResource("DropZoneActiveBrush")
            : new SolidColorBrush(Color.FromArgb(0x22, 0x00, 0x7A, 0xCC));

        OverlayBorder.BorderThickness = isMouseInZone
            ? new Thickness(3) : new Thickness(2);
    }
}