using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using WindowPilot.Controls;
using WindowPilot.Diagnostics;
using WindowPilot.Models;
using WindowPilot.Native;

namespace WindowPilot.Services;

/// <summary>
/// 管理侧边栏悬停时的缩略图预览浮窗。
/// <para>
/// 对于通过 <c>SetParent</c> 嵌入的子窗口（WS_CHILD），跳过 DWM Thumbnail，
/// 改用「Z-order 临时置顶 + PrintWindow(PW_RENDERFULLCONTENT)」截图。
/// </para>
/// <para>
/// 为何需要临时置顶：<br/>
/// <c>PW_RENDERFULLCONTENT</c> 对子窗口的实现是从父窗口的 GPU 共享合成面读取像素，
/// 而非让目标窗口重新自绘——因此读取到的始终是 Z-order 最顶层兄弟窗口的内容。
/// 将目标临时提到顶层后，合成面会在同一帧内更新为它的内容，PrintWindow 随即捕获到正确画面。
/// 整个置顶→截图→还原在一次消息泵内同步完成（远不足 1ms），DWM 不会渲染中间帧，用户无感。
/// </para>
/// </summary>
public class ThumbnailPreviewService : IDisposable
{
    private const string Cat = "ThumbnailPreview";

    private ThumbnailPreviewWindow? _previewWindow;
    private IntPtr _thumbnailHandle = IntPtr.Zero;
    private ManagedWindow? _currentTarget;

    private const double PreviewWidth  = 320.0;
    private const double PreviewHeight = 220.0;
    private const double GapDip        = 10.0;

    public void Show(ManagedWindow target,
                     Point          itemScreenTopLeft,
                     double         sidebarScreenRight,
                     double         dpiScale)
    {
        // 同目标复用：只更新位置，不重新截图
        if (_currentTarget?.Handle == target.Handle && _previewWindow?.IsVisible == true)
        {
            RepositionWindow(itemScreenTopLeft, sidebarScreenRight, dpiScale);
            return;
        }

        HideInternal();
        _currentTarget = target;

        _previewWindow ??= new ThumbnailPreviewWindow();
        _previewWindow.SetWindowInfo(target.Title, target.ProcessName, target.Icon);
        RepositionWindow(itemScreenTopLeft, sidebarScreenRight, dpiScale);
        _previewWindow.Show();

        // 检查是否为通过 SetParent 嵌入的子窗口
        long windowStyle   = NativeMethods.GetWindowLongSafe(target.Handle, NativeConstants.GWL_STYLE);
        bool isChildWindow = (windowStyle & NativeConstants.WS_CHILD) != 0;

        if (!isChildWindow)
        {
            // 顶层窗口：优先尝试 DWM 实时合成
            var previewHwnd = new WindowInteropHelper(_previewWindow).Handle;
            if (previewHwnd != IntPtr.Zero)
            {
                int hr = NativeMethods.DwmRegisterThumbnail(
                    previewHwnd, target.Handle, out _thumbnailHandle);

                if (hr == 0 && _thumbnailHandle != IntPtr.Zero)
                {
                    _previewWindow.SetDwmMode();
                    ApplyThumbnailProperties(_previewWindow.GetThumbnailDestRect());
                    Logger.Debug($"DWM 预览成功: \"{target.Title}\"  thumb=0x{_thumbnailHandle:X}", Cat);
                    return;
                }

                Logger.Trace(
                    $"DwmRegisterThumbnail 失败 hr=0x{hr:X8}，回退到 PrintWindow: \"{target.Title}\"", Cat);
                _thumbnailHandle = IntPtr.Zero;
            }
        }
        else
        {
            Logger.Trace($"WS_CHILD 窗口，跳过 DWM，使用 Z-swap + PrintWindow: \"{target.Title}\"", Cat);
        }

        // 子窗口固定路径 / 顶层窗口 DWM 失败时的回退
        var bitmap = CaptureWithPrintWindow(target.Handle, isChildWindow);
        _previewWindow.SetFallbackImage(bitmap);

        Logger.Debug(
            bitmap != null
                ? $"PrintWindow 截图成功: \"{target.Title}\"  {bitmap.PixelWidth}×{bitmap.PixelHeight}"
                : $"PrintWindow 截图失败: \"{target.Title}\"",
            Cat);
    }

    public void Hide()
    {
        Logger.Trace("ThumbnailPreviewService.Hide()", Cat);
        HideInternal();
        _currentTarget = null;
    }

    private void HideInternal()
    {
        if (_thumbnailHandle != IntPtr.Zero)
        {
            NativeMethods.DwmUnregisterThumbnail(_thumbnailHandle);
            Logger.Trace($"DwmUnregisterThumbnail: 0x{_thumbnailHandle:X}", Cat);
            _thumbnailHandle = IntPtr.Zero;
        }
        _previewWindow?.Hide();
    }

    private void RepositionWindow(Point itemScreenTopLeft, double sidebarScreenRight, double dpiScale)
    {
        if (_previewWindow == null) return;

        double left = sidebarScreenRight / dpiScale + GapDip;
        double top  = itemScreenTopLeft.Y / dpiScale;

        var workArea = SystemParameters.WorkArea;
        double maxTop = workArea.Bottom - PreviewHeight;
        if (top > maxTop) top = maxTop;
        if (top < workArea.Top) top = workArea.Top;

        _previewWindow.Left = left;
        _previewWindow.Top  = top;
    }

    private void ApplyThumbnailProperties(RECT destRect)
    {
        if (_thumbnailHandle == IntPtr.Zero) return;

        var props = new DWM_THUMBNAIL_PROPERTIES
        {
            dwFlags =
                NativeConstants.DWM_TNP_RECTDESTINATION      |
                NativeConstants.DWM_TNP_VISIBLE              |
                NativeConstants.DWM_TNP_OPACITY              |
                NativeConstants.DWM_TNP_SOURCECLIENTAREAONLY,
            rcDestination         = destRect,
            opacity               = 255,
            fVisible              = true,
            fSourceClientAreaOnly = true,
        };

        int hr = NativeMethods.DwmUpdateThumbnailProperties(_thumbnailHandle, ref props);
        if (hr != 0)
            Logger.Warning($"DwmUpdateThumbnailProperties 失败: hr=0x{hr:X8}", Cat);
    }

    /// <summary>
    /// 使用 PrintWindow + GDI 截取窗口画面。
    /// <para>
    /// 对于子窗口（<paramref name="isChild"/> = true），截图前先将其临时置于兄弟窗口 Z-order 顶层，
    /// 使 GPU 合成面更新为该窗口内容，截完后立刻还原原先置顶的兄弟窗口。
    /// 同步操作在一次消息泵内完成，不产生可见闪烁。
    /// </para>
    /// </summary>
    private static BitmapSource? CaptureWithPrintWindow(IntPtr hwnd, bool isChild)
    {
        if (!NativeMethods.GetWindowRect(hwnd, out RECT rect)) return null;
        int w = rect.Width;
        int h = rect.Height;
        if (w <= 0 || h <= 0) return null;

        // ── 子窗口：临时 Z-order 置顶 ──────────────────────────────────
        // PW_RENDERFULLCONTENT 从父窗口 GPU 合成面读取像素，必须让目标先成为最顶层子窗口，
        // 合成面才会反映其内容。置顶→截图→还原三步同步完成，不会被 DWM 渲染中间帧。
        IntPtr prevTopChild    = IntPtr.Zero;
        bool   needRestoreZOrder = false;

        if (isChild)
        {
            IntPtr parentHwnd = NativeMethods.GetParent(hwnd);
            prevTopChild      = NativeMethods.GetTopWindow(parentHwnd); // 当前最顶层兄弟

            if (prevTopChild != hwnd) // 目标不是已经在顶层，才需要置换
            {
                NativeMethods.SetWindowPos(
                    hwnd, NativeConstants.HWND_TOP,
                    0, 0, 0, 0,
                    NativeConstants.SWP_NOMOVE   |
                    NativeConstants.SWP_NOSIZE   |
                    NativeConstants.SWP_NOACTIVATE);
                needRestoreZOrder = true;

                Logger.Trace(
                    $"Z-swap: 0x{hwnd:X} 置顶（原顶层=0x{prevTopChild:X}）", "ThumbnailPreview");
            }
        }
        // ────────────────────────────────────────────────────────────────

        IntPtr hdcScreen = NativeMethods.GetDC(IntPtr.Zero);
        IntPtr hdcMem    = IntPtr.Zero;
        IntPtr hBmp      = IntPtr.Zero;
        IntPtr hOld      = IntPtr.Zero;

        try
        {
            if (hdcScreen == IntPtr.Zero) return null;

            hdcMem = NativeMethods.CreateCompatibleDC(hdcScreen);
            hBmp   = NativeMethods.CreateCompatibleBitmap(hdcScreen, w, h);
            hOld   = NativeMethods.SelectObject(hdcMem, hBmp);

            bool ok = NativeMethods.PrintWindow(hwnd, hdcMem, NativeMethods.PW_RENDERFULLCONTENT);
            if (!ok) return null;

            var bmpSource = Imaging.CreateBitmapSourceFromHBitmap(
                hBmp, IntPtr.Zero, Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
            bmpSource.Freeze();
            return bmpSource;
        }
        catch (Exception ex)
        {
            Logger.Warning($"CaptureWithPrintWindow 异常: {ex.Message}", "ThumbnailPreview");
            return null;
        }
        finally
        {
            // GDI 资源释放
            if (hOld      != IntPtr.Zero) NativeMethods.SelectObject(hdcMem, hOld);
            if (hBmp      != IntPtr.Zero) NativeMethods.DeleteObject(hBmp);
            if (hdcMem    != IntPtr.Zero) NativeMethods.DeleteDC(hdcMem);
            if (hdcScreen != IntPtr.Zero) NativeMethods.ReleaseDC(IntPtr.Zero, hdcScreen);

            // ── Z-order 还原：把之前的顶层窗口重新置顶 ──
            if (needRestoreZOrder && prevTopChild != IntPtr.Zero)
            {
                NativeMethods.SetWindowPos(
                    prevTopChild, NativeConstants.HWND_TOP,
                    0, 0, 0, 0,
                    NativeConstants.SWP_NOMOVE   |
                    NativeConstants.SWP_NOSIZE   |
                    NativeConstants.SWP_NOACTIVATE);

                Logger.Trace(
                    $"Z-swap 还原: 0x{prevTopChild:X} 重新置顶", "ThumbnailPreview");
            }
        }
    }

    public void Dispose()
    {
        Logger.Debug("ThumbnailPreviewService.Dispose()", Cat);
        HideInternal();
        _previewWindow?.Close();
        _previewWindow = null;
        GC.SuppressFinalize(this);
    }
}