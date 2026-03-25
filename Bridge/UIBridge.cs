using System.IO;
using System.Text.Json;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Web.WebView2.Core;
using WindowPilot.Diagnostics;
using WindowPilot.Services;

namespace WindowPilot.Bridge;

/// <summary>
/// 统一管理 3 个 WebView2 区域（工具栏 / 侧边栏 / 状态栏）的通信桥。
///
/// JS → C# : window.chrome.webview.postMessage({type, ...})  （直接传对象）
/// C# → JS : CoreWebView2.PostWebMessageAsString(json)
/// </summary>
public class UIBridge
{
    private const string Cat = "UIBridge";

    private CoreWebView2? _toolbar;
    private CoreWebView2? _sidebar;
    private CoreWebView2? _status;

    private readonly WindowManagerService _windowManager;
    private readonly LayoutService        _layout;
    private readonly Dispatcher           _dispatcher;

    // ── 业务事件（MainWindow 订阅） ────────────────────────────────────
    public event Action<LayoutService.LayoutMode>? LayoutModeRequested;
    public event Action?                           RefreshLayoutRequested;
    public event Action?                           ReleaseAllRequested;
    public event Action?                           ToggleSidebarRequested;

    public UIBridge(WindowManagerService windowManager, LayoutService layout, Dispatcher dispatcher)
    {
        _windowManager = windowManager;
        _layout        = layout;
        _dispatcher    = dispatcher;
    }

    // ── 初始化 ────────────────────────────────────────────────────────

    public void AttachToolbar(CoreWebView2 core)
    {
        _toolbar = core;
        _toolbar.WebMessageReceived += OnMessageReceived;
        Logger.Info("UIBridge 已绑定 ToolbarWebView。", Cat);
    }

    public void AttachSidebar(CoreWebView2 core)
    {
        _sidebar = core;
        _sidebar.WebMessageReceived += OnMessageReceived;
        Logger.Info("UIBridge 已绑定 SidebarWebView。", Cat);
    }

    public void AttachStatus(CoreWebView2 core)
    {
        _status = core;
        _status.WebMessageReceived += OnMessageReceived;
        Logger.Info("UIBridge 已绑定 StatusWebView。", Cat);
    }

    // ── JS → C#（所有区域共用同一个处理器） ─────────────────────────────

    private void OnMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            var json = e.WebMessageAsJson;
            Logger.Debug($"收到 UI 消息: {json}", Cat);

            using var doc  = JsonDocument.Parse(json);
            var       root = doc.RootElement;
            var       type = root.GetProperty("type").GetString();

            switch (type)
            {
                // ── 初始化 ──────────────────────────────────────────
                case "toolbarReady":
                    PushLayoutMode(_layout.CurrentMode);
                    PushWindowCount(_windowManager.ManagedWindows.Count);
                    break;

                case "sidebarReady":
                    PushWindowList();
                    PushSidebarCollapsed(!IsSidebarExpanded());
                    break;

                case "statusReady":
                    PostStatus("就绪 — 拖拽窗口到侧边栏以接管，或按 Ctrl+Alt+G │ Ctrl+Alt+M 拖拽移动窗口");
                    break;

                // ── 工具栏按钮 ──────────────────────────────────────
                case "layoutMode":
                {
                    var modeName = root.GetProperty("mode").GetString();
                    if (Enum.TryParse<LayoutService.LayoutMode>(modeName, out var mode))
                        LayoutModeRequested?.Invoke(mode);
                    break;
                }

                case "refreshLayout":
                    RefreshLayoutRequested?.Invoke();
                    break;

                case "releaseAll":
                    ReleaseAllRequested?.Invoke();
                    break;

                case "toggleSidebar":
                    ToggleSidebarRequested?.Invoke();
                    break;

                // ── 侧边栏交互 ──────────────────────────────────────
                case "activate":
                {
                    var handle = ParseHandle(root, "handle");
                    _dispatcher.BeginInvoke(() =>
                    {
                        var target = _windowManager.FindByHandle(handle);
                        if (target == null) return;
                        _layout.SwitchToWindow(target);
                        Logger.Info($"activate → \"{target.Title}\"", Cat);
                        PushWindowList();
                    });
                    break;
                }

                case "release":
                {
                    var handle = ParseHandle(root, "handle");
                    _dispatcher.BeginInvoke(() => _windowManager.ReleaseWindow(handle));
                    break;
                }

                case "swapWindows":
                {
                    var h1 = ParseHandle(root, "handle1");
                    var h2 = ParseHandle(root, "handle2");
                    _dispatcher.BeginInvoke(() =>
                    {
                        var w1 = _windowManager.FindByHandle(h1);
                        var w2 = _windowManager.FindByHandle(h2);
                        if (w1 == null || w2 == null) return;
                        _windowManager.SwapOrder(h1, h2);
                        _layout.ApplyLayout();
                        PushWindowList();
                        PushStatus($"已交换位置：{w1.Title}  ⇄  {w2.Title}");
                    });
                    break;
                }

                default:
                    Logger.Warning($"未知 UI 消息类型: {type}", Cat);
                    break;
            }
        }
        catch (Exception ex)
        {
            Logger.Error("UIBridge 消息处理异常", ex, Cat);
        }
    }

    private static IntPtr ParseHandle(JsonElement root, string prop)
        => new IntPtr(long.Parse(root.GetProperty(prop).GetString()!));

    // ── 内部状态查询（供 sidebarReady 使用） ──────────────────────────
    private bool _sidebarExpanded = true;
    public void SetSidebarExpanded(bool expanded) => _sidebarExpanded = expanded;
    private bool IsSidebarExpanded() => _sidebarExpanded;

    // ── C# → JS：推送方法 ────────────────────────────────────────────

    /// <summary>
    /// 推送完整窗口列表到侧边栏，并同步工具栏的窗口数量。
    /// </summary>
    public void PushWindowList()
    {
        if (_sidebar == null && _toolbar == null) return;

        _dispatcher.BeginInvoke(() =>
        {
            try
            {
                var windows = _windowManager.ManagedWindows.Select(w => new
                {
                    handle      = w.Handle.ToString(),
                    title       = w.Title,
                    processName = w.ProcessName,
                    icon        = IconToBase64(w.Icon),
                    isActive    = w.IsActive,
                }).ToArray();

                var listPayload = JsonSerializer.Serialize(new
                {
                    type    = "windowList",
                    windows,
                });

                PostTo(_sidebar, listPayload);

                // 同步工具栏的窗口计数
                var countPayload = JsonSerializer.Serialize(new
                {
                    type  = "windowCount",
                    count = windows.Length,
                });
                PostTo(_toolbar, countPayload);
            }
            catch (Exception ex)
            {
                Logger.Error("PushWindowList 异常", ex, Cat);
            }
        });
    }

    /// <summary>推送布局模式高亮到工具栏。</summary>
    public void PushLayoutMode(LayoutService.LayoutMode mode)
    {
        if (_toolbar == null) return;
        _dispatcher.BeginInvoke(() =>
            PostTo(_toolbar, JsonSerializer.Serialize(new
            {
                type       = "layoutMode",
                activeMode = mode.ToString(),
            })));
    }

    /// <summary>推送窗口数量到工具栏。</summary>
    public void PushWindowCount(int count)
    {
        if (_toolbar == null) return;
        _dispatcher.BeginInvoke(() =>
            PostTo(_toolbar, JsonSerializer.Serialize(new { type = "windowCount", count })));
    }

    /// <summary>推送状态文本到状态栏。</summary>
    public void PushStatus(string message)
    {
        if (_status == null) return;
        _dispatcher.BeginInvoke(() =>
            PostTo(_status, JsonSerializer.Serialize(new { type = "status", message })));
    }

    /// <summary>推送侧边栏折叠状态（不展开时为 true）。</summary>
    public void PushSidebarCollapsed(bool collapsed)
    {
        if (_sidebar == null) return;
        _sidebarExpanded = !collapsed;
        _dispatcher.BeginInvoke(() =>
            PostTo(_sidebar, JsonSerializer.Serialize(new { type = "collapsed", collapsed })));
    }

    /// <summary>推送拖入遮罩显示/隐藏到侧边栏。</summary>
    public void PushDropOverlay(bool visible)
    {
        if (_sidebar == null) return;
        _dispatcher.BeginInvoke(() =>
            PostTo(_sidebar, JsonSerializer.Serialize(new { type = "dropOverlay", visible })));
    }

    // ── 工具方法 ──────────────────────────────────────────────────────

    // 包装一下，直接传 json 字符串避免二次序列化
    private static void PostTo(CoreWebView2? core, string json)
    {
        if (core == null) return;
        try { core.PostWebMessageAsString(json); }
        catch (Exception ex) { Logger.Error($"PostTo 失败: {ex.Message}", Cat); }
    }

    // 便捷重载，供外部直接传对象
    private static void PostStatus(string message) { /* 通过 PushStatus 调用 */ }

    private static string? IconToBase64(ImageSource? icon)
    {
        if (icon is not BitmapSource bitmap) return null;
        try
        {
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bitmap));
            using var ms = new MemoryStream();
            encoder.Save(ms);
            return "data:image/png;base64," + Convert.ToBase64String(ms.ToArray());
        }
        catch { return null; }
    }
}