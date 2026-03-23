using WindowPilot.Services;

namespace WindowPilot.Models;

/// <summary>
/// 应用程序持久化配置数据模型，所有字段均应提供合理的默认值，
/// 以便配置文件缺失时程序也能正常启动。
/// </summary>
public class AppSettings
{
    /// <summary>
    /// 上次退出时的布局模式，下次启动时自动恢复。
    /// </summary>
    public LayoutService.LayoutMode LastLayoutMode { get; set; } = LayoutService.LayoutMode.Stacked;
}
