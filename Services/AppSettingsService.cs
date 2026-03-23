using System.IO;
using System.Text.Json;
using WindowPilot.Diagnostics;
using WindowPilot.Models;

namespace WindowPilot.Services;

/// <summary>
/// 轻量级应用配置持久化服务，将设置以 JSON 格式存储在可执行文件同级目录的
/// <c>settings.json</c> 文件中。读取失败时静默降级为默认值，写入失败时仅记录日志，
/// 不影响主流程。
/// </summary>
public class AppSettingsService
{
    private const string Cat      = "AppSettingsService";
    private const string FileName = "settings.json";

    // 配置文件存放在可执行文件同目录，随程序一起部署
    private static readonly string FilePath =
        Path.Combine(AppContext.BaseDirectory, FileName);

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented        = true,   // 人类可读格式，方便手动编辑
        AllowTrailingCommas  = true,   // 容忍手动编辑时遗留的尾逗号
    };

    /// <summary>当前生效的配置对象，Load 之前为默认值。</summary>
    public AppSettings Settings { get; private set; } = new();

    /// <summary>
    /// 从磁盘读取配置文件并反序列化。
    /// 文件不存在或解析失败时，<see cref="Settings"/> 保留默认值，不抛出异常。
    /// </summary>
    public void Load()
    {
        try
        {
            if (!File.Exists(FilePath))
            {
                Logger.Info($"配置文件不存在，使用默认值：{FilePath}", Cat);
                return;
            }

            var json   = File.ReadAllText(FilePath);
            var loaded = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions);

            if (loaded != null)
            {
                Settings = loaded;
                Logger.Info($"配置已加载 ← {FilePath}  LastLayoutMode={Settings.LastLayoutMode}", Cat);
            }
            else
            {
                Logger.Warning("配置文件内容为空，使用默认值。", Cat);
            }
        }
        catch (Exception ex)
        {
            Logger.Error("加载配置失败，使用默认值", ex, Cat);
            Settings = new AppSettings();
        }
    }

    /// <summary>
    /// 将 <see cref="Settings"/> 序列化并写入磁盘。
    /// 写入失败时仅记录错误日志，不抛出异常，不影响主流程。
    /// </summary>
    public void Save()
    {
        try
        {
            var json = JsonSerializer.Serialize(Settings, JsonOptions);
            File.WriteAllText(FilePath, json);
            Logger.Debug($"配置已保存 → {FilePath}  LastLayoutMode={Settings.LastLayoutMode}", Cat);
        }
        catch (Exception ex)
        {
            Logger.Error("保存配置失败", ex, Cat);
        }
    }
}
