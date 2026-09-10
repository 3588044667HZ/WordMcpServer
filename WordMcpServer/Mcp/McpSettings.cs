using System;
using System.IO;
using Newtonsoft.Json;
using WordMcpServer.Logging;

namespace WordMcpServer.Mcp
{
    /// <summary>
    /// 服务器运行配置。默认监听 127.0.0.1:5101（仅回环地址，拒绝远程连接）。
    /// 可通过插件目录下的 mcp.config.json 覆盖；配置文件不存在时自动生成模板。
    /// 端口与 PptMcpServer（5100）错开，两个插件可同时运行。
    /// </summary>
    public sealed class McpSettings
    {
        public const string DefaultHost = "127.0.0.1";
        public const int DefaultPort = 5101;
        public const string ConfigFileName = "mcp.config.json";

        /// <summary>监听主机名。出于安全考虑仅允许 127.0.0.1 / localhost。</summary>
        public string Host { get; set; } = DefaultHost;

        /// <summary>监听端口（1-65535）。</summary>
        public int Port { get; set; } = DefaultPort;

        /// <summary>日志目录，为空时使用默认目录（%APPDATA%\WordMcpServer\Logs）。</summary>
        public string LogDirectory { get; set; }

        public string GetHostForListening()
        {
            // 边界防御：即使配置被改成其他地址，也强制收敛到回环地址
            return "127.0.0.1";
        }

        /// <summary>
        /// 从插件所在目录加载配置；加载失败或值非法时回退到默认值并记录警告。
        /// </summary>
        public static McpSettings Load(string configDirectory, IMcpLogger logger)
        {
            var settings = new McpSettings();
            var configPath = Path.Combine(configDirectory, ConfigFileName);
            try
            {
                if (!File.Exists(configPath))
                {
                    // 首次运行：生成模板文件，方便用户按需修改
                    File.WriteAllText(configPath, JsonConvert.SerializeObject(settings, Formatting.Indented));
                    logger?.Info($"已生成默认配置文件：{configPath}");
                    return settings;
                }

                var parsed = JsonConvert.DeserializeObject<McpSettings>(File.ReadAllText(configPath));
                if (parsed != null)
                {
                    settings = parsed;
                }
            }
            catch (Exception ex)
            {
                logger?.Warn($"读取配置文件失败，使用默认配置：{ex.Message}");
                return new McpSettings();
            }

            if (settings.Port < 1 || settings.Port > 65535)
            {
                logger?.Warn($"配置的端口 {settings.Port} 非法，回退到默认端口 {DefaultPort}");
                settings.Port = DefaultPort;
            }

            if (string.IsNullOrWhiteSpace(settings.Host))
            {
                settings.Host = DefaultHost;
            }

            return settings;
        }
    }
}
