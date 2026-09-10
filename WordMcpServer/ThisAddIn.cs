using System;
using System.Windows.Forms;
using WordMcpServer.Logging;
using WordMcpServer.Mcp;
using WordMcpServer.Mcp.Protocol;
using WordMcpServer.Mcp.Server;
using WordMcpServer.Mcp.Threading;
using WordMcpServer.Mcp.Tools;
using WordMcpServer.Services;

namespace WordMcpServer
{
    /// <summary>
    /// VSTO 插件入口。
    ///
    /// 生命周期：
    ///   Startup —— 安装 UI 线程同步上下文 → 组装服务（DI）→ 启动 Kestrel
    ///   Shutdown —— 优雅停止 Kestrel → 释放资源
    ///
    /// 线程说明：Startup/Shutdown 都发生在 Word 主 UI 线程；
    /// HTTP 请求在线程池线程处理，Word 调用通过 OfficeThreadMarshaller
    /// 封送回主线程（详见该类的注释）。
    /// </summary>
    public partial class ThisAddIn
    {
        private FileMcpLogger _logger;
        private McpHttpHost _httpHost;
        private OfficeThreadMarshaller _marshaller;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                // 1. 日志（先于一切，保证后续失败可追踪）
                _logger = new FileMcpLogger();

                // 2. 加载配置（插件目录下的 mcp.config.json，不存在时自动生成模板）
                var settings = McpSettings.Load(AppDomain.CurrentDomain.BaseDirectory, _logger);

                // 3. 安装 Word 主线程同步上下文（必须在主线程上执行）
                _marshaller = OfficeThreadMarshaller.Install(_logger);

                // 4. 组装服务：Word 服务 → 工具注册表 → MCP 协议处理器 → HTTP 宿主
                var wordService = new WordService(this.Application, _marshaller, _logger);
                var toolRegistry = new ToolRegistry(wordService, _logger);
                var handler = new McpProtocolProcessor(toolRegistry, _logger);
                _httpHost = new McpHttpHost(settings, _logger, handler, toolRegistry);

                // 5. 启动本地 HTTP 服务器（Kestrel 在线程池运行，不阻塞 UI）
                _httpHost.Start();

                _logger.Info($"WordMcpServer 已就绪：http://{settings.GetHostForListening()}:{settings.Port}/mcp");
            }
            catch (Exception ex)
            {
                // 启动失败不影响 Word 本身的使用，弹窗提示用户
                _logger?.Error("WordMcpServer 启动失败", ex);
                MessageBox.Show("WordMcpServer MCP 服务器启动失败：\n" + ex.Message,
                    "WordMcpServer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                _httpHost?.Stop();
                _marshaller?.Dispose();
                _logger?.Info("WordMcpServer 已关闭");
            }
            catch (Exception ex)
            {
                // 关闭阶段异常不能阻止插件卸载
                _logger?.Error("WordMcpServer 关闭异常", ex);
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
