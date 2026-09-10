using System;
using System.Net;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using WordMcpServer.Logging;
using WordMcpServer.Mcp.Protocol;
using WordMcpServer.Mcp.Tools;

namespace WordMcpServer.Mcp.Server
{
    /// <summary>
    /// Kestrel 自托管 HTTP 服务器生命周期管理。
    /// Start 在插件启动时（Word 主线程）调用；Kestrel 自身运行在线程池上，不阻塞 UI。
    /// Stop 在插件关闭时调用，优雅停止并释放资源。
    /// 安全约束：只监听 127.0.0.1 回环地址（McpSettings.GetHostForListening 强制收敛）。
    /// </summary>
    public sealed class McpHttpHost : IDisposable
    {
        private readonly McpSettings _settings;
        private readonly IMcpLogger _logger;
        private readonly IMcpRequestHandler _handler;
        private readonly ToolRegistry _toolRegistry;
        private IWebHost _host;

        public McpHttpHost(McpSettings settings, IMcpLogger logger, IMcpRequestHandler handler, ToolRegistry toolRegistry)
        {
            _settings = settings;
            _logger = logger;
            _handler = handler;
            _toolRegistry = toolRegistry;
        }

        /// <summary>是否已成功启动。</summary>
        public bool IsRunning => _host != null;

        /// <summary>启动 HTTP 服务；端口被占用等失败会抛出异常，由调用方记录日志。</summary>
        public void Start()
        {
            if (_host != null)
            {
                return;
            }

            // 注意：此处不使用 CreateDefaultBuilder —— 它是 netcoreapp 专用；
            // .NET Framework 上必须显式用 WebHostBuilder 逐项配置。
            var builder = new WebHostBuilder()
                .UseKestrel(options =>
                {
                    // 仅监听回环地址；即使配置被篡改，GetHostForListening 也会强制收敛
                    options.Listen(IPAddress.Loopback, _settings.Port);
                    // 防止超大请求体耗尽内存
                    options.Limits.MaxRequestBodySize = 1024 * 1024;
                })
                .UseStartup<Startup>()
                .ConfigureServices(services =>
                {
                    // 共享的单例服务注入 WebHost 的 DI 容器
                    services.AddSingleton(_settings);
                    services.AddSingleton<IMcpLogger>(_logger);
                    services.AddSingleton<IMcpRequestHandler>(_handler);
                    services.AddSingleton(_toolRegistry);
                });

            _host = builder.Build();
            _host.Start(); // Start 立即返回，Kestrel 在后台线程池接收请求

            _logger.Info($"HTTP 服务已启动，监听 http://{_settings.GetHostForListening()}:{_settings.Port}/mcp");
        }

        /// <summary>优雅停止并释放服务器。</summary>
        public void Stop()
        {
            if (_host == null)
            {
                return;
            }

            try
            {
                // 等待在途请求完成；GetAwaiter().GetResult() 避免同步等待死锁（UI 线程无同步上下文 await）
                _host.StopAsync().GetAwaiter().GetResult();
                _logger.Info("HTTP 服务已停止");
            }
            catch (Exception ex)
            {
                _logger.Warn($"停止 HTTP 服务时发生异常：{ex.Message}");
            }
            finally
            {
                _host.Dispose();
                _host = null;
            }
        }

        public void Dispose() => Stop();
    }
}
