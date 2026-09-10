using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;

namespace WordMcpServer.Mcp.Server
{
    /// <summary>
    /// ASP.NET Core 启动配置。业务服务已通过 WebHostBuilder.ConfigureServices 注入，
    /// 这里只装配中间件管线（本服务器仅需一个自定义 MCP 中间件）。
    /// </summary>
    public sealed class Startup
    {
        public void ConfigureServices(IServiceCollection services)
        {
            // 无需额外注册；服务由 McpHttpHost.ConfigureServices 注入
        }

        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            app.UseMiddleware<McpHttpMiddleware>();
        }
    }
}
