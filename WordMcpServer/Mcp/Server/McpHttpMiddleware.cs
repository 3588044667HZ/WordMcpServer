using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Mcp.Protocol;

namespace WordMcpServer.Mcp.Server
{
    /// <summary>
    /// MCP Streamable HTTP 传输的 HTTP 层：
    ///   POST /mcp  —— JSON-RPC 请求（响应为 application/json；若客户端 Accept
    ///                 text/event-stream 则用 SSE 单事件包装返回）
    ///   POST /     —— 与 /mcp 等效（便于手工测试）
    ///   GET  /health —— 健康检查
    /// 通知类消息返回 202 空响应（协议约定）。
    /// </summary>
    public sealed class McpHttpMiddleware
    {
        private readonly RequestDelegate _next;
        private readonly IMcpRequestHandler _handler;
        private readonly IMcpLogger _logger;

        public McpHttpMiddleware(RequestDelegate next, IMcpRequestHandler handler, IMcpLogger logger)
        {
            _next = next;
            _handler = handler;
            _logger = logger;
        }

        public async Task Invoke(HttpContext context)
        {
            var path = context.Request.Path.Value ?? string.Empty;

            if (string.Equals(path, "/health", StringComparison.OrdinalIgnoreCase))
            {
                context.Response.ContentType = "text/plain; charset=utf-8";
                await context.Response.WriteAsync("ok");
                return;
            }

            if (!string.Equals(path, "/mcp", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(path, "/", StringComparison.OrdinalIgnoreCase))
            {
                context.Response.StatusCode = StatusCodes.Status404NotFound;
                await context.Response.WriteAsync("{\"error\":\"not_found\"}");
                return;
            }

            if (!HttpMethods.IsPost(context.Request.Method))
            {
                context.Response.StatusCode = StatusCodes.Status405MethodNotAllowed;
                await context.Response.WriteAsync("{\"error\":\"method_not_allowed\"}");
                return;
            }

            string rawBody;
            try
            {
                using (var reader = new StreamReader(context.Request.Body, Encoding.UTF8))
                {
                    rawBody = await reader.ReadToEndAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取请求体失败：{ex.Message}");
                context.Response.StatusCode = StatusCodes.Status400BadRequest;
                await context.Response.WriteAsync("{\"error\":\"bad_request\"}");
                return;
            }

            _logger.Info($"收到请求：{context.Connection.RemoteIpAddress}:{context.Connection.RemotePort} -> {path}，长度 {rawBody?.Length ?? 0}");

            var stopwatch = Stopwatch.StartNew();
            JToken response;
            try
            {
                response = _handler.Process(rawBody);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                _logger.Error("处理 MCP 请求时抛出未预期异常", ex);
                response = new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = null,
                    ["error"] = new JObject
                    {
                        ["code"] = -32603,
                        ["message"] = "服务器内部错误",
                    },
                };
            }

            stopwatch.Stop();
            _logger.Info($"请求处理完成，耗时 {stopwatch.ElapsedMilliseconds} ms");

            if (response == null)
            {
                // 通知类消息：无响应体
                context.Response.StatusCode = StatusCodes.Status202Accepted;
                return;
            }

            var json = response.ToString(Formatting.None);
            var accept = context.Request.Headers["Accept"].ToString() ?? string.Empty;
            if (accept.IndexOf("text/event-stream", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // SSE 单事件封装：协议允许服务器以事件流形式返回单个响应
                context.Response.ContentType = "text/event-stream; charset=utf-8";
                context.Response.Headers["Cache-Control"] = "no-cache";
                var sb = new StringBuilder();
                sb.Append("event: message\n");
                sb.Append("data: ").Append(json).Append("\n\n");
                await context.Response.WriteAsync(sb.ToString());
            }
            else
            {
                context.Response.ContentType = "application/json; charset=utf-8";
                await context.Response.WriteAsync(json);
            }
        }
    }
}
