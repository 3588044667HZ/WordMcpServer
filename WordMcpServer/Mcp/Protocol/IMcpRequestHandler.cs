using Newtonsoft.Json.Linq;

namespace WordMcpServer.Mcp.Protocol
{
    /// <summary>
    /// MCP（JSON-RPC）请求处理器：解析请求并返回响应。
    /// 返回 null 表示通知类消息（无需响应体）。
    /// </summary>
    public interface IMcpRequestHandler
    {
        JToken Process(string rawBody);
    }
}
