using System;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>
    /// 工具执行期间的可预期业务异常。携带错误码（见 McpErrorCodes）和人类可读消息，
    /// 由协议层转换为 MCP 的 isError 结果返回给客户端，而不是 JSON-RPC 错误。
    /// </summary>
    public sealed class McpToolException : Exception
    {
        public McpToolException(string code, string message) : base(message)
        {
            Code = code;
        }

        /// <summary>机器可读错误码（如 PPT_NOT_OPEN、SLIDE_OUT_OF_RANGE）。</summary>
        public string Code { get; }
    }
}
