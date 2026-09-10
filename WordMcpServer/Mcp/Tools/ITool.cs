using Newtonsoft.Json.Linq;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>
    /// MCP 工具抽象。每个工具提供名称、描述、JSON Schema 输入声明，
    /// 以及实际执行逻辑（内部通过 IWordService 访问 Word，线程已封送）。
    /// </summary>
    public interface ITool
    {
        /// <summary>工具名（MCP tools/call 中的 name）。</summary>
        string Name { get; }

        /// <summary>工具描述（供 AI 理解用途）。</summary>
        string Description { get; }

        /// <summary>输入参数的 JSON Schema（tools/list 中返回）。</summary>
        JObject InputSchema { get; }

        /// <summary>
        /// 执行工具。参数已通过 InputSchema 的 required 校验；
        /// 语义校验与范围检查在实现内部完成，失败时抛 McpToolException。
        /// 返回结构化的结果对象（JSON 序列化后作为 MCP text 内容返回）。
        /// </summary>
        JObject Execute(JObject arguments);
    }
}
