using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>
    /// 工具基类：收敛 Name/Description/InputSchema/Execute 的样板代码，
    /// 并提供 JSON Schema 片段构造辅助方法（schema 统一 snake_case）。
    /// </summary>
    public abstract class ToolBase : ITool
    {
        protected ToolBase(IWordService service, IMcpLogger logger)
        {
            Service = service;
            Logger = logger;
        }

        protected IWordService Service { get; }

        protected IMcpLogger Logger { get; }

        public abstract string Name { get; }

        public abstract string Description { get; }

        public abstract JObject InputSchema { get; }

        public abstract JObject Execute(JObject arguments);

        protected static JObject EmptySchema() => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject(),
            ["required"] = new JArray(),
        };

        protected static JObject Schema(JObject properties, params string[] required) => new JObject
        {
            ["type"] = "object",
            ["properties"] = properties,
            ["required"] = new JArray(required),
        };

        protected static JObject Str(string description, params string[] enums)
        {
            var prop = new JObject
            {
                ["type"] = "string",
                ["description"] = description,
            };
            if (enums != null && enums.Length > 0)
            {
                prop["enum"] = new JArray(enums);
            }

            return prop;
        }

        protected static JObject Int(string description, int? min = null, int? max = null)
        {
            var prop = new JObject
            {
                ["type"] = "integer",
                ["description"] = description,
            };
            if (min.HasValue) prop["minimum"] = min.Value;
            if (max.HasValue) prop["maximum"] = max.Value;
            return prop;
        }

        protected static JObject Num(string description, double? min = null, double? max = null)
        {
            var prop = new JObject
            {
                ["type"] = "number",
                ["description"] = description,
            };
            if (min.HasValue) prop["minimum"] = min.Value;
            if (max.HasValue) prop["maximum"] = max.Value;
            return prop;
        }

        protected static JObject Bool(string description) => new JObject
        {
            ["type"] = "boolean",
            ["description"] = description,
        };

        protected static JObject StrArray(string description) => new JObject
        {
            ["type"] = "array",
            ["description"] = description,
            ["items"] = new JObject { ["type"] = "string" },
        };

        protected static JObject NumArray(string description) => new JObject
        {
            ["type"] = "array",
            ["description"] = description,
            ["items"] = new JObject { ["type"] = "number" },
        };

        protected static JObject StrMatrix(string description) => new JObject
        {
            ["type"] = "array",
            ["description"] = description,
            ["items"] = new JObject
            {
                ["type"] = "array",
                ["items"] = new JObject { ["type"] = "string" },
            },
        };
    }
}
