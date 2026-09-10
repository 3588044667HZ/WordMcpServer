using System;
using System.Diagnostics;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Mcp.Tools;

namespace WordMcpServer.Mcp.Protocol
{
    /// <summary>
    /// MCP 协议核心：JSON-RPC 2.0 消息处理。
    /// 支持方法：
    ///   initialize             —— 协议握手，返回能力声明
    ///   notifications/initialized —— 客户端初始化完成通知（无响应）
    ///   notifications/*        —— 其他通知（无响应）
    ///   ping                   —— 保活
    ///   tools/list             —— 返回全部工具定义（含 JSON Schema）
    ///   tools/call             —— 执行工具（失败返回 isError 结果而非 JSON-RPC 错误）
    /// 其余方法返回 JSON-RPC -32601。
    /// 注意：本处理器在 Kestrel 线程池线程上执行，Word 调用已由服务层封送到 UI 线程。
    /// </summary>
    public sealed class McpProtocolProcessor : IMcpRequestHandler
    {
        /// <summary>本服务器支持的 MCP 协议版本。</summary>
        private const string ProtocolVersion = "2025-06-18";

        private const string ServerName = "word-mcp-server";
        private const string ServerVersion = "1.0.0";

        private static readonly string[] KnownProtocolVersions =
        {
            "2024-11-05", "2025-03-26", "2025-06-18",
        };

        private readonly ToolRegistry _tools;
        private readonly IMcpLogger _logger;

        public McpProtocolProcessor(ToolRegistry tools, IMcpLogger logger)
        {
            _tools = tools;
            _logger = logger;
        }

        public JToken Process(string rawBody)
        {
            JToken request;
            try
            {
                request = JToken.Parse(rawBody);
            }
            catch (JsonException ex)
            {
                _logger.Warn($"JSON 解析失败：{ex.Message}");
                return ErrorResponse(null, -32700, "无效的 JSON 请求：解析错误");
            }

            if (request is JArray batch)
            {
                // JSON-RPC 批量请求：逐个处理（每个元素的结果或 null 通知占位）
                var responses = new JArray();
                foreach (var item in batch)
                {
                    var response = ProcessSingle(item);
                    if (response != null)
                    {
                        responses.Add(response);
                    }
                }

                return responses.Count > 0 ? (JToken)responses : null;
            }

            return ProcessSingle(request);
        }

        private JToken ProcessSingle(JToken item)
        {
            var req = item as JObject;
            if (req == null)
            {
                return ErrorResponse(null, -32600, "请求必须是 JSON 对象");
            }

            var id = req["id"]; // 通知请求没有 id
            var method = req["method"] as JValue;
            if (method == null || string.IsNullOrEmpty(method.Value<string>()))
            {
                return ErrorResponse(id, -32600, "缺少 method 字段");
            }

            var methodName = method.Value<string>();
            JObject parameters = req["params"] as JObject;
            if (req["params"] != null && parameters == null)
            {
                return ErrorResponse(id, -32600, "params 必须是对象");
            }

            _logger.Debug($"收到 MCP 请求：method={methodName}, id={id?.ToString() ?? "(notification)"}");

            switch (methodName)
            {
                case "initialize":
                    return Ok(id, new JObject
                    {
                        ["protocolVersion"] = NegotiateProtocolVersion(parameters),
                        ["capabilities"] = new JObject
                        {
                            ["tools"] = new JObject { ["listChanged"] = false },
                        },
                        ["serverInfo"] = new JObject
                        {
                            ["name"] = ServerName,
                            ["version"] = ServerVersion,
                        },
                    });

                // 客户端通知：按协议约定返回空（HTTP 层返回 202）
                case "notifications/initialized":
                case "notifications/cancelled":
                case "notifications/roots/list_changed":
                    return null;

                case "ping":
                    return Ok(id, new JObject());

                case "tools/list":
                    return Ok(id, new JObject { ["tools"] = _tools.ToDefinitions() });

                case "tools/call":
                    return HandleToolCall(id, parameters);

                default:
                    if (methodName.StartsWith("notifications/", StringComparison.Ordinal))
                    {
                        // 未知通知同样按协议静默处理
                        return null;
                    }

                    return ErrorResponse(id, -32601, $"未知方法：{methodName}");
            }
        }

        private JToken HandleToolCall(JToken id, JObject parameters)
        {
            if (parameters == null)
            {
                return ErrorResponse(id, -32602, "tools/call 缺少 params");
            }

            var name = parameters["name"] as JValue;
            if (name == null || string.IsNullOrEmpty(name.Value<string>()))
            {
                return ErrorResponse(id, -32602, "tools/call 缺少工具名称 name");
            }

            var tool = _tools.Find(name.Value<string>());
            if (tool == null)
            {
                return ErrorResponse(id, -32602, $"未知工具：{name.Value<string>()}",
                    new JObject { ["code"] = "UNKNOWN_TOOL" });
            }

            JObject arguments = parameters["arguments"] as JObject;
            if (parameters["arguments"] != null && arguments == null)
            {
                return ErrorResponse(id, -32602, "tools/call 的 arguments 必须是对象");
            }

            var stopwatch = Stopwatch.StartNew();
            try
            {
                var result = tool.Execute(arguments ?? new JObject());
                stopwatch.Stop();
                _logger.Info($"工具 {tool.Name} 执行成功，耗时 {stopwatch.ElapsedMilliseconds} ms");

                // 工具可通过 "$content" 返回自定义 MCP content 数组（如附带图片），
                // 否则默认把结果对象序列化为单个 text 内容
                JArray content;
                if (result["$content"] is JArray customContent)
                {
                    content = customContent;
                }
                else
                {
                    content = new JArray
                    {
                        new JObject
                        {
                            ["type"] = "text",
                            ["text"] = result.ToString(Formatting.None),
                        },
                    };
                }

                return Ok(id, new JObject
                {
                    ["content"] = content,
                    ["isError"] = false,
                });
            }
            catch (McpToolException ex)
            {
                stopwatch.Stop();
                _logger.Warn($"工具 {tool.Name} 执行失败：{ex.Code} - {ex.Message}");
                return Ok(id, ErrorToolResult(ex.Code, ex.Message));
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                _logger.Error($"工具 {tool.Name} 抛出未预期异常", ex);
                return Ok(id, ErrorToolResult(McpErrorCodes.InternalError, "执行工具时发生内部错误：" + ex.Message));
            }
        }

        /// <summary>工具失败结果：MCP 规范要求通过 isError=true 的 result 返回。</summary>
        private static JObject ErrorToolResult(string code, string message) => new JObject
        {
            ["content"] = new JArray
            {
                new JObject
                {
                    ["type"] = "text",
                    ["text"] = new JObject
                    {
                        ["error"] = new JObject { ["code"] = code, ["message"] = message },
                    }.ToString(Formatting.None),
                },
            },
            ["isError"] = true,
        };

        private static JToken Ok(JToken id, JObject result) => new JObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id,
            ["result"] = result,
        };

        private static JToken ErrorResponse(JToken id, int code, string message, JObject data = null)
        {
            var error = new JObject
            {
                ["code"] = code,
                ["message"] = message,
            };
            if (data != null)
            {
                error["data"] = data;
            }

            return new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id,
                ["error"] = error,
            };
        }

        /// <summary>
        /// 协议版本协商：若客户端请求的版本在本服务器已知集合内则沿用，
        /// 否则返回本服务器支持的最新版本（客户端会据此降级）。
        /// </summary>
        private static string NegotiateProtocolVersion(JObject parameters)
        {
            if (parameters != null)
            {
                var requested = parameters["protocolVersion"] as JValue;
                if (requested != null)
                {
                    var text = requested.Value<string>();
                    foreach (var known in KnownProtocolVersions)
                    {
                        if (string.Equals(known, text, StringComparison.Ordinal))
                        {
                            return known;
                        }
                    }
                }
            }

            return ProtocolVersion;
        }
    }
}
