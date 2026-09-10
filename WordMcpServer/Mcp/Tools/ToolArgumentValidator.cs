using System;
using System.Drawing;
using Newtonsoft.Json.Linq;
using WordMcpServer.Mcp;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>
    /// 工具参数运行时校验工具集。
    /// JSON Schema 用于 tools/list 声明；这里是实际执行前的强校验，
    /// 保证进入 Word 对象模型之前参数已经合法（边界防御）。
    /// 所有方法在校验失败时抛 McpToolException(INVALID_PARAMS, ...)。
    /// </summary>
    public static class ToolArgumentValidator
    {
        /// <summary>要求参数对象非空。</summary>
        public static JObject RequireObject(JToken arguments)
        {
            var obj = arguments as JObject;
            if (obj == null)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, "参数必须是 JSON 对象（arguments）");
            }

            return obj;
        }

        /// <summary>必填整型参数，可附带取值范围。</summary>
        public static int RequireInt(JObject args, string name, int? min = null, int? max = null)
        {
            var token = args[name];
            int value;
            if (token == null || token.Type == JTokenType.Null)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"缺少必填参数：{name}");
            }

            if (token.Type == JTokenType.Integer)
            {
                value = token.Value<int>();
            }
            else if (token.Type == JTokenType.Float)
            {
                // 允许 1.0 这类整数形式的浮点数
                var d = token.Value<double>();
                if (Math.Abs(d - Math.Round(d)) > 0.0001)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 必须是整数");
                }

                value = (int)Math.Round(d);
            }
            else
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 必须是整数");
            }

            if (min.HasValue && value < min.Value)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 不能小于 {min.Value}");
            }

            if (max.HasValue && value > max.Value)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 不能大于 {max.Value}");
            }

            return value;
        }

        /// <summary>可选整型参数（缺省返回 null）。</summary>
        public static int? OptionalInt(JObject args, string name, int? min = null, int? max = null)
        {
            if (args[name] == null || args[name].Type == JTokenType.Null)
            {
                return null;
            }

            return RequireInt(args, name, min, max);
        }

        /// <summary>必填字符串参数（去除首尾空白）。</summary>
        public static string RequireString(JObject args, string name)
        {
            var value = OptionalString(args, name);
            if (string.IsNullOrEmpty(value))
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"缺少必填参数：{name}");
            }

            return value;
        }

        /// <summary>可选字符串参数（缺省或 null 返回 null）。</summary>
        public static string OptionalString(JObject args, string name)
        {
            var token = args[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            if (token.Type != JTokenType.String)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 必须是字符串");
            }

            return token.Value<string>().Trim();
        }

        /// <summary>可选布尔参数。</summary>
        public static bool? OptionalBool(JObject args, string name)
        {
            var token = args[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            if (token.Type != JTokenType.Boolean)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 必须是布尔值");
            }

            return token.Value<bool>();
        }

        /// <summary>可选浮点参数，附带范围校验。</summary>
        public static double? OptionalDouble(JObject args, string name, double? min = null, double? max = null)
        {
            var token = args[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            double value;
            if (token.Type == JTokenType.Integer || token.Type == JTokenType.Float)
            {
                value = token.Value<double>();
            }
            else
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 必须是数值");
            }

            if (min.HasValue && value < min.Value)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 不能小于 {min.Value}");
            }

            if (max.HasValue && value > max.Value)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 不能大于 {max.Value}");
            }

            return value;
        }

        /// <summary>
        /// 可选颜色参数。支持 "#RRGGBB"、八位 "#AARRGGBB" 或颜色名（如 "red"、"蓝"）。
        /// 返回 ARGB 整型；Word 的 WdColor（BGR 顺序）转换在服务层完成。
        /// </summary>
        public static int? OptionalColor(JObject args, string name)
        {
            var text = OptionalString(args, name);
            if (text == null)
            {
                return null;
            }

            return ParseColorString(text, name);
        }

        /// <summary>解析颜色字符串为 ARGB 整型（供颜色参数共用）。</summary>
        public static int ParseColorString(string text, string parameterName)
        {
            if (text.StartsWith("#", StringComparison.Ordinal))
            {
                var hex = text.Substring(1);
                if (hex.Length == 6)
                {
                    if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out var rgb))
                    {
                        // 补全 alpha 通道
                        return unchecked((int)0xFF000000) | rgb;
                    }
                }
                else if (hex.Length == 8)
                {
                    if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out var argb))
                    {
                        return argb;
                    }
                }
            }
            else
            {
                var known = Color.FromName(text);
                if (known.IsKnownColor)
                {
                    return known.ToArgb();
                }
            }

            throw new McpToolException(McpErrorCodes.InvalidParams,
                $"参数 {parameterName} 不是合法的颜色值：\"{text}\"（支持 #RRGGBB、#AARRGGBB 或颜色名）");
        }

        /// <summary>
        /// 可选富文本 runs 参数：[{text, font_name, font_size, color, bold, italic, underline, highlight}]。
        /// 返回服务层 TextRunSpec 列表；未提供返回 null。
        /// </summary>
        public static System.Collections.Generic.List<Services.TextRunSpec> OptionalRuns(JObject args, string name)
        {
            var token = args[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            if (token.Type != JTokenType.Array)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 必须是数组");
            }

            var runs = new System.Collections.Generic.List<Services.TextRunSpec>();
            var index = 0;
            foreach (var item in (JArray)token)
            {
                index++;
                var runObj = item as JObject;
                if (runObj == null)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, $"{name}[{index}] 必须是对象");
                }

                var textToken = runObj["text"];
                if (textToken == null || textToken.Type == JTokenType.Null)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, $"{name}[{index}] 缺少 text");
                }

                var run = new Services.TextRunSpec { Text = textToken.ToString() };

                var fontName = runObj["font_name"];
                if (fontName != null && fontName.Type != JTokenType.Null)
                {
                    run.FontName = fontName.Value<string>();
                }

                var fontSize = runObj["font_size"];
                if (fontSize != null && fontSize.Type != JTokenType.Null)
                {
                    var size = fontSize.Value<double>();
                    if (size < 1 || size > 400)
                    {
                        throw new McpToolException(McpErrorCodes.InvalidParams, $"{name}[{index}].font_size 必须在 1-400 之间");
                    }

                    run.FontSize = size;
                }

                var color = runObj["color"];
                if (color != null && color.Type != JTokenType.Null)
                {
                    run.ColorArgb = ParseColorString(color.Value<string>(), $"{name}[{index}].color");
                }

                var bold = runObj["bold"];
                if (bold != null && bold.Type != JTokenType.Null)
                {
                    run.Bold = bold.Value<bool>();
                }

                var italic = runObj["italic"];
                if (italic != null && italic.Type != JTokenType.Null)
                {
                    run.Italic = italic.Value<bool>();
                }

                var underline = runObj["underline"];
                if (underline != null && underline.Type != JTokenType.Null)
                {
                    run.Underline = underline.Value<bool>();
                }

                var highlight = runObj["highlight"];
                if (highlight != null && highlight.Type != JTokenType.Null)
                {
                    run.Highlight = highlight.Value<string>();
                }

                runs.Add(run);
            }

            return runs;
        }

        /// <summary>可选字符串键值对（对象形式，值必须为字符串或可转字符串的标量）。</summary>
        public static System.Collections.Generic.Dictionary<string, string> OptionalStringDictionary(JObject args, string name)
        {
            var token = args[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            var obj = token as JObject;
            if (obj == null)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 必须是对象（键值对）");
            }

            var result = new System.Collections.Generic.Dictionary<string, string>();
            foreach (var property in obj.Properties())
            {
                if (property.Value == null || property.Value.Type == JTokenType.Null)
                {
                    result[property.Name] = string.Empty;
                    continue;
                }

                if (property.Value is JObject || property.Value is JArray)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams,
                        $"参数 {name}.{property.Name} 必须是字符串");
                }

                result[property.Name] = property.Value.ToString();
            }

            return result;
        }

        /// <summary>可选字符串数组（必须是字符串组成的数组）。</summary>
        public static System.Collections.Generic.List<string> OptionalStringArray(JObject args, string name)
        {
            var token = args[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            if (token.Type != JTokenType.Array)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 必须是字符串数组");
            }

            var result = new System.Collections.Generic.List<string>();
            foreach (var item in (JArray)token)
            {
                if (item.Type != JTokenType.String)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 的数组元素必须是字符串");
                }

                result.Add(item.Value<string>());
            }

            return result;
        }

        /// <summary>可选数值数组（如列宽）。</summary>
        public static System.Collections.Generic.List<double> OptionalDoubleArray(JObject args, string name)
        {
            var token = args[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            if (token.Type != JTokenType.Array)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 必须是数值数组");
            }

            var result = new System.Collections.Generic.List<double>();
            foreach (var item in (JArray)token)
            {
                if (item.Type != JTokenType.Integer && item.Type != JTokenType.Float)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 的数组元素必须是数值");
                }

                result.Add(item.Value<double>());
            }

            return result;
        }

        /// <summary>
        /// 可选字符串矩阵（如表格数据）：必须是二维字符串数组，行长度可不一致（部分更新语义）。
        /// </summary>
        public static System.Collections.Generic.List<System.Collections.Generic.List<string>> OptionalStringMatrix(JObject args, string name)
        {
            var token = args[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            if (token.Type != JTokenType.Array)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name} 必须是二维字符串数组");
            }

            var rows = new System.Collections.Generic.List<System.Collections.Generic.List<string>>();
            var rowIndex = 0;
            foreach (var rowToken in (JArray)token)
            {
                rowIndex++;
                if (rowToken.Type != JTokenType.Array)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name}[{rowIndex}] 必须是字符串数组");
                }

                var row = new System.Collections.Generic.List<string>();
                foreach (var cellToken in (JArray)rowToken)
                {
                    if (cellToken.Type != JTokenType.String && cellToken.Type != JTokenType.Null &&
                        cellToken.Type != JTokenType.Integer && cellToken.Type != JTokenType.Float &&
                        cellToken.Type != JTokenType.Boolean)
                    {
                        throw new McpToolException(McpErrorCodes.InvalidParams, $"参数 {name}[{rowIndex}] 的单元格必须是字符串");
                    }

                    row.Add(cellToken.Type == JTokenType.Null ? string.Empty : cellToken.ToString());
                }

                rows.Add(row);
            }

            return rows;
        }
    }
}
