using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>添加段落。</summary>
    public sealed class AddParagraphTool : ToolBase
    {
        public AddParagraphTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "add_paragraph";

        public override string Description => "在文档末尾追加一个段落（复用末尾空段落），或插入到指定段落之前。可选 style 套用段落样式（如 heading1、normal、quote 或样式名）。返回新段落的索引，可用于后续格式化。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("段落文本；文本内的换行会转为 Word 手动换行符"),
            ["position"] = Int("插入位置（1 基段落索引，插入到该段落之前）；缺省追加到文档末尾", 1),
            ["style"] = Str("段落样式（可选）：heading1-heading9 / normal / title / subtitle / quote / list-paragraph / caption 或样式名"),
        }, "text");

        public override JObject Execute(JObject arguments)
        {
            var request = new AddParagraphRequest
            {
                Text = ToolArgumentValidator.RequireString(arguments, "text"),
                Position = ToolArgumentValidator.OptionalInt(arguments, "position", 1),
                Style = ToolArgumentValidator.OptionalString(arguments, "style"),
            };
            Logger.Debug($"执行 add_paragraph，position={request.Position}，style={request.Style}");
            return JObject.FromObject(Service.AddParagraph(request));
        }
    }

    /// <summary>添加标题。</summary>
    public sealed class AddHeadingTool : ToolBase
    {
        public AddHeadingTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "add_heading";

        public override string Description => "在文档末尾（或指定段落之前）添加一个标题段落，自动套用 Heading 1-9 样式（级别与语言无关）。返回新段落的索引。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("标题文本"),
            ["level"] = Int("标题级别 1-9（1 为最高级）", 1, 9),
            ["position"] = Int("插入位置（1 基段落索引，插入到该段落之前）；缺省追加到文档末尾", 1),
        }, "text", "level");

        public override JObject Execute(JObject arguments)
        {
            var text = ToolArgumentValidator.RequireString(arguments, "text");
            var level = ToolArgumentValidator.RequireInt(arguments, "level", 1, 9);
            var position = ToolArgumentValidator.OptionalInt(arguments, "position", 1);
            Logger.Debug($"执行 add_heading，level={level}，position={position}");
            return JObject.FromObject(Service.AddHeading(text, level, position));
        }
    }

    /// <summary>设置段落格式。</summary>
    public sealed class SetParagraphFormatTool : ToolBase
    {
        public SetParagraphFormatTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_paragraph_format";

        public override string Description => "设置段落格式（未提供的参数保持不变）。用 paragraph_index 指定段落，或用 find_text 定位包含该文本的第一个段落。支持对齐、缩进、段前段后间距、行距、分页控制与大纲级别。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["paragraph_index"] = Int("段落索引（1 基）", 1),
            ["find_text"] = Str("按文本定位段落（包含该文本的第一个段落）"),
            ["alignment"] = Str("水平对齐", "left", "center", "right", "justify", "distribute"),
            ["left_indent"] = Num("左缩进（磅）"),
            ["right_indent"] = Num("右缩进（磅）"),
            ["first_line_indent"] = Num("首行缩进（磅，可为负表示悬挂缩进）"),
            ["space_before"] = Num("段前间距（磅）", 0),
            ["space_after"] = Num("段后间距（磅）", 0),
            ["line_spacing"] = Str("行距模式", "single", "1.5", "double", "at-least", "exactly", "multiple"),
            ["line_spacing_value"] = Num("行距数值（磅）：at-least / exactly / multiple 时必填"),
            ["keep_with_next"] = Bool("与下段同页"),
            ["keep_together"] = Bool("段中不分页"),
            ["page_break_before"] = Bool("段前分页"),
            ["outline_level"] = Str("大纲级别：body-text / level-1 .. level-9"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetParagraphFormatRequest
            {
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
                FindText = ToolArgumentValidator.OptionalString(arguments, "find_text"),
                Alignment = ToolArgumentValidator.OptionalString(arguments, "alignment"),
                LeftIndent = ToolArgumentValidator.OptionalDouble(arguments, "left_indent"),
                RightIndent = ToolArgumentValidator.OptionalDouble(arguments, "right_indent"),
                FirstLineIndent = ToolArgumentValidator.OptionalDouble(arguments, "first_line_indent"),
                SpaceBefore = ToolArgumentValidator.OptionalDouble(arguments, "space_before", 0),
                SpaceAfter = ToolArgumentValidator.OptionalDouble(arguments, "space_after", 0),
                LineSpacing = ToolArgumentValidator.OptionalString(arguments, "line_spacing"),
                LineSpacingValue = ToolArgumentValidator.OptionalDouble(arguments, "line_spacing_value", 0),
                KeepWithNext = ToolArgumentValidator.OptionalBool(arguments, "keep_with_next"),
                KeepTogether = ToolArgumentValidator.OptionalBool(arguments, "keep_together"),
                PageBreakBefore = ToolArgumentValidator.OptionalBool(arguments, "page_break_before"),
                OutlineLevel = ToolArgumentValidator.OptionalString(arguments, "outline_level"),
            };
            Logger.Debug($"执行 set_paragraph_format，paragraphIndex={request.ParagraphIndex}，findText={request.FindText}");
            return JObject.FromObject(Service.SetParagraphFormat(request));
        }
    }

    /// <summary>设置文本样式（字体/字号/加粗/颜色等）。</summary>
    public sealed class SetTextStyleTool : ToolBase
    {
        public SetTextStyleTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_text_style";

        public override string Description => "设置文本样式（未提供的参数保持不变）。用 paragraph_index 指定段落，或用 find_text 对所有匹配文本应用样式。支持字体、字号、加粗、斜体、下划线、颜色、全部大写、上下标。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["paragraph_index"] = Int("段落索引（1 基）", 1),
            ["find_text"] = Str("查找文本：对该文本的所有出现位置应用样式"),
            ["font_name"] = Str("字体名称（同时设置中英文字体），如 宋体 / 微软雅黑 / Arial"),
            ["font_size"] = Num("字号（磅，1-400）", 1, 400),
            ["bold"] = Bool("加粗"),
            ["italic"] = Bool("斜体"),
            ["underline"] = Bool("下划线"),
            ["color"] = Str("字体颜色：#RRGGBB、#AARRGGBB 或颜色名（red / blue / 红 等）"),
            ["highlight"] = Str("突出显示颜色", "yellow", "bright-green", "turquoise", "pink", "blue", "red", "dark-blue", "teal", "violet", "dark-red", "dark-yellow", "gray", "gray-25", "black", "white", "none"),
            ["all_caps"] = Bool("全部大写"),
            ["subscript"] = Bool("下标"),
            ["superscript"] = Bool("上标"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetTextStyleRequest
            {
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
                FindText = ToolArgumentValidator.OptionalString(arguments, "find_text"),
                FontName = ToolArgumentValidator.OptionalString(arguments, "font_name"),
                FontSize = ToolArgumentValidator.OptionalDouble(arguments, "font_size", 1, 400),
                Bold = ToolArgumentValidator.OptionalBool(arguments, "bold"),
                Italic = ToolArgumentValidator.OptionalBool(arguments, "italic"),
                Underline = ToolArgumentValidator.OptionalBool(arguments, "underline"),
                ColorArgb = ToolArgumentValidator.OptionalColor(arguments, "color"),
                Highlight = ToolArgumentValidator.OptionalString(arguments, "highlight"),
                AllCaps = ToolArgumentValidator.OptionalBool(arguments, "all_caps"),
                Subscript = ToolArgumentValidator.OptionalBool(arguments, "subscript"),
                Superscript = ToolArgumentValidator.OptionalBool(arguments, "superscript"),
            };
            Logger.Debug($"执行 set_text_style，paragraphIndex={request.ParagraphIndex}，findText={request.FindText}");
            return JObject.FromObject(Service.SetTextStyle(request));
        }
    }

    /// <summary>插入分页符。</summary>
    public sealed class InsertPageBreakTool : ToolBase
    {
        public InsertPageBreakTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_page_break";

        public override string Description => "插入分页符：before_paragraph 指定在哪个段落之前分页（缺省在文档末尾分页）。可选 text 在分页符后写入文本。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["before_paragraph"] = Int("在该段落之前分页（1 基）；缺省在文档末尾", 1),
            ["text"] = Str("分页符之后要写入的文本（可选）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var before = ToolArgumentValidator.OptionalInt(arguments, "before_paragraph", 1);
            var text = ToolArgumentValidator.OptionalString(arguments, "text");
            Logger.Debug($"执行 insert_page_break，beforeParagraph={before}");
            return JObject.FromObject(Service.InsertPageBreak(before, text));
        }
    }

    /// <summary>插入分节符。</summary>
    public sealed class InsertSectionBreakTool : ToolBase
    {
        public InsertSectionBreakTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_section_break";

        public override string Description => "插入分节符（下一页分节）：before_paragraph 指定在哪个段落之前分节（缺省在文档末尾）。可选 text 在分节符后写入文本。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["before_paragraph"] = Int("在该段落之前分节（1 基）；缺省在文档末尾", 1),
            ["text"] = Str("分节符之后要写入的文本（可选）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var before = ToolArgumentValidator.OptionalInt(arguments, "before_paragraph", 1);
            var text = ToolArgumentValidator.OptionalString(arguments, "text");
            Logger.Debug($"执行 insert_section_break，beforeParagraph={before}");
            return JObject.FromObject(Service.InsertSectionBreak(before, text));
        }
    }

    /// <summary>应用段落样式。</summary>
    public sealed class ApplyStyleTool : ToolBase
    {
        public ApplyStyleTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "apply_style";

        public override string Description => "对段落应用样式：paragraph_index 指定段落、find_text 匹配所有出现位置、all=true 应用到整篇文档。style 支持内置别名（heading1-9 / normal / title / subtitle / quote / intense-quote / list-paragraph / caption / strong / emphasis）或样式名。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["style"] = Str("样式：heading1-heading9 / normal / title / subtitle / quote / intense-quote / list-paragraph / caption / strong / emphasis 或样式名"),
            ["paragraph_index"] = Int("段落索引（1 基）", 1),
            ["find_text"] = Str("按文本匹配（对所有出现位置应用样式）"),
            ["all"] = Bool("是否应用到整篇文档的所有段落"),
        }, "style");

        public override JObject Execute(JObject arguments)
        {
            var request = new ApplyStyleRequest
            {
                Style = ToolArgumentValidator.RequireString(arguments, "style"),
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
                FindText = ToolArgumentValidator.OptionalString(arguments, "find_text"),
                All = ToolArgumentValidator.OptionalBool(arguments, "all"),
            };
            Logger.Debug($"执行 apply_style，style={request.Style}，paragraphIndex={request.ParagraphIndex}，all={request.All}");
            return JObject.FromObject(Service.ApplyStyle(request));
        }
    }

    /// <summary>列出文档可用样式。</summary>
    public sealed class ListStylesTool : ToolBase
    {
        public ListStylesTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "list_styles";

        public override string Description => "列出当前文档可用的样式（名称、类型、是否内置、是否使用中）。可用返回的 name 作为 apply_style / add_paragraph 的 style 参数。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 list_styles");
            var styles = Service.ListStyles();
            return new JObject
            {
                ["count"] = styles.Count,
                ["styles"] = JArray.FromObject(styles),
            };
        }
    }

    /// <summary>在光标处插入文本。</summary>
    public sealed class InsertTextAtCursorTool : ToolBase
    {
        public InsertTextAtCursorTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_text_at_cursor";

        public override string Description => "在 Word 当前光标位置插入文本（使用 Selection.TypeText），可选 style 套用样式。适合模拟继续输入的场景。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("要插入的文本"),
            ["style"] = Str("样式（可选）"),
        }, "text");

        public override JObject Execute(JObject arguments)
        {
            var text = ToolArgumentValidator.RequireString(arguments, "text");
            var style = ToolArgumentValidator.OptionalString(arguments, "style");
            Logger.Debug("执行 insert_text_at_cursor");
            return JObject.FromObject(Service.InsertTextAtCursor(text, style));
        }
    }

    /// <summary>撤销最近操作。</summary>
    public sealed class UndoLastTool : ToolBase
    {
        public UndoLastTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "undo_last";

        public override string Description => "撤销当前文档最近的若干次操作（count 默认 1）。每个 MCP 写操作都作为独立撤销单元，可整批撤销。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["count"] = Int("撤销次数（默认 1）", 1, 100),
        });

        public override JObject Execute(JObject arguments)
        {
            var count = ToolArgumentValidator.OptionalInt(arguments, "count", 1, 100) ?? 1;
            Logger.Debug($"执行 undo_last，count={count}");
            var undone = Service.UndoLast(count);
            return new JObject
            {
                ["undone"] = undone,
                ["message"] = undone == 0 ? "没有可撤销的操作" : $"已撤销 {undone} 次操作",
            };
        }
    }

    /// <summary>读取单个段落的完整信息（含格式）。</summary>
    public sealed class GetParagraphTool : ToolBase
    {
        public GetParagraphTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_paragraph";

        public override string Description => "读取指定段落的完整信息：文本、样式、对齐、大纲级别、缩进、段前段后、行距、字体/字号/粗斜体/颜色/突出显示、列表类型与级别、是否在表格内。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["paragraph_index"] = Int("段落索引（1 基）", 1),
        }, "paragraph_index");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "paragraph_index", 1);
            Logger.Debug($"执行 get_paragraph，paragraphIndex={index}");
            return JObject.FromObject(Service.GetParagraph(index));
        }
    }

    /// <summary>删除段落。</summary>
    public sealed class DeleteParagraphTool : ToolBase
    {
        public DeleteParagraphTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_paragraph";

        public override string Description => "删除指定段落（含段落内容；文档始终保留至少一个段落）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["paragraph_index"] = Int("段落索引（1 基）", 1),
        }, "paragraph_index");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "paragraph_index", 1);
            Logger.Debug($"执行 delete_paragraph，paragraphIndex={index}");
            return JObject.FromObject(Service.DeleteParagraph(index));
        }
    }

    /// <summary>删除文本。</summary>
    public sealed class DeleteTextTool : ToolBase
    {
        public DeleteTextTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_text";

        public override string Description => "按文本删除：删除所有匹配文本，或只删除第 N 处（occurrence）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("要删除的文本"),
            ["occurrence"] = Int("只删除第 N 处匹配（1 基）；缺省删除全部", 1),
            ["match_case"] = Bool("区分大小写，默认 false"),
            ["match_whole_word"] = Bool("全字匹配，默认 false"),
        }, "text");

        public override JObject Execute(JObject arguments)
        {
            var text = ToolArgumentValidator.RequireString(arguments, "text");
            var occurrence = ToolArgumentValidator.OptionalInt(arguments, "occurrence", 1);
            var matchCase = ToolArgumentValidator.OptionalBool(arguments, "match_case") ?? false;
            var matchWholeWord = ToolArgumentValidator.OptionalBool(arguments, "match_whole_word") ?? false;
            Logger.Debug($"执行 delete_text，text={text}，occurrence={occurrence}");
            return JObject.FromObject(Service.DeleteText(text, occurrence, matchCase, matchWholeWord));
        }
    }

    /// <summary>添加富文本段落（同一段落多段不同格式）。</summary>
    public sealed class AddParagraphRunsTool : ToolBase
    {
        public AddParagraphRunsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "add_paragraph_runs";

        public override string Description => "添加一个富文本段落：runs 数组中每个片段可独立设置字体/字号/颜色/粗斜体/下划线/突出显示，适合“一段话里混排多种格式”的场景。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["runs"] = new JObject
            {
                ["type"] = "array",
                ["description"] = "文本片段数组，每项：{text, font_name, font_size, color, bold, italic, underline, highlight}",
                ["items"] = new JObject
                {
                    ["type"] = "object",
                    ["properties"] = new JObject
                    {
                        ["text"] = Str("片段文本"),
                        ["font_name"] = Str("字体名称"),
                        ["font_size"] = Num("字号（磅）", 1, 400),
                        ["color"] = Str("字体颜色：#RRGGBB / 颜色名"),
                        ["bold"] = Bool("加粗"),
                        ["italic"] = Bool("斜体"),
                        ["underline"] = Bool("下划线"),
                        ["highlight"] = Str("突出显示颜色"),
                    },
                    ["required"] = new JArray { "text" },
                },
            },
            ["position"] = Int("插入位置（1 基段落索引，插入到该段落之前）；缺省追加到文档末尾", 1),
            ["style"] = Str("段落样式（可选）"),
        }, "runs");

        public override JObject Execute(JObject arguments)
        {
            var request = new AddParagraphRunsRequest
            {
                Runs = ToolArgumentValidator.OptionalRuns(arguments, "runs"),
                Position = ToolArgumentValidator.OptionalInt(arguments, "position", 1),
                Style = ToolArgumentValidator.OptionalString(arguments, "style"),
            };
            if (request.Runs == null || request.Runs.Count == 0)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：runs（至少一个片段）");
            }

            Logger.Debug($"执行 add_paragraph_runs，runs={request.Runs.Count}");
            return JObject.FromObject(Service.AddParagraphRuns(request));
        }
    }

    /// <summary>设置列表格式（项目符号 / 编号 / 多级列表）。</summary>
    public sealed class SetListFormatTool : ToolBase
    {
        public SetListFormatTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_list_format";

        public override string Description => "把段落设置为列表：bullet（项目符号）/ number（编号）/ outline（多级列表）/ none（取消列表）。用 paragraph_index 指定单段，或用 start_paragraph+end_paragraph 指定区间；缺省为最后一个段落。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["type"] = Str("列表类型", "bullet", "number", "outline", "none"),
            ["paragraph_index"] = Int("单个目标段落（1 基，可选）", 1),
            ["start_paragraph"] = Int("起始段落（1 基，可选）", 1),
            ["end_paragraph"] = Int("结束段落（1 基，含，可选）", 1),
            ["level"] = Int("列表级别（1-9，默认 1）", 1, 9),
        }, "type");

        public override JObject Execute(JObject arguments)
        {
            var request = new SetListFormatRequest
            {
                Type = ToolArgumentValidator.RequireString(arguments, "type"),
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
                StartParagraph = ToolArgumentValidator.OptionalInt(arguments, "start_paragraph", 1),
                EndParagraph = ToolArgumentValidator.OptionalInt(arguments, "end_paragraph", 1),
                Level = ToolArgumentValidator.OptionalInt(arguments, "level", 1, 9),
            };
            Logger.Debug($"执行 set_list_format，type={request.Type}");
            return JObject.FromObject(Service.SetListFormat(request));
        }
    }

    /// <summary>设置段落底纹与边框。</summary>
    public sealed class SetParagraphShadingTool : ToolBase
    {
        public SetParagraphShadingTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_paragraph_shading";

        public override string Description => "设置段落底纹与边框（提示框/引用块效果）：用 paragraph_index 或 find_text 定位段落，可设置背景色、四周边框颜色/样式/粗细。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["paragraph_index"] = Int("段落索引（1 基，可选）", 1),
            ["find_text"] = Str("按文本定位段落（包含该文本的第一个段落，可选）"),
            ["shading"] = Str("底纹颜色：#RRGGBB / #AARRGGBB / 颜色名"),
            ["border_color"] = Str("边框颜色：#RRGGBB / #AARRGGBB / 颜色名"),
            ["border_style"] = Str("边框样式", "single", "double", "dotted", "dashed", "none"),
            ["border_width"] = Str("边框宽度：thin / medium / thick 或磅值（如 1.5）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetParagraphShadingRequest
            {
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
                FindText = ToolArgumentValidator.OptionalString(arguments, "find_text"),
                ShadingArgb = ToolArgumentValidator.OptionalColor(arguments, "shading"),
                BorderArgb = ToolArgumentValidator.OptionalColor(arguments, "border_color"),
                BorderStyle = ToolArgumentValidator.OptionalString(arguments, "border_style"),
                BorderWidth = ToolArgumentValidator.OptionalString(arguments, "border_width"),
            };
            Logger.Debug($"执行 set_paragraph_shading，paragraphIndex={request.ParagraphIndex}");
            return JObject.FromObject(Service.SetParagraphShading(request));
        }
    }
}
