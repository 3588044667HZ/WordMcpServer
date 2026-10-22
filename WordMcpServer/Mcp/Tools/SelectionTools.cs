using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>读取当前选区。</summary>
    public sealed class GetSelectionTool : ToolBase
    {
        public GetSelectionTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_selection";

        public override string Description => "读取当前选区信息：选区类型、起止位置、选中文本、所在段落、页码、是否为空（光标）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 get_selection");
            return JObject.FromObject(Service.GetSelection());
        }
    }

    /// <summary>用文本替换当前选区。</summary>
    public sealed class ReplaceSelectionTool : ToolBase
    {
        public ReplaceSelectionTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "replace_selection";

        public override string Description => "用新文本替换当前选区（光标处则直接插入），可选套用样式。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("替换为的文本"),
            ["style"] = Str("替换后段落样式（可选）"),
        }, "text");

        public override JObject Execute(JObject arguments)
        {
            var text = ToolArgumentValidator.RequireString(arguments, "text");
            var style = ToolArgumentValidator.OptionalString(arguments, "style");
            Logger.Debug("执行 replace_selection");
            return JObject.FromObject(Service.ReplaceSelection(text, style));
        }
    }

    /// <summary>跳转到指定页。</summary>
    public sealed class GotoPageTool : ToolBase
    {
        public GotoPageTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "goto_page";

        public override string Description => "把光标/视图跳转到指定页，返回跳转后的选区信息（页码、段落、文本）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["page"] = Int("页码（1 基）", 1),
        }, "page");

        public override JObject Execute(JObject arguments)
        {
            var page = ToolArgumentValidator.RequireInt(arguments, "page", 1);
            Logger.Debug($"执行 goto_page，page={page}");
            return JObject.FromObject(Service.GotoPage(page));
        }
    }

    /// <summary>移动段落。</summary>
    public sealed class MoveParagraphTool : ToolBase
    {
        public MoveParagraphTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "move_paragraph";

        public override string Description => "把指定段落上移或下移若干位置（保留格式），返回移动后的段落信息。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["paragraph_index"] = Int("段落索引（1 基）", 1),
            ["offset"] = Int("移动偏移：负数上移、正数下移（如 -1 上移一位）"),
        }, "paragraph_index", "offset");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "paragraph_index", 1);
            var offset = ToolArgumentValidator.RequireInt(arguments, "offset");
            Logger.Debug($"执行 move_paragraph，index={index}，offset={offset}");
            return JObject.FromObject(Service.MoveParagraph(index, offset));
        }
    }

    /// <summary>复制内容到指定位置。</summary>
    public sealed class CopyRangeTool : ToolBase
    {
        public CopyRangeTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "copy_range";

        public override string Description => "复制段落区间或匹配文本，插入到目标段落之前（缺省文末），保留原格式。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["source_start_paragraph"] = Int("源起始段落（1 基，可选）", 1),
            ["source_end_paragraph"] = Int("源结束段落（1 基，含，可选）", 1),
            ["find_text"] = Str("按文本定位要复制的内容（可选，优先于段落区间）"),
            ["target_paragraph"] = Int("插入到该段落之前（1 基）；缺省追加到文档末尾", 1),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new CopyRangeRequest
            {
                SourceStartParagraph = ToolArgumentValidator.OptionalInt(arguments, "source_start_paragraph", 1),
                SourceEndParagraph = ToolArgumentValidator.OptionalInt(arguments, "source_end_paragraph", 1),
                FindText = ToolArgumentValidator.OptionalString(arguments, "find_text"),
                TargetParagraph = ToolArgumentValidator.OptionalInt(arguments, "target_paragraph", 1),
            };
            Logger.Debug("执行 copy_range");
            return JObject.FromObject(Service.CopyRange(request));
        }
    }
}
