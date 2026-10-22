using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>插入尾注。</summary>
    public sealed class InsertEndnoteTool : ToolBase
    {
        public InsertEndnoteTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_endnote";

        public override string Description => "插入尾注：默认在指定段落末尾插入引用标记；提供 anchor_text 时在该文本之后插入。尾注内容显示在文档/节末尾。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("尾注内容"),
            ["paragraph_index"] = Int("目标段落（1 基，可选）；缺省最后一个段落", 1),
            ["anchor_text"] = Str("角标锚点文本：在该文本之后插入引用标记（可选，优先于 paragraph_index）"),
        }, "text");

        public override JObject Execute(JObject arguments)
        {
            var text = ToolArgumentValidator.RequireString(arguments, "text");
            var paragraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1);
            var anchorText = ToolArgumentValidator.OptionalString(arguments, "anchor_text");
            Logger.Debug($"执行 insert_endnote，paragraphIndex={paragraphIndex}，anchorText={anchorText}");
            return JObject.FromObject(Service.InsertEndnote(text, paragraphIndex, anchorText));
        }
    }

    /// <summary>列出脚注/尾注。</summary>
    public sealed class ListNotesTool : ToolBase
    {
        public ListNotesTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "list_notes";

        public override string Description => "列出脚注或尾注（索引、内容、引用所在段落）。note_type 支持 footnote / endnote，默认 footnote。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["note_type"] = Str("注释类型", "footnote", "endnote"),
        });

        public override JObject Execute(JObject arguments)
        {
            var noteType = ToolArgumentValidator.OptionalString(arguments, "note_type") ?? "footnote";
            Logger.Debug($"执行 list_notes，noteType={noteType}");
            var notes = Service.ListNotes(noteType);
            return new JObject
            {
                ["note_type"] = noteType,
                ["count"] = notes.Count,
                ["notes"] = JArray.FromObject(notes),
            };
        }
    }

    /// <summary>读取单条脚注/尾注。</summary>
    public sealed class GetNoteTool : ToolBase
    {
        public GetNoteTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_note";

        public override string Description => "读取指定脚注/尾注的详细信息（索引、内容、引用所在段落与段落预览）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["index"] = Int("注释索引（1 基）", 1),
            ["note_type"] = Str("注释类型", "footnote", "endnote"),
        }, "index");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "index", 1);
            var noteType = ToolArgumentValidator.OptionalString(arguments, "note_type") ?? "footnote";
            Logger.Debug($"执行 get_note，noteType={noteType}，index={index}");
            return JObject.FromObject(Service.GetNote(noteType, index));
        }
    }

    /// <summary>删除脚注/尾注。</summary>
    public sealed class DeleteNoteTool : ToolBase
    {
        public DeleteNoteTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_note";

        public override string Description => "按索引删除脚注或尾注（同时删除正文中的引用标记与注释内容）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["index"] = Int("注释索引（1 基）", 1),
            ["note_type"] = Str("注释类型", "footnote", "endnote"),
        }, "index");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "index", 1);
            var noteType = ToolArgumentValidator.OptionalString(arguments, "note_type") ?? "footnote";
            Logger.Debug($"执行 delete_note，noteType={noteType}，index={index}");
            Service.DeleteNote(noteType, index);
            return new JObject { ["deleted"] = true, ["note_type"] = noteType, ["index"] = index };
        }
    }

    /// <summary>设置脚注/尾注编号选项。</summary>
    public sealed class SetNoteOptionsTool : ToolBase
    {
        public SetNoteOptionsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_note_options";

        public override string Description => "设置脚注/尾注编号：编号样式（阿拉伯数字/罗马数字/字母）、重编规则（连续/每节/每页）、起始编号；脚注还可设置位置（页面底部/文字下方）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["note_type"] = Str("注释类型", "footnote", "endnote"),
            ["number_style"] = Str("编号样式", "arabic", "upper-roman", "lower-roman", "upper-letter", "lower-letter"),
            ["numbering_rule"] = Str("重编规则", "continuous", "restart-section", "restart-page"),
            ["starting_number"] = Int("起始编号（>=1）", 1),
            ["location"] = Str("脚注位置（仅 footnote）", "bottom", "beneath-text"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetNoteOptionsRequest
            {
                NoteType = ToolArgumentValidator.OptionalString(arguments, "note_type") ?? "footnote",
                NumberStyle = ToolArgumentValidator.OptionalString(arguments, "number_style"),
                NumberingRule = ToolArgumentValidator.OptionalString(arguments, "numbering_rule"),
                StartingNumber = ToolArgumentValidator.OptionalInt(arguments, "starting_number", 1),
                Location = ToolArgumentValidator.OptionalString(arguments, "location"),
            };
            Logger.Debug($"执行 set_note_options，noteType={request.NoteType}");
            return JObject.FromObject(Service.SetNoteOptions(request));
        }
    }
}
