using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>插入交叉引用。</summary>
    public sealed class InsertCrossReferenceTool : ToolBase
    {
        public InsertCrossReferenceTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_cross_reference";

        public override string Description => "插入交叉引用：支持引用标题（heading）、书签（bookmark）、脚注/尾注编号（footnote/endnote）、题注（caption，需 caption_label）。可插入为可点击超链接。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["reference_type"] = Str("引用类型", "heading", "bookmark", "footnote", "endnote", "caption"),
            ["reference_item"] = Str("引用目标：标题文本、书签名、脚注/尾注编号、题注编号"),
            ["caption_label"] = Str("题注标签（仅 caption 类型，如 图 / 表）"),
            ["reference_kind"] = Str("引用方式", "text", "number", "page", "label-and-number", "entire-caption"),
            ["insert_as_hyperlink"] = Bool("是否插入为可点击超链接，默认 true"),
            ["paragraph_index"] = Int("插入到该段落末尾（1 基，可选）；缺省追加到文档末尾", 1),
        }, "reference_type", "reference_item");

        public override JObject Execute(JObject arguments)
        {
            var request = new CrossReferenceRequest
            {
                ReferenceType = ToolArgumentValidator.RequireString(arguments, "reference_type"),
                ReferenceItem = ToolArgumentValidator.RequireString(arguments, "reference_item"),
                CaptionLabel = ToolArgumentValidator.OptionalString(arguments, "caption_label"),
                ReferenceKind = ToolArgumentValidator.OptionalString(arguments, "reference_kind"),
                InsertAsHyperlink = ToolArgumentValidator.OptionalBool(arguments, "insert_as_hyperlink") ?? true,
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
            };
            Logger.Debug($"执行 insert_cross_reference，type={request.ReferenceType}，item={request.ReferenceItem}");
            return JObject.FromObject(Service.InsertCrossReference(request));
        }
    }

    /// <summary>插入题注。</summary>
    public sealed class InsertCaptionTool : ToolBase
    {
        public InsertCaptionTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_caption";

        public override string Description => "为图/表插入自动编号题注（SEQ 域）。label 为标签（图/表/Figure），title 为题注文字，position 控制题注在目标段落上方或下方。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["label"] = Str("题注标签（默认“图”，表格可传“表”）"),
            ["title"] = Str("题注文字（可选，如“模型架构示意图”）"),
            ["position"] = Str("题注位置", "above", "below"),
            ["paragraph_index"] = Int("题注插入到该段落之后（below）或之前（above）；缺省最后一个段落", 1),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new InsertCaptionRequest
            {
                Label = ToolArgumentValidator.OptionalString(arguments, "label") ?? "图",
                Title = ToolArgumentValidator.OptionalString(arguments, "title"),
                Position = ToolArgumentValidator.OptionalString(arguments, "position") ?? "below",
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
            };
            Logger.Debug($"执行 insert_caption，label={request.Label}，position={request.Position}");
            return JObject.FromObject(Service.InsertCaption(request));
        }
    }

    /// <summary>插入通用域。</summary>
    public sealed class InsertFieldTool : ToolBase
    {
        public InsertFieldTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_field";

        public override string Description => "插入 Word 域（通用）：如 PAGEREF 书签 \\h（页码引用）、SEQ 图 \\* ARABIC（自动编号）、DATE、FILENAME 等。插入后可立即更新域。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["field_code"] = Str("域代码，如 \"PAGEREF mybookmark \\\\h\" 或 \"SEQ 图 \\\\* ARABIC\""),
            ["paragraph_index"] = Int("插入到该段落末尾（1 基，可选）；缺省追加到文档末尾", 1),
            ["update"] = Bool("插入后是否立即更新域，默认 true"),
        }, "field_code");

        public override JObject Execute(JObject arguments)
        {
            var request = new InsertFieldRequest
            {
                FieldCode = ToolArgumentValidator.RequireString(arguments, "field_code"),
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
                Update = ToolArgumentValidator.OptionalBool(arguments, "update") ?? true,
            };
            Logger.Debug($"执行 insert_field，code={request.FieldCode}");
            return JObject.FromObject(Service.InsertField(request));
        }
    }
}
