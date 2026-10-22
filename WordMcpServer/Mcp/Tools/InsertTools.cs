using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>插入图片。</summary>
    public sealed class InsertImageTool : ToolBase
    {
        public InsertImageTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_image";

        public override string Description => "插入内嵌图片：默认追加到文档末尾（新段落），或用 paragraph_index 插入到指定段落开头。width/height 单位为磅；只给一个时等比缩放。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["path"] = Str("图片文件完整路径（png / jpg / bmp / gif 等）"),
            ["width"] = Num("宽度（磅），可选", 1),
            ["height"] = Num("高度（磅），可选", 1),
            ["alignment"] = Str("图片所在段落对齐", "left", "center", "right"),
            ["paragraph_index"] = Int("插入到该段落开头（1 基，可选）；缺省追加到文档末尾", 1),
        }, "path");

        public override JObject Execute(JObject arguments)
        {
            var request = new InsertImageRequest
            {
                Path = ToolArgumentValidator.RequireString(arguments, "path"),
                Width = ToolArgumentValidator.OptionalDouble(arguments, "width", 1),
                Height = ToolArgumentValidator.OptionalDouble(arguments, "height", 1),
                Alignment = ToolArgumentValidator.OptionalString(arguments, "alignment"),
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
            };
            Logger.Debug($"执行 insert_image，path={request.Path}");
            return JObject.FromObject(Service.InsertImage(request));
        }
    }

    /// <summary>插入目录。</summary>
    public sealed class InsertTocTool : ToolBase
    {
        public InsertTocTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_toc";

        public override string Description => "在文档末尾插入自动目录（基于 Heading 样式/大纲级别）。levels 控制包含的标题级别（默认 3）；title 可选，作为目录上方的标题文本。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["levels"] = Int("包含的标题级别数（1-9，默认 3）", 1, 9),
            ["title"] = Str("目录标题文字（可选）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new InsertTocRequest
            {
                Levels = ToolArgumentValidator.OptionalInt(arguments, "levels", 1, 9) ?? 3,
                Title = ToolArgumentValidator.OptionalString(arguments, "title"),
            };
            Logger.Debug($"执行 insert_toc，levels={request.Levels}");
            return JObject.FromObject(Service.InsertToc(request));
        }
    }

    /// <summary>更新目录。</summary>
    public sealed class UpdateTocTool : ToolBase
    {
        public UpdateTocTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "update_toc";

        public override string Description => "更新文档中所有自动目录的页码与条目。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 update_toc");
            return JObject.FromObject(Service.UpdateToc());
        }
    }

    /// <summary>设置页眉页脚。</summary>
    public sealed class SetHeaderFooterTool : ToolBase
    {
        public SetHeaderFooterTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_header_footer";

        public override string Description => "设置指定节的页眉/页脚文本（未提供的保持不变）；include_page_number=true 时在页脚追加页码域。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["section_index"] = Int("节索引（1 基，默认 1）", 1),
            ["header_text"] = Str("页眉文本（可选）"),
            ["footer_text"] = Str("页脚文本（可选）"),
            ["include_page_number"] = Bool("是否在页脚追加页码，默认 false"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetHeaderFooterRequest
            {
                SectionIndex = ToolArgumentValidator.OptionalInt(arguments, "section_index", 1) ?? 1,
                HeaderText = ToolArgumentValidator.OptionalString(arguments, "header_text"),
                FooterText = ToolArgumentValidator.OptionalString(arguments, "footer_text"),
                IncludePageNumber = ToolArgumentValidator.OptionalBool(arguments, "include_page_number") ?? false,
            };
            Logger.Debug($"执行 set_header_footer，sectionIndex={request.SectionIndex}");
            return JObject.FromObject(Service.SetHeaderFooter(request));
        }
    }

    /// <summary>读取页眉页脚。</summary>
    public sealed class GetHeaderFooterTool : ToolBase
    {
        public GetHeaderFooterTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_header_footer";

        public override string Description => "读取指定节的页眉与页脚文本。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["section_index"] = Int("节索引（1 基，默认 1）", 1),
        });

        public override JObject Execute(JObject arguments)
        {
            var sectionIndex = ToolArgumentValidator.OptionalInt(arguments, "section_index", 1) ?? 1;
            Logger.Debug($"执行 get_header_footer，sectionIndex={sectionIndex}");
            return JObject.FromObject(Service.GetHeaderFooter(sectionIndex));
        }
    }

    /// <summary>插入批注。</summary>
    public sealed class InsertCommentTool : ToolBase
    {
        public InsertCommentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_comment";

        public override string Description => "在指定位置插入批注：用 paragraph_index 对整个段落加批注，或用 anchor_text 查找文本后加批注。text 为批注内容。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("批注内容"),
            ["anchor_text"] = Str("批注锚点文本（查找该文本并加批注）"),
            ["paragraph_index"] = Int("对指定段落加批注（1 基，可选）", 1),
        }, "text");

        public override JObject Execute(JObject arguments)
        {
            var text = ToolArgumentValidator.RequireString(arguments, "text");
            var anchorText = ToolArgumentValidator.OptionalString(arguments, "anchor_text");
            var paragraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1);
            Logger.Debug($"执行 insert_comment，anchorText={anchorText}，paragraphIndex={paragraphIndex}");
            return JObject.FromObject(Service.InsertComment(text, anchorText, paragraphIndex));
        }
    }

    /// <summary>列出批注。</summary>
    public sealed class ListCommentsTool : ToolBase
    {
        public ListCommentsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "list_comments";

        public override string Description => "列出文档中的所有批注（索引、作者、内容、锚定文本）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 list_comments");
            var comments = Service.ListComments();
            return new JObject
            {
                ["count"] = comments.Count,
                ["comments"] = JArray.FromObject(comments),
            };
        }
    }

    /// <summary>删除批注。</summary>
    public sealed class DeleteCommentTool : ToolBase
    {
        public DeleteCommentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_comment";

        public override string Description => "按索引删除批注（索引来自 list_comments）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["index"] = Int("批注索引（1 基）", 1),
        }, "index");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "index", 1);
            Logger.Debug($"执行 delete_comment，index={index}");
            Service.DeleteComment(index);
            return new JObject { ["deleted"] = true, ["index"] = index };
        }
    }

    /// <summary>添加文字水印。</summary>
    public sealed class AddWatermarkTool : ToolBase
    {
        public AddWatermarkTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "add_watermark";

        public override string Description => "添加文字水印（每个节的页眉中插入 WordArt，幂等：重复调用会先移除旧水印）。支持斜向/水平布局与透明度。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("水印文字（如“内部资料”“DRAFT”）"),
            ["font_name"] = Str("字体名称（默认微软雅黑）"),
            ["font_size"] = Num("字号（磅，默认按文字长度自动计算 24-72）", 1, 300),
            ["color"] = Str("颜色：#RRGGBB / #AARRGGBB / 颜色名（默认浅灰 #C0C0C0）"),
            ["layout"] = Str("布局", "diagonal", "horizontal"),
            ["transparency"] = Num("透明度 0-1（默认 0.5）", 0, 1),
        }, "text");

        public override JObject Execute(JObject arguments)
        {
            var request = new AddWatermarkRequest
            {
                Text = ToolArgumentValidator.RequireString(arguments, "text"),
                FontName = ToolArgumentValidator.OptionalString(arguments, "font_name"),
                FontSize = ToolArgumentValidator.OptionalDouble(arguments, "font_size", 1, 300),
                ColorArgb = ToolArgumentValidator.OptionalColor(arguments, "color"),
                Layout = ToolArgumentValidator.OptionalString(arguments, "layout"),
                Transparency = ToolArgumentValidator.OptionalDouble(arguments, "transparency", 0, 1),
            };
            Logger.Debug($"执行 add_watermark，text={request.Text}");
            return JObject.FromObject(Service.AddWatermark(request));
        }
    }

    /// <summary>删除水印。</summary>
    public sealed class RemoveWatermarkTool : ToolBase
    {
        public RemoveWatermarkTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "remove_watermark";

        public override string Description => "删除由本插件添加的文字水印（所有节的页眉）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 remove_watermark");
            return JObject.FromObject(Service.RemoveWatermark());
        }
    }

    /// <summary>插入脚注。</summary>
    public sealed class InsertFootnoteTool : ToolBase
    {
        public InsertFootnoteTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_footnote";

        public override string Description => "插入脚注（真角标）：默认在指定段落末尾插入引用标记；提供 anchor_text 时在该文本之后插入。脚注内容显示在页面底部。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("脚注内容"),
            ["paragraph_index"] = Int("目标段落（1 基，可选）；缺省最后一个段落", 1),
            ["anchor_text"] = Str("角标锚点文本：在该文本之后插入引用标记（可选，优先于 paragraph_index）"),
        }, "text");

        public override JObject Execute(JObject arguments)
        {
            var text = ToolArgumentValidator.RequireString(arguments, "text");
            var paragraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1);
            var anchorText = ToolArgumentValidator.OptionalString(arguments, "anchor_text");
            Logger.Debug($"执行 insert_footnote，paragraphIndex={paragraphIndex}，anchorText={anchorText}");
            return JObject.FromObject(Service.InsertFootnote(text, paragraphIndex, anchorText));
        }
    }
}
