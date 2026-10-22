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

        public override string Description => "插入图片：内嵌（inline，默认）或浮动（floating，可设位置、环绕、定位基准）。width/height 单位为磅；只给一个时等比缩放。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["path"] = Str("图片文件完整路径（png / jpg / bmp / gif 等）"),
            ["width"] = Num("宽度（磅），可选", 1),
            ["height"] = Num("高度（磅），可选", 1),
            ["alignment"] = Str("图片所在段落对齐", "left", "center", "right"),
            ["layout"] = Str("布局", "inline", "floating"),
            ["wrap"] = Str("浮动环绕方式（floating 时可用）", "square", "tight", "through", "none", "top-bottom", "behind", "front"),
            ["left"] = Num("浮动图片左边距（磅）"),
            ["top"] = Num("浮动图片上边距（磅）"),
            ["relative_to"] = Str("浮动定位基准", "page", "margin", "column", "paragraph"),
            ["alt_text"] = Str("替代文字（可选）"),
            ["paragraph_index"] = Int("插入到该段落（1 基，可选）；缺省追加到文档末尾", 1),
        }, "path");

        public override JObject Execute(JObject arguments)
        {
            var request = new InsertImageRequest
            {
                Path = ToolArgumentValidator.RequireString(arguments, "path"),
                Width = ToolArgumentValidator.OptionalDouble(arguments, "width", 1),
                Height = ToolArgumentValidator.OptionalDouble(arguments, "height", 1),
                Alignment = ToolArgumentValidator.OptionalString(arguments, "alignment"),
                Layout = ToolArgumentValidator.OptionalString(arguments, "layout") ?? "inline",
                Wrap = ToolArgumentValidator.OptionalString(arguments, "wrap"),
                Left = ToolArgumentValidator.OptionalDouble(arguments, "left"),
                Top = ToolArgumentValidator.OptionalDouble(arguments, "top"),
                RelativeTo = ToolArgumentValidator.OptionalString(arguments, "relative_to") ?? "page",
                AltText = ToolArgumentValidator.OptionalString(arguments, "alt_text"),
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
            };
            Logger.Debug($"执行 insert_image，path={request.Path}，layout={request.Layout}");
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

        public override string Description => "设置指定节的页眉/页脚（primary / first-page / even-pages），可控制首页不同、奇偶页不同、链接前一节，并可在页脚追加页码域。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["section_index"] = Int("节索引（1 基，默认 1）", 1),
            ["header_type"] = Str("页眉页脚类型", "primary", "first-page", "even-pages"),
            ["header_text"] = Str("页眉文本（可选）"),
            ["footer_text"] = Str("页脚文本（可选）"),
            ["include_page_number"] = Bool("是否在页脚追加页码，默认 false"),
            ["link_to_previous"] = Bool("是否链接到前一节（可选）"),
            ["different_first_page"] = Bool("首页页眉页脚不同（可选）"),
            ["odd_even_pages"] = Bool("奇偶页页眉页脚不同（可选）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetHeaderFooterRequest
            {
                SectionIndex = ToolArgumentValidator.OptionalInt(arguments, "section_index", 1) ?? 1,
                HeaderType = ToolArgumentValidator.OptionalString(arguments, "header_type") ?? "primary",
                HeaderText = ToolArgumentValidator.OptionalString(arguments, "header_text"),
                FooterText = ToolArgumentValidator.OptionalString(arguments, "footer_text"),
                IncludePageNumber = ToolArgumentValidator.OptionalBool(arguments, "include_page_number") ?? false,
                LinkToPrevious = ToolArgumentValidator.OptionalBool(arguments, "link_to_previous"),
                DifferentFirstPage = ToolArgumentValidator.OptionalBool(arguments, "different_first_page"),
                OddEvenPages = ToolArgumentValidator.OptionalBool(arguments, "odd_even_pages"),
            };
            Logger.Debug($"执行 set_header_footer，sectionIndex={request.SectionIndex}，headerType={request.HeaderType}");
            return JObject.FromObject(Service.SetHeaderFooter(request));
        }
    }

    /// <summary>读取页眉页脚。</summary>
    public sealed class GetHeaderFooterTool : ToolBase
    {
        public GetHeaderFooterTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_header_footer";

        public override string Description => "读取指定节、指定类型（primary / first-page / even-pages）的页眉与页脚文本，以及链接/首页不同/奇偶页设置。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["section_index"] = Int("节索引（1 基，默认 1）", 1),
            ["header_type"] = Str("页眉页脚类型", "primary", "first-page", "even-pages"),
        });

        public override JObject Execute(JObject arguments)
        {
            var sectionIndex = ToolArgumentValidator.OptionalInt(arguments, "section_index", 1) ?? 1;
            var headerType = ToolArgumentValidator.OptionalString(arguments, "header_type") ?? "primary";
            Logger.Debug($"执行 get_header_footer，sectionIndex={sectionIndex}，headerType={headerType}");
            return JObject.FromObject(Service.GetHeaderFooter(sectionIndex, headerType));
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

    /// <summary>列出图片。</summary>
    public sealed class ListImagesTool : ToolBase
    {
        public ListImagesTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "list_images";

        public override string Description => "列出文档中的图片：内嵌图片（索引/宽高/替代文字/段落）与浮动图片（名称/位置/环绕/替代文字）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 list_images");
            var images = Service.ListImages();
            return new JObject
            {
                ["count"] = images.Count,
                ["images"] = JArray.FromObject(images),
            };
        }
    }

    /// <summary>修改图片。</summary>
    public sealed class UpdateImageTool : ToolBase
    {
        public UpdateImageTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "update_image";

        public override string Description => "修改图片：内嵌图片（按 index）可改宽高/缩放百分比/裁剪（百分比）/替代文字；浮动图片（按 name 或 index）可改宽高/替代文字。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["kind"] = Str("图片类型", "inline", "floating"),
            ["index"] = Int("内嵌图片索引，或浮动图片索引（1 基）", 1),
            ["name"] = Str("浮动图片名称（可选，优先于 index）"),
            ["width"] = Num("宽度（磅）", 1),
            ["height"] = Num("高度（磅）", 1),
            ["scale_width_percent"] = Num("宽度缩放百分比（如 50）", 1),
            ["scale_height_percent"] = Num("高度缩放百分比（如 50）", 1),
            ["crop_top"] = Num("顶部裁剪百分比（0-100）", 0, 100),
            ["crop_bottom"] = Num("底部裁剪百分比（0-100）", 0, 100),
            ["crop_left"] = Num("左侧裁剪百分比（0-100）", 0, 100),
            ["crop_right"] = Num("右侧裁剪百分比（0-100）", 0, 100),
            ["alt_text"] = Str("替代文字"),
            ["title"] = Str("标题（可选）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new UpdateImageRequest
            {
                Kind = ToolArgumentValidator.OptionalString(arguments, "kind") ?? "inline",
                Index = ToolArgumentValidator.OptionalInt(arguments, "index", 1),
                Name = ToolArgumentValidator.OptionalString(arguments, "name"),
                Width = ToolArgumentValidator.OptionalDouble(arguments, "width", 1),
                Height = ToolArgumentValidator.OptionalDouble(arguments, "height", 1),
                ScaleWidthPercent = ToolArgumentValidator.OptionalDouble(arguments, "scale_width_percent", 1),
                ScaleHeightPercent = ToolArgumentValidator.OptionalDouble(arguments, "scale_height_percent", 1),
                CropTop = ToolArgumentValidator.OptionalDouble(arguments, "crop_top", 0, 100),
                CropBottom = ToolArgumentValidator.OptionalDouble(arguments, "crop_bottom", 0, 100),
                CropLeft = ToolArgumentValidator.OptionalDouble(arguments, "crop_left", 0, 100),
                CropRight = ToolArgumentValidator.OptionalDouble(arguments, "crop_right", 0, 100),
                AltText = ToolArgumentValidator.OptionalString(arguments, "alt_text"),
                Title = ToolArgumentValidator.OptionalString(arguments, "title"),
            };
            Logger.Debug($"执行 update_image，kind={request.Kind}");
            return JObject.FromObject(Service.UpdateImage(request));
        }
    }

    /// <summary>删除图片。</summary>
    public sealed class DeleteImageTool : ToolBase
    {
        public DeleteImageTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_image";

        public override string Description => "删除图片：内嵌图片按 index，浮动图片按 name 或 index。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["kind"] = Str("图片类型", "inline", "floating"),
            ["index"] = Int("图片索引（1 基，可选）", 1),
            ["name"] = Str("浮动图片名称（可选）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var kind = ToolArgumentValidator.OptionalString(arguments, "kind") ?? "inline";
            var index = ToolArgumentValidator.OptionalInt(arguments, "index", 1);
            var name = ToolArgumentValidator.OptionalString(arguments, "name");
            Logger.Debug($"执行 delete_image，kind={kind}，index={index}，name={name}");
            Service.DeleteImage(kind, index, name);
            return new JObject { ["deleted"] = true, ["kind"] = kind };
        }
    }

    /// <summary>设置页码格式。</summary>
    public sealed class SetPageNumbersTool : ToolBase
    {
        public SetPageNumbersTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_page_numbers";

        public override string Description => "设置指定节的页码格式：编号样式（阿拉伯/罗马/字母）、起始编号、是否从本节重新编号。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["section_index"] = Int("节索引（1 基，默认 1）", 1),
            ["number_style"] = Str("页码样式", "arabic", "upper-roman", "lower-roman", "upper-letter", "lower-letter"),
            ["starting_number"] = Int("起始编号（>=1）", 1),
            ["restart_at_section"] = Bool("是否从本节重新编号"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetPageNumbersRequest
            {
                SectionIndex = ToolArgumentValidator.OptionalInt(arguments, "section_index", 1) ?? 1,
                NumberStyle = ToolArgumentValidator.OptionalString(arguments, "number_style"),
                StartingNumber = ToolArgumentValidator.OptionalInt(arguments, "starting_number", 1),
                RestartAtSection = ToolArgumentValidator.OptionalBool(arguments, "restart_at_section"),
            };
            Logger.Debug($"执行 set_page_numbers，sectionIndex={request.SectionIndex}");
            return JObject.FromObject(Service.SetPageNumbers(request));
        }
    }

    /// <summary>读取批注。</summary>
    public sealed class GetCommentTool : ToolBase
    {
        public GetCommentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_comment";

        public override string Description => "读取指定批注的详细信息（作者、内容、锚定文本、是否已解决、回复数）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["index"] = Int("批注索引（1 基）", 1),
        }, "index");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "index", 1);
            Logger.Debug($"执行 get_comment，index={index}");
            return JObject.FromObject(Service.GetComment(index));
        }
    }

    /// <summary>回复批注。</summary>
    public sealed class ReplyCommentTool : ToolBase
    {
        public ReplyCommentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "reply_comment";

        public override string Description => "回复指定批注（Word 批注回复）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["index"] = Int("父批注索引（1 基）", 1),
            ["text"] = Str("回复内容"),
        }, "index", "text");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "index", 1);
            var text = ToolArgumentValidator.RequireString(arguments, "text");
            Logger.Debug($"执行 reply_comment，index={index}");
            return JObject.FromObject(Service.ReplyComment(index, text));
        }
    }

    /// <summary>标记批注已解决/未解决。</summary>
    public sealed class ResolveCommentTool : ToolBase
    {
        public ResolveCommentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "resolve_comment";

        public override string Description => "把批注标记为已解决（done=true）或重新打开（done=false）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["index"] = Int("批注索引（1 基）", 1),
            ["done"] = Bool("true 标记已解决（默认），false 取消"),
        }, "index");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "index", 1);
            var done = ToolArgumentValidator.OptionalBool(arguments, "done") ?? true;
            Logger.Debug($"执行 resolve_comment，index={index}，done={done}");
            return JObject.FromObject(Service.ResolveComment(index, done));
        }
    }

    /// <summary>插入封面页。</summary>
    public sealed class InsertCoverPageTool : ToolBase
    {
        public InsertCoverPageTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_cover_page";

        public override string Description => "在文档开头生成封面页：标题、副标题、作者、单位、日期（居中排版），可选在封面后插入分页符。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["title"] = Str("封面标题"),
            ["subtitle"] = Str("副标题（可选）"),
            ["author"] = Str("作者（可选）"),
            ["institution"] = Str("单位/机构（可选）"),
            ["date"] = Str("日期文字（可选，默认当天日期）"),
            ["page_break"] = Bool("封面后是否分页，默认 true"),
        }, "title");

        public override JObject Execute(JObject arguments)
        {
            var request = new CoverPageRequest
            {
                Title = ToolArgumentValidator.RequireString(arguments, "title"),
                Subtitle = ToolArgumentValidator.OptionalString(arguments, "subtitle"),
                Author = ToolArgumentValidator.OptionalString(arguments, "author"),
                Institution = ToolArgumentValidator.OptionalString(arguments, "institution"),
                Date = ToolArgumentValidator.OptionalString(arguments, "date"),
                PageBreak = ToolArgumentValidator.OptionalBool(arguments, "page_break") ?? true,
            };
            Logger.Debug($"执行 insert_cover_page，title={request.Title}");
            return JObject.FromObject(Service.InsertCoverPage(request));
        }
    }
}
