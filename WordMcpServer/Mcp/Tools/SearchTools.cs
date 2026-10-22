using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>查找替换文本。</summary>
    public sealed class FindReplaceTextTool : ToolBase
    {
        public FindReplaceTextTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "find_replace_text";

        public override string Description => "在全文范围查找并替换文本（replace_text 传空字符串表示删除匹配文本）。默认替换全部（replace_all=true）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["find_text"] = Str("要查找的文本"),
            ["replace_text"] = Str("替换为的文本（可为空字符串表示删除）"),
            ["match_case"] = Bool("区分大小写，默认 false"),
            ["match_whole_word"] = Bool("全字匹配，默认 false"),
            ["replace_all"] = Bool("是否替换全部，默认 true（false 只替换第一处）"),
        }, "find_text");

        public override JObject Execute(JObject arguments)
        {
            var request = new FindReplaceRequest
            {
                FindText = ToolArgumentValidator.RequireString(arguments, "find_text"),
                ReplaceText = ToolArgumentValidator.OptionalString(arguments, "replace_text") ?? string.Empty,
                MatchCase = ToolArgumentValidator.OptionalBool(arguments, "match_case") ?? false,
                MatchWholeWord = ToolArgumentValidator.OptionalBool(arguments, "match_whole_word") ?? false,
                ReplaceAll = ToolArgumentValidator.OptionalBool(arguments, "replace_all") ?? true,
            };
            Logger.Debug($"执行 find_replace_text，findText={request.FindText}");
            return JObject.FromObject(Service.FindReplaceText(request));
        }
    }

    /// <summary>查找文本位置。</summary>
    public sealed class FindTextTool : ToolBase
    {
        public FindTextTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "find_text";

        public override string Description => "查找文本在文档中的所有出现位置（段落索引与字符位置），最多返回 max_results 条。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("要查找的文本"),
            ["match_case"] = Bool("区分大小写，默认 false"),
            ["match_whole_word"] = Bool("全字匹配，默认 false"),
            ["max_results"] = Int("最多返回的匹配数（默认 20）", 1, 200),
        }, "text");

        public override JObject Execute(JObject arguments)
        {
            var request = new FindTextRequest
            {
                Text = ToolArgumentValidator.RequireString(arguments, "text"),
                MatchCase = ToolArgumentValidator.OptionalBool(arguments, "match_case") ?? false,
                MatchWholeWord = ToolArgumentValidator.OptionalBool(arguments, "match_whole_word") ?? false,
                MaxResults = ToolArgumentValidator.OptionalInt(arguments, "max_results", 1, 200) ?? 20,
            };
            Logger.Debug($"执行 find_text，text={request.Text}");
            return JObject.FromObject(Service.FindText(request));
        }
    }

    /// <summary>插入书签。</summary>
    public sealed class InsertBookmarkTool : ToolBase
    {
        public InsertBookmarkTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_bookmark";

        public override string Description => "在文档末尾追加一段文本并将其标记为书签（名称不能与现有书签重复）。后续可用 get_bookmark_text 读取。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["name"] = Str("书签名称（字母开头，不含空格）"),
            ["text"] = Str("书签内容文本（可选）"),
        }, "name");

        public override JObject Execute(JObject arguments)
        {
            var name = ToolArgumentValidator.RequireString(arguments, "name");
            var text = ToolArgumentValidator.OptionalString(arguments, "text");
            Logger.Debug($"执行 insert_bookmark，name={name}");
            return JObject.FromObject(Service.InsertBookmark(name, text));
        }
    }

    /// <summary>读取书签内容。</summary>
    public sealed class GetBookmarkTextTool : ToolBase
    {
        public GetBookmarkTextTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_bookmark_text";

        public override string Description => "读取指定书签的范围文本与字符位置。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["name"] = Str("书签名称"),
        }, "name");

        public override JObject Execute(JObject arguments)
        {
            var name = ToolArgumentValidator.RequireString(arguments, "name");
            Logger.Debug($"执行 get_bookmark_text，name={name}");
            return JObject.FromObject(Service.GetBookmarkText(name));
        }
    }

    /// <summary>删除书签。</summary>
    public sealed class DeleteBookmarkTool : ToolBase
    {
        public DeleteBookmarkTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_bookmark";

        public override string Description => "删除指定书签（只删除书签标记，不删除书签内的文本）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["name"] = Str("书签名称"),
        }, "name");

        public override JObject Execute(JObject arguments)
        {
            var name = ToolArgumentValidator.RequireString(arguments, "name");
            Logger.Debug($"执行 delete_bookmark，name={name}");
            Service.DeleteBookmark(name);
            return new JObject { ["deleted"] = true, ["name"] = name };
        }
    }

    /// <summary>插入超链接。</summary>
    public sealed class InsertHyperlinkTool : ToolBase
    {
        public InsertHyperlinkTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_hyperlink";

        public override string Description => "插入超链接：address 为外部链接（URL/邮件），bookmark 为内部锚点（跳转到书签），二者至少提供一个。默认追加到文档末尾（新段落），或用 paragraph_index 插入到指定段落开头。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["text"] = Str("超链接显示文本（缺省使用 address / bookmark）"),
            ["address"] = Str("外部目标地址，如 https://example.com 或 mailto:a@b.com"),
            ["bookmark"] = Str("内部锚点书签名称（可选，需先用 insert_bookmark 创建）"),
            ["screen_tip"] = Str("悬停提示（可选）"),
            ["paragraph_index"] = Int("插入到该段落开头（1 基，可选）；缺省追加到文档末尾", 1),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new InsertHyperlinkRequest
            {
                Text = ToolArgumentValidator.OptionalString(arguments, "text"),
                Address = ToolArgumentValidator.OptionalString(arguments, "address"),
                Bookmark = ToolArgumentValidator.OptionalString(arguments, "bookmark"),
                ScreenTip = ToolArgumentValidator.OptionalString(arguments, "screen_tip"),
                ParagraphIndex = ToolArgumentValidator.OptionalInt(arguments, "paragraph_index", 1),
            };
            Logger.Debug($"执行 insert_hyperlink，address={request.Address}，bookmark={request.Bookmark}");
            return JObject.FromObject(Service.InsertHyperlink(request));
        }
    }

    /// <summary>列出超链接。</summary>
    public sealed class ListHyperlinksTool : ToolBase
    {
        public ListHyperlinksTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "list_hyperlinks";

        public override string Description => "列出文档中的所有超链接（索引、显示文本、外部地址、内部锚点、悬停提示、所在段落）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 list_hyperlinks");
            var links = Service.ListHyperlinks();
            return new JObject
            {
                ["count"] = links.Count,
                ["hyperlinks"] = JArray.FromObject(links),
            };
        }
    }

    /// <summary>修改超链接。</summary>
    public sealed class UpdateHyperlinkTool : ToolBase
    {
        public UpdateHyperlinkTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "update_hyperlink";

        public override string Description => "修改已有超链接（未提供的字段保持不变）：显示文本、外部地址、内部锚点、悬停提示。索引来自 list_hyperlinks。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["index"] = Int("超链接索引（1 基）", 1),
            ["text"] = Str("新的显示文本"),
            ["address"] = Str("新的外部地址"),
            ["bookmark"] = Str("新的内部锚点书签名称"),
            ["screen_tip"] = Str("新的悬停提示"),
        }, "index");

        public override JObject Execute(JObject arguments)
        {
            var request = new UpdateHyperlinkRequest
            {
                Index = ToolArgumentValidator.RequireInt(arguments, "index", 1),
                Text = ToolArgumentValidator.OptionalString(arguments, "text"),
                Address = ToolArgumentValidator.OptionalString(arguments, "address"),
                Bookmark = ToolArgumentValidator.OptionalString(arguments, "bookmark"),
                ScreenTip = ToolArgumentValidator.OptionalString(arguments, "screen_tip"),
            };
            Logger.Debug($"执行 update_hyperlink，index={request.Index}");
            return JObject.FromObject(Service.UpdateHyperlink(request));
        }
    }

    /// <summary>删除超链接。</summary>
    public sealed class DeleteHyperlinkTool : ToolBase
    {
        public DeleteHyperlinkTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_hyperlinks";

        public override string Description => "删除超链接：提供 index 删除单个，或 all=true 删除全部（只删除链接，保留显示文本）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["index"] = Int("要删除的超链接索引（1 基，可选）", 1),
            ["all"] = Bool("是否删除全部超链接"),
        });

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.OptionalInt(arguments, "index", 1);
            var all = ToolArgumentValidator.OptionalBool(arguments, "all") ?? false;
            Logger.Debug($"执行 delete_hyperlinks，index={index}，all={all}");
            var deleted = Service.DeleteHyperlinks(index, all);
            return new JObject
            {
                ["deleted_count"] = deleted,
                ["message"] = deleted == 0 ? "没有可删除的超链接" : $"已删除 {deleted} 个超链接",
            };
        }
    }
}
