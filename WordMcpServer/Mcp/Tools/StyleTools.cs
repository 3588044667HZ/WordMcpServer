using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>创建自定义样式。</summary>
    public sealed class CreateStyleTool : ToolBase
    {
        public CreateStyleTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "create_style";

        public override string Description => "创建自定义样式（paragraph / character / table / list 类型），可指定基准样式与字体/段落格式。样式创建后可被 apply_style / add_paragraph 等工具引用。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["name"] = Str("样式名称（不能与现有样式重复）"),
            ["type"] = Str("样式类型", "paragraph", "character", "table", "list"),
            ["base_style"] = Str("基准样式名称（可选，如“正文”）"),
            ["font_name"] = Str("字体名称"),
            ["font_size"] = Num("字号（磅）", 1, 400),
            ["bold"] = Bool("加粗"),
            ["italic"] = Bool("斜体"),
            ["color"] = Str("字体颜色：#RRGGBB / 颜色名"),
            ["alignment"] = Str("段落对齐", "left", "center", "right", "justify"),
            ["space_before"] = Num("段前间距（磅）", 0),
            ["space_after"] = Num("段后间距（磅）", 0),
            ["line_spacing"] = Str("行距", "single", "1.5", "double"),
        }, "name");

        public override JObject Execute(JObject arguments)
        {
            var request = BuildRequest(arguments, true);
            Logger.Debug($"执行 create_style，name={request.Name}");
            return JObject.FromObject(Service.CreateStyle(request));
        }

        internal static StyleFormatRequest BuildRequest(JObject arguments, bool requireName)
        {
            return new StyleFormatRequest
            {
                Name = requireName
                    ? ToolArgumentValidator.RequireString(arguments, "name")
                    : ToolArgumentValidator.OptionalString(arguments, "name"),
                Type = ToolArgumentValidator.OptionalString(arguments, "type"),
                BaseStyle = ToolArgumentValidator.OptionalString(arguments, "base_style"),
                FontName = ToolArgumentValidator.OptionalString(arguments, "font_name"),
                FontSize = ToolArgumentValidator.OptionalDouble(arguments, "font_size", 1, 400),
                Bold = ToolArgumentValidator.OptionalBool(arguments, "bold"),
                Italic = ToolArgumentValidator.OptionalBool(arguments, "italic"),
                ColorArgb = ToolArgumentValidator.OptionalColor(arguments, "color"),
                Alignment = ToolArgumentValidator.OptionalString(arguments, "alignment"),
                SpaceBefore = ToolArgumentValidator.OptionalDouble(arguments, "space_before", 0),
                SpaceAfter = ToolArgumentValidator.OptionalDouble(arguments, "space_after", 0),
                LineSpacing = ToolArgumentValidator.OptionalString(arguments, "line_spacing"),
            };
        }
    }

    /// <summary>修改现有样式。</summary>
    public sealed class ModifyStyleTool : ToolBase
    {
        public ModifyStyleTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "modify_style";

        public override string Description => "修改现有样式（未提供的格式参数保持不变）：可改基准样式、字体、段落格式。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["name"] = Str("样式名称（支持本地化名称，如“标题 1”）"),
            ["base_style"] = Str("基准样式名称（可选）"),
            ["font_name"] = Str("字体名称"),
            ["font_size"] = Num("字号（磅）", 1, 400),
            ["bold"] = Bool("加粗"),
            ["italic"] = Bool("斜体"),
            ["color"] = Str("字体颜色：#RRGGBB / 颜色名"),
            ["alignment"] = Str("段落对齐", "left", "center", "right", "justify"),
            ["space_before"] = Num("段前间距（磅）", 0),
            ["space_after"] = Num("段后间距（磅）", 0),
            ["line_spacing"] = Str("行距", "single", "1.5", "double"),
        }, "name");

        public override JObject Execute(JObject arguments)
        {
            var request = CreateStyleTool.BuildRequest(arguments, true);
            Logger.Debug($"执行 modify_style，name={request.Name}");
            return JObject.FromObject(Service.ModifyStyle(request));
        }
    }

    /// <summary>删除自定义样式。</summary>
    public sealed class DeleteStyleTool : ToolBase
    {
        public DeleteStyleTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_style";

        public override string Description => "删除自定义样式（内置样式不能删除）；使用该样式的文本会回到基准样式。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["name"] = Str("样式名称"),
        }, "name");

        public override JObject Execute(JObject arguments)
        {
            var name = ToolArgumentValidator.RequireString(arguments, "name");
            Logger.Debug($"执行 delete_style，name={name}");
            Service.DeleteStyle(name);
            return new JObject { ["deleted"] = true, ["name"] = name };
        }
    }
}
