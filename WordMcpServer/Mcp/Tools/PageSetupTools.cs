using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>读取页面设置。</summary>
    public sealed class GetPageSetupTool : ToolBase
    {
        public GetPageSetupTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_page_setup";

        public override string Description => "读取指定节的页面设置：纸张大小、方向、页边距（磅）、页眉页脚距离、装订线、分栏数。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["section_index"] = Int("节索引（1 基，默认 1）", 1),
        });

        public override JObject Execute(JObject arguments)
        {
            var sectionIndex = ToolArgumentValidator.OptionalInt(arguments, "section_index", 1) ?? 1;
            Logger.Debug($"执行 get_page_setup，sectionIndex={sectionIndex}");
            return JObject.FromObject(Service.GetPageSetup(sectionIndex));
        }
    }

    /// <summary>设置页面设置。</summary>
    public sealed class SetPageSetupTool : ToolBase
    {
        public SetPageSetupTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_page_setup";

        public override string Description => "设置指定节的页面设置（未提供的参数保持不变）：方向、纸张大小、页边距（磅）、页眉页脚距离、装订线、分栏数。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["section_index"] = Int("节索引（1 基，默认 1）", 1),
            ["orientation"] = Str("页面方向", "portrait", "landscape"),
            ["paper_size"] = Str("纸张大小", "A3", "A4", "A5", "B4", "B5", "Letter", "Legal", "Executive", "Tabloid"),
            ["top_margin"] = Num("上边距（磅）", 0),
            ["bottom_margin"] = Num("下边距（磅）", 0),
            ["left_margin"] = Num("左边距（磅）", 0),
            ["right_margin"] = Num("右边距（磅）", 0),
            ["header_distance"] = Num("页眉距边界（磅）", 0),
            ["footer_distance"] = Num("页脚距边界（磅）", 0),
            ["gutter"] = Num("装订线（磅）", 0),
            ["text_columns"] = Int("分栏数（1-8）", 1, 8),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetPageSetupRequest
            {
                SectionIndex = ToolArgumentValidator.OptionalInt(arguments, "section_index", 1) ?? 1,
                Orientation = ToolArgumentValidator.OptionalString(arguments, "orientation"),
                PaperSize = ToolArgumentValidator.OptionalString(arguments, "paper_size"),
                TopMargin = ToolArgumentValidator.OptionalDouble(arguments, "top_margin", 0),
                BottomMargin = ToolArgumentValidator.OptionalDouble(arguments, "bottom_margin", 0),
                LeftMargin = ToolArgumentValidator.OptionalDouble(arguments, "left_margin", 0),
                RightMargin = ToolArgumentValidator.OptionalDouble(arguments, "right_margin", 0),
                HeaderDistance = ToolArgumentValidator.OptionalDouble(arguments, "header_distance", 0),
                FooterDistance = ToolArgumentValidator.OptionalDouble(arguments, "footer_distance", 0),
                Gutter = ToolArgumentValidator.OptionalDouble(arguments, "gutter", 0),
                TextColumns = ToolArgumentValidator.OptionalInt(arguments, "text_columns", 1, 8),
            };
            Logger.Debug($"执行 set_page_setup，sectionIndex={request.SectionIndex}");
            return JObject.FromObject(Service.SetPageSetup(request));
        }
    }
}
