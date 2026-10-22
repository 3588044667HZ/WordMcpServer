using System.Collections.Generic;
using Newtonsoft.Json;

namespace WordMcpServer.Services
{
    // ============================== DTO 定义 ==============================
    // 说明：所有 DTO 字段以 snake_case 序列化（MCP 工具参数/返回约定），
    // 由工具层 JObject.FromObject 直接输出。

    /// <summary>文档基本信息。</summary>
    public sealed class DocumentInfo
    {
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("path")] public string Path { get; set; }
        [JsonProperty("full_name")] public string FullName { get; set; }
        [JsonProperty("saved")] public bool Saved { get; set; }
        [JsonProperty("read_only")] public bool ReadOnly { get; set; }
        [JsonProperty("paragraph_count")] public int ParagraphCount { get; set; }
        [JsonProperty("table_count")] public int TableCount { get; set; }
        [JsonProperty("section_count")] public int SectionCount { get; set; }
        [JsonProperty("page_count")] public int PageCount { get; set; }
        [JsonProperty("word_count")] public int WordCount { get; set; }
        [JsonProperty("character_count")] public int CharacterCount { get; set; }
    }

    /// <summary>已打开文档的摘要（list_documents）。</summary>
    public sealed class DocumentSummary
    {
        [JsonProperty("index")] public int Index { get; set; }
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("path")] public string Path { get; set; }
        [JsonProperty("saved")] public bool Saved { get; set; }
        [JsonProperty("read_only")] public bool ReadOnly { get; set; }
        [JsonProperty("active")] public bool Active { get; set; }
    }

    /// <summary>文档纯文本内容。</summary>
    public sealed class DocumentText
    {
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("character_count")] public int CharacterCount { get; set; }
        [JsonProperty("paragraph_count")] public int ParagraphCount { get; set; }
        [JsonProperty("truncated")] public bool Truncated { get; set; }
    }

    /// <summary>大纲条目（标题段落）。</summary>
    public sealed class OutlineItem
    {
        [JsonProperty("paragraph_index")] public int ParagraphIndex { get; set; }
        [JsonProperty("level")] public int Level { get; set; }
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("style")] public string Style { get; set; }
    }

    /// <summary>文档大纲。</summary>
    public sealed class DocumentOutline
    {
        [JsonProperty("heading_count")] public int HeadingCount { get; set; }
        [JsonProperty("items")] public List<OutlineItem> Items { get; set; } = new List<OutlineItem>();
        [JsonProperty("truncated")] public bool Truncated { get; set; }
    }

    /// <summary>文档统计。</summary>
    public sealed class DocumentStatistics
    {
        [JsonProperty("pages")] public int Pages { get; set; }
        [JsonProperty("words")] public int Words { get; set; }
        [JsonProperty("characters")] public int Characters { get; set; }
        [JsonProperty("characters_with_spaces")] public int CharactersWithSpaces { get; set; }
        [JsonProperty("paragraphs")] public int Paragraphs { get; set; }
        [JsonProperty("lines")] public int Lines { get; set; }
    }

    /// <summary>段落信息。</summary>
    public sealed class ParagraphInfo
    {
        [JsonProperty("paragraph_index")] public int ParagraphIndex { get; set; }
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("style")] public string Style { get; set; }
        [JsonProperty("alignment")] public string Alignment { get; set; }
        [JsonProperty("paragraph_count")] public int ParagraphCount { get; set; }
    }

    /// <summary>样式信息。</summary>
    public sealed class StyleInfo
    {
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("type")] public string Type { get; set; }
        [JsonProperty("built_in")] public bool BuiltIn { get; set; }
        [JsonProperty("in_use")] public bool InUse { get; set; }
    }

    /// <summary>应用样式的结果。</summary>
    public sealed class StyleApplyResult
    {
        [JsonProperty("style")] public string Style { get; set; }
        [JsonProperty("applied_count")] public int AppliedCount { get; set; }
        [JsonProperty("paragraph_index")] public int? ParagraphIndex { get; set; }
        [JsonProperty("paragraph_count")] public int ParagraphCount { get; set; }
    }

    /// <summary>应用文本样式的结果。</summary>
    public sealed class TextStyleResult
    {
        [JsonProperty("styled_count")] public int StyledCount { get; set; }
        [JsonProperty("paragraph_index")] public int? ParagraphIndex { get; set; }
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("paragraph_count")] public int ParagraphCount { get; set; }
    }

    /// <summary>备份结果。</summary>
    public sealed class BackupResult
    {
        [JsonProperty("backup_path")] public string BackupPath { get; set; }
        [JsonProperty("size_bytes")] public long SizeBytes { get; set; }
        [JsonProperty("source_path")] public string SourcePath { get; set; }

        /// <summary>file_copy（已保存文档直接复制）或 save_as（未保存文档另存）。</summary>
        [JsonProperty("method")] public string Method { get; set; }
    }

    /// <summary>PDF 导出结果。</summary>
    public sealed class ExportResult
    {
        [JsonProperty("output_path")] public string OutputPath { get; set; }
        [JsonProperty("size_bytes")] public long SizeBytes { get; set; }
        [JsonProperty("page_count")] public int PageCount { get; set; }
    }

    /// <summary>表格信息。</summary>
    public sealed class TableInfo
    {
        [JsonProperty("table_index")] public int TableIndex { get; set; }
        [JsonProperty("rows")] public int Rows { get; set; }
        [JsonProperty("columns")] public int Columns { get; set; }
        [JsonProperty("style")] public string Style { get; set; }
        [JsonProperty("alignment")] public string Alignment { get; set; }
        [JsonProperty("table_count")] public int TableCount { get; set; }
        [JsonProperty("column_widths")] public List<double> ColumnWidths { get; set; }
    }

    /// <summary>表格数据（二维数组）。</summary>
    public sealed class TableData
    {
        [JsonProperty("table_index")] public int TableIndex { get; set; }
        [JsonProperty("rows")] public int Rows { get; set; }
        [JsonProperty("columns")] public int Columns { get; set; }
        [JsonProperty("data")] public List<List<string>> Data { get; set; } = new List<List<string>>();
    }

    /// <summary>查找/替换结果。</summary>
    public sealed class FindReplaceResult
    {
        [JsonProperty("find_text")] public string FindText { get; set; }
        [JsonProperty("replace_text")] public string ReplaceText { get; set; }
        [JsonProperty("replacement_count")] public int ReplacementCount { get; set; }
        [JsonProperty("message")] public string Message { get; set; }
    }

    /// <summary>查找匹配项。</summary>
    public sealed class FindMatch
    {
        [JsonProperty("paragraph_index")] public int ParagraphIndex { get; set; }
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("start")] public int Start { get; set; }
        [JsonProperty("end")] public int End { get; set; }
    }

    /// <summary>查找结果。</summary>
    public sealed class FindResult
    {
        [JsonProperty("find_text")] public string FindText { get; set; }
        [JsonProperty("match_count")] public int MatchCount { get; set; }
        [JsonProperty("matches")] public List<FindMatch> Matches { get; set; } = new List<FindMatch>();
    }

    /// <summary>书签信息。</summary>
    public sealed class BookmarkInfo
    {
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("exists")] public bool Exists { get; set; }
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("start")] public int Start { get; set; }
        [JsonProperty("end")] public int End { get; set; }
    }

    /// <summary>超链接信息。</summary>
    public sealed class HyperlinkInfo
    {
        [JsonProperty("index")] public int Index { get; set; }
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("address")] public string Address { get; set; }
        [JsonProperty("sub_address")] public string SubAddress { get; set; }
        [JsonProperty("screen_tip")] public string ScreenTip { get; set; }
        [JsonProperty("paragraph_index")] public int ParagraphIndex { get; set; }
        [JsonProperty("hyperlink_count")] public int HyperlinkCount { get; set; }
    }

    /// <summary>图片信息。</summary>
    public sealed class ImageInfo
    {
        [JsonProperty("path")] public string Path { get; set; }
        [JsonProperty("width")] public double Width { get; set; }
        [JsonProperty("height")] public double Height { get; set; }
        [JsonProperty("paragraph_index")] public int ParagraphIndex { get; set; }
        [JsonProperty("inline_shape_count")] public int InlineShapeCount { get; set; }
    }

    /// <summary>目录信息。</summary>
    public sealed class TocInfo
    {
        [JsonProperty("toc_count")] public int TocCount { get; set; }
        [JsonProperty("upper_level")] public int UpperLevel { get; set; }
        [JsonProperty("lower_level")] public int LowerLevel { get; set; }
        [JsonProperty("updated")] public bool Updated { get; set; }
    }

    /// <summary>页眉页脚信息。</summary>
    public sealed class HeaderFooterInfo
    {
        [JsonProperty("section_index")] public int SectionIndex { get; set; }
        [JsonProperty("header_text")] public string HeaderText { get; set; }
        [JsonProperty("footer_text")] public string FooterText { get; set; }
        [JsonProperty("section_count")] public int SectionCount { get; set; }
    }

    /// <summary>批注信息。</summary>
    public sealed class CommentInfo
    {
        [JsonProperty("index")] public int Index { get; set; }
        [JsonProperty("author")] public string Author { get; set; }
        [JsonProperty("initial")] public string Initial { get; set; }
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("scope_text")] public string ScopeText { get; set; }
    }

    /// <summary>页面设置信息。</summary>
    public sealed class PageSetupInfo
    {
        [JsonProperty("section_index")] public int SectionIndex { get; set; }
        [JsonProperty("page_width")] public double PageWidth { get; set; }
        [JsonProperty("page_height")] public double PageHeight { get; set; }
        [JsonProperty("orientation")] public string Orientation { get; set; }
        [JsonProperty("paper_size")] public string PaperSize { get; set; }
        [JsonProperty("top_margin")] public double TopMargin { get; set; }
        [JsonProperty("bottom_margin")] public double BottomMargin { get; set; }
        [JsonProperty("left_margin")] public double LeftMargin { get; set; }
        [JsonProperty("right_margin")] public double RightMargin { get; set; }
        [JsonProperty("header_distance")] public double HeaderDistance { get; set; }
        [JsonProperty("footer_distance")] public double FooterDistance { get; set; }
        [JsonProperty("gutter")] public double Gutter { get; set; }
        [JsonProperty("text_columns")] public int TextColumns { get; set; }
        [JsonProperty("section_count")] public int SectionCount { get; set; }
    }

    // ============================== 请求参数 ==============================

    /// <summary>add_paragraph 请求。</summary>
    public sealed class AddParagraphRequest
    {
        public string Text { get; set; }

        /// <summary>插入位置（1 基段落索引，插入到该段落之前）；为空表示追加到文档末尾。</summary>
        public int? Position { get; set; }

        /// <summary>段落样式（内置别名如 heading1-9 / normal / title，或样式名称）。</summary>
        public string Style { get; set; }
    }

    /// <summary>set_paragraph_format 请求。</summary>
    public sealed class SetParagraphFormatRequest
    {
        public int? ParagraphIndex { get; set; }
        public string FindText { get; set; }
        public string Alignment { get; set; }
        public double? LeftIndent { get; set; }
        public double? RightIndent { get; set; }
        public double? FirstLineIndent { get; set; }
        public double? SpaceBefore { get; set; }
        public double? SpaceAfter { get; set; }

        /// <summary>行距模式：single / 1.5 / double / at-least / exactly / multiple。</summary>
        public string LineSpacing { get; set; }

        /// <summary>行距数值（磅）：at-least/exactly/multiple 时使用。</summary>
        public double? LineSpacingValue { get; set; }

        public bool? KeepWithNext { get; set; }
        public bool? KeepTogether { get; set; }
        public bool? PageBreakBefore { get; set; }

        /// <summary>大纲级别：body-text 或 level-1 .. level-9。</summary>
        public string OutlineLevel { get; set; }
    }

    /// <summary>set_text_style 请求。</summary>
    public sealed class SetTextStyleRequest
    {
        public int? ParagraphIndex { get; set; }

        /// <summary>查找文本：对该文本的所有出现位置应用样式。</summary>
        public string FindText { get; set; }

        public string FontName { get; set; }
        public double? FontSize { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public bool? Underline { get; set; }

        /// <summary>字体颜色（ARGB 整型，工具层已解析）。</summary>
        public int? ColorArgb { get; set; }

        /// <summary>突出显示颜色：yellow / bright-green / turquoise / pink / blue / red / dark-blue / teal / violet / dark-red / dark-yellow / gray / black / white / none。</summary>
        public string Highlight { get; set; }

        public bool? AllCaps { get; set; }
        public bool? Subscript { get; set; }
        public bool? Superscript { get; set; }
    }

    /// <summary>apply_style 请求。</summary>
    public sealed class ApplyStyleRequest
    {
        public int? ParagraphIndex { get; set; }
        public string FindText { get; set; }

        /// <summary>true 表示应用到整篇文档的所有段落。</summary>
        public bool? All { get; set; }

        /// <summary>样式（内置别名或样式名称）。</summary>
        public string Style { get; set; }
    }

    /// <summary>insert_table 请求。</summary>
    public sealed class InsertTableRequest
    {
        public int Rows { get; set; }
        public int Columns { get; set; }
        public List<List<string>> Data { get; set; }

        /// <summary>表格样式名称（可选）。</summary>
        public string Style { get; set; }

        /// <summary>插入到该段落之前（1 基）；为空表示追加到文档末尾。</summary>
        public int? Position { get; set; }

        /// <summary>是否自动适应窗口宽度（默认 true）。</summary>
        public bool AutoFit { get; set; } = true;
    }

    /// <summary>insert_table_rows 请求。</summary>
    public sealed class InsertTableRowsRequest
    {
        public int TableIndex { get; set; }
        public int Count { get; set; } = 1;

        /// <summary>在该行之前插入（1 基）；为空表示追加到表格末尾。</summary>
        public int? BeforeRow { get; set; }

        public List<List<string>> Data { get; set; }
    }

    /// <summary>set_table_cell_style 请求。</summary>
    public sealed class SetTableCellStyleRequest
    {
        public int TableIndex { get; set; }

        /// <summary>目标行（1 基）；为空表示所有行。</summary>
        public int? Row { get; set; }

        /// <summary>目标列（1 基）；为空表示所有列。</summary>
        public int? Column { get; set; }

        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public string FontName { get; set; }
        public double? FontSize { get; set; }
        public int? ColorArgb { get; set; }

        /// <summary>单元格底纹颜色（ARGB 整型）。</summary>
        public int? ShadingArgb { get; set; }

        /// <summary>垂直对齐：top / center / bottom。</summary>
        public string VerticalAlignment { get; set; }

        /// <summary>段落水平对齐：left / center / right / justify。</summary>
        public string Alignment { get; set; }
    }

    /// <summary>set_table_layout 请求。</summary>
    public sealed class SetTableLayoutRequest
    {
        public int TableIndex { get; set; }

        /// <summary>表格对齐：left / center / right。</summary>
        public string Alignment { get; set; }

        /// <summary>自动调整：fixed / content / window。</summary>
        public string AutoFit { get; set; }

        /// <summary>各列宽度（磅），按列顺序提供。</summary>
        public List<double> ColumnWidths { get; set; }

        public bool? RepeatHeaderRow { get; set; }
        public bool? AllowBreakAcrossPages { get; set; }
    }

    /// <summary>find_replace_text 请求。</summary>
    public sealed class FindReplaceRequest
    {
        public string FindText { get; set; }
        public string ReplaceText { get; set; }
        public bool MatchCase { get; set; }
        public bool MatchWholeWord { get; set; }

        /// <summary>true 替换全部（默认），false 只替换第一处。</summary>
        public bool ReplaceAll { get; set; } = true;
    }

    /// <summary>find_text 请求。</summary>
    public sealed class FindTextRequest
    {
        public string Text { get; set; }
        public bool MatchCase { get; set; }
        public bool MatchWholeWord { get; set; }
        public int MaxResults { get; set; } = 20;
    }

    /// <summary>insert_image 请求。</summary>
    public sealed class InsertImageRequest
    {
        public string Path { get; set; }

        /// <summary>宽度（磅）；与 height 只给一个时等比缩放。</summary>
        public double? Width { get; set; }

        /// <summary>高度（磅）。</summary>
        public double? Height { get; set; }

        /// <summary>图片所在段落的对齐：left / center / right。</summary>
        public string Alignment { get; set; }

        /// <summary>插入到该段落之前（1 基）；为空表示追加到文档末尾。</summary>
        public int? ParagraphIndex { get; set; }
    }

    /// <summary>insert_hyperlink 请求。</summary>
    public sealed class InsertHyperlinkRequest
    {
        public string Text { get; set; }
        public string Address { get; set; }

        /// <summary>内部锚点书签名称（与 address 二选一）。</summary>
        public string Bookmark { get; set; }

        public string ScreenTip { get; set; }

        /// <summary>插入到该段落之前（1 基）；为空表示追加到文档末尾。</summary>
        public int? ParagraphIndex { get; set; }
    }

    /// <summary>insert_toc 请求。</summary>
    public sealed class InsertTocRequest
    {
        /// <summary>包含的标题级别数（1-9，默认 3）。</summary>
        public int Levels { get; set; } = 3;

        /// <summary>目录标题文字（可选，作为独立段落插在目录前）。</summary>
        public string Title { get; set; }
    }

    /// <summary>set_header_footer 请求。</summary>
    public sealed class SetHeaderFooterRequest
    {
        /// <summary>节索引（1 基，默认 1）。</summary>
        public int SectionIndex { get; set; } = 1;

        public string HeaderText { get; set; }
        public string FooterText { get; set; }

        /// <summary>是否在页脚附加页码域（默认 false）。</summary>
        public bool IncludePageNumber { get; set; }
    }

    /// <summary>set_page_setup 请求。</summary>
    public sealed class SetPageSetupRequest
    {
        /// <summary>节索引（1 基，默认 1）。</summary>
        public int SectionIndex { get; set; } = 1;

        /// <summary>portrait / landscape。</summary>
        public string Orientation { get; set; }

        /// <summary>A4 / A3 / A5 / Letter / Legal / B5。</summary>
        public string PaperSize { get; set; }

        public double? TopMargin { get; set; }
        public double? BottomMargin { get; set; }
        public double? LeftMargin { get; set; }
        public double? RightMargin { get; set; }
        public double? HeaderDistance { get; set; }
        public double? FooterDistance { get; set; }
        public double? Gutter { get; set; }

        /// <summary>分栏数（1-8）。</summary>
        public int? TextColumns { get; set; }
    }

    // ============================== 扩展 DTO（P0） ==============================

    /// <summary>段落详细信息（含格式回读）。</summary>
    public sealed class ParagraphDetail
    {
        [JsonProperty("paragraph_index")] public int ParagraphIndex { get; set; }
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("style")] public string Style { get; set; }
        [JsonProperty("alignment")] public string Alignment { get; set; }
        [JsonProperty("outline_level")] public string OutlineLevel { get; set; }
        [JsonProperty("in_table")] public bool InTable { get; set; }
        [JsonProperty("left_indent")] public double LeftIndent { get; set; }
        [JsonProperty("right_indent")] public double RightIndent { get; set; }
        [JsonProperty("first_line_indent")] public double FirstLineIndent { get; set; }
        [JsonProperty("space_before")] public double SpaceBefore { get; set; }
        [JsonProperty("space_after")] public double SpaceAfter { get; set; }
        [JsonProperty("line_spacing_rule")] public string LineSpacingRule { get; set; }
        [JsonProperty("line_spacing")] public double LineSpacing { get; set; }
        [JsonProperty("keep_with_next")] public bool? KeepWithNext { get; set; }
        [JsonProperty("keep_together")] public bool? KeepTogether { get; set; }
        [JsonProperty("page_break_before")] public bool? PageBreakBefore { get; set; }
        [JsonProperty("font_name")] public string FontName { get; set; }
        [JsonProperty("font_size")] public double? FontSize { get; set; }
        [JsonProperty("bold")] public bool? Bold { get; set; }
        [JsonProperty("italic")] public bool? Italic { get; set; }
        [JsonProperty("underline")] public bool? Underline { get; set; }
        [JsonProperty("font_color")] public string FontColor { get; set; }
        [JsonProperty("highlight")] public string Highlight { get; set; }
        [JsonProperty("list_type")] public string ListType { get; set; }
        [JsonProperty("list_level")] public int? ListLevel { get; set; }
        [JsonProperty("paragraph_count")] public int ParagraphCount { get; set; }
    }

    /// <summary>删除类操作结果。</summary>
    public sealed class DeleteResult
    {
        [JsonProperty("deleted")] public bool Deleted { get; set; }
        [JsonProperty("deleted_count")] public int DeletedCount { get; set; }
        [JsonProperty("paragraph_index")] public int? ParagraphIndex { get; set; }
        [JsonProperty("paragraph_count")] public int ParagraphCount { get; set; }
        [JsonProperty("message")] public string Message { get; set; }
    }

    /// <summary>域更新结果。</summary>
    public sealed class FieldUpdateResult
    {
        [JsonProperty("fields_updated")] public bool FieldsUpdated { get; set; }
        [JsonProperty("toc_count")] public int TocCount { get; set; }
        [JsonProperty("section_count")] public int SectionCount { get; set; }
        [JsonProperty("message")] public string Message { get; set; }
    }

    /// <summary>文档属性集合（内置 + 自定义）。</summary>
    public sealed class DocumentPropertySet
    {
        [JsonProperty("title")] public string Title { get; set; }
        [JsonProperty("author")] public string Author { get; set; }
        [JsonProperty("subject")] public string Subject { get; set; }
        [JsonProperty("keywords")] public string Keywords { get; set; }
        [JsonProperty("comments")] public string Comments { get; set; }
        [JsonProperty("category")] public string Category { get; set; }
        [JsonProperty("company")] public string Company { get; set; }
        [JsonProperty("manager")] public string Manager { get; set; }
        [JsonProperty("last_author")] public string LastAuthor { get; set; }
        [JsonProperty("revision_number")] public string RevisionNumber { get; set; }
        [JsonProperty("custom")] public Dictionary<string, string> Custom { get; set; } = new Dictionary<string, string>();
    }

    /// <summary>set_document_properties 请求（null 字段保持不变）。</summary>
    public sealed class SetDocumentPropertiesRequest
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public string Subject { get; set; }
        public string Keywords { get; set; }
        public string Comments { get; set; }
        public string Category { get; set; }
        public string Company { get; set; }
        public string Manager { get; set; }
        public Dictionary<string, string> Custom { get; set; }
    }

    /// <summary>富文本 run 定义（add_paragraph_runs）。</summary>
    public sealed class TextRunSpec
    {
        public string Text { get; set; }
        public string FontName { get; set; }
        public double? FontSize { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public bool? Underline { get; set; }
        public int? ColorArgb { get; set; }
        public string Highlight { get; set; }
        public bool? Superscript { get; set; }
        public bool? Subscript { get; set; }
    }

    /// <summary>add_paragraph_runs 请求。</summary>
    public sealed class AddParagraphRunsRequest
    {
        public int? Position { get; set; }
        public string Style { get; set; }
        public List<TextRunSpec> Runs { get; set; }
    }

    /// <summary>set_list_format 请求。</summary>
    public sealed class SetListFormatRequest
    {
        /// <summary>bullet / number / outline / none。</summary>
        public string Type { get; set; }

        /// <summary>单个目标段落（1 基）；缺省为最后一个段落。</summary>
        public int? ParagraphIndex { get; set; }

        /// <summary>起始段落（与 end_paragraph 配合，1 基）。</summary>
        public int? StartParagraph { get; set; }

        /// <summary>结束段落（1 基，含）。</summary>
        public int? EndParagraph { get; set; }

        /// <summary>列表级别（1-9，默认 1）。</summary>
        public int? Level { get; set; }
    }

    /// <summary>set_row_height 请求。</summary>
    public sealed class SetRowHeightRequest
    {
        public int TableIndex { get; set; }

        /// <summary>目标行（1 基）；为空表示所有行。</summary>
        public int? Row { get; set; }

        /// <summary>行高（磅）；rule=auto 时可省略。</summary>
        public double? Height { get; set; }

        /// <summary>auto / at-least / exactly（默认 at-least）。</summary>
        public string Rule { get; set; }
    }

    /// <summary>insert_table_columns 请求。</summary>
    public sealed class InsertTableColumnsRequest
    {
        public int TableIndex { get; set; }
        public int Count { get; set; } = 1;

        /// <summary>在该列之前插入（1 基）；为空表示追加到末尾。</summary>
        public int? BeforeColumn { get; set; }

        public List<List<string>> Data { get; set; }
    }

    /// <summary>create_style / modify_style 请求。</summary>
    public sealed class StyleFormatRequest
    {
        /// <summary>样式名称（create 时必填，modify 时用于定位）。</summary>
        public string Name { get; set; }

        /// <summary>样式类型：paragraph / character / table / list（默认 paragraph）。</summary>
        public string Type { get; set; }

        /// <summary>基准样式名称（可选）。</summary>
        public string BaseStyle { get; set; }

        public string FontName { get; set; }
        public double? FontSize { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public int? ColorArgb { get; set; }
        public string Alignment { get; set; }
        public double? SpaceBefore { get; set; }
        public double? SpaceAfter { get; set; }

        /// <summary>single / 1.5 / double。</summary>
        public string LineSpacing { get; set; }
    }

    /// <summary>水印操作结果。</summary>
    public sealed class WatermarkResult
    {
        [JsonProperty("applied")] public bool Applied { get; set; }
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("layout")] public string Layout { get; set; }
        [JsonProperty("section_count")] public int SectionCount { get; set; }
        [JsonProperty("removed_count")] public int RemovedCount { get; set; }
    }

    /// <summary>add_watermark 请求。</summary>
    public sealed class AddWatermarkRequest
    {
        public string Text { get; set; }
        public string FontName { get; set; }
        public double? FontSize { get; set; }
        public int? ColorArgb { get; set; }

        /// <summary>diagonal（斜向，默认）/ horizontal（水平）。</summary>
        public string Layout { get; set; }

        /// <summary>透明度 0-1（默认 0.5）。</summary>
        public double? Transparency { get; set; }
    }

    /// <summary>修订操作结果。</summary>
    public sealed class RevisionResult
    {
        [JsonProperty("track_changes")] public bool TrackChanges { get; set; }
        [JsonProperty("revision_count")] public int RevisionCount { get; set; }
        [JsonProperty("message")] public string Message { get; set; }
    }

    /// <summary>set_paragraph_shading 请求。</summary>
    public sealed class SetParagraphShadingRequest
    {
        public int? ParagraphIndex { get; set; }
        public string FindText { get; set; }

        /// <summary>底纹颜色（ARGB 整型）。</summary>
        public int? ShadingArgb { get; set; }

        /// <summary>边框颜色（ARGB 整型）。</summary>
        public int? BorderArgb { get; set; }

        /// <summary>边框样式：single / double / dotted / dashed / none（默认 single）。</summary>
        public string BorderStyle { get; set; }

        /// <summary>边框宽度：thin / medium / thick 或磅值（如 1.5）。</summary>
        public string BorderWidth { get; set; }
    }

    // ============================== 论文增强 DTO ==============================

    /// <summary>脚注/尾注信息。</summary>
    public sealed class NoteInfo
    {
        [JsonProperty("note_type")] public string NoteType { get; set; }
        [JsonProperty("index")] public int Index { get; set; }
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("reference_paragraph_index")] public int ReferenceParagraphIndex { get; set; }
        [JsonProperty("reference_text")] public string ReferenceText { get; set; }
        [JsonProperty("note_count")] public int NoteCount { get; set; }
    }

    /// <summary>脚注/尾注编号选项。</summary>
    public sealed class NoteOptionsInfo
    {
        [JsonProperty("note_type")] public string NoteType { get; set; }
        [JsonProperty("number_style")] public string NumberStyle { get; set; }
        [JsonProperty("numbering_rule")] public string NumberingRule { get; set; }
        [JsonProperty("starting_number")] public int StartingNumber { get; set; }
        [JsonProperty("location")] public string Location { get; set; }
        [JsonProperty("note_count")] public int NoteCount { get; set; }
    }

    /// <summary>set_note_options 请求。</summary>
    public sealed class SetNoteOptionsRequest
    {
        /// <summary>footnote / endnote。</summary>
        public string NoteType { get; set; }

        /// <summary>arabic / upper-roman / lower-roman / upper-letter / lower-letter。</summary>
        public string NumberStyle { get; set; }

        /// <summary>continuous / restart-section / restart-page。</summary>
        public string NumberingRule { get; set; }

        public int? StartingNumber { get; set; }

        /// <summary>仅脚注：bottom / beneath-text。</summary>
        public string Location { get; set; }
    }

    /// <summary>insert_cross_reference 请求。</summary>
    public sealed class CrossReferenceRequest
    {
        /// <summary>heading / bookmark / footnote / endnote / caption。</summary>
        public string ReferenceType { get; set; }

        /// <summary>标题文本、书签名、脚注/尾注编号、题注编号。</summary>
        public string ReferenceItem { get; set; }

        /// <summary>caption 类型的题注标签（如 图 / 表）。</summary>
        public string CaptionLabel { get; set; }

        /// <summary>text / number / page / label-and-number / entire-caption（缺省按类型）。</summary>
        public string ReferenceKind { get; set; }

        /// <summary>是否插入为可点击超链接（默认 true）。</summary>
        public bool InsertAsHyperlink { get; set; } = true;

        /// <summary>插入到该段落末尾（1 基）；缺省追加到文档末尾。</summary>
        public int? ParagraphIndex { get; set; }
    }

    /// <summary>交叉引用结果。</summary>
    public sealed class CrossReferenceResult
    {
        [JsonProperty("inserted")] public bool Inserted { get; set; }
        [JsonProperty("reference_type")] public string ReferenceType { get; set; }
        [JsonProperty("reference_kind")] public string ReferenceKind { get; set; }
        [JsonProperty("paragraph_index")] public int ParagraphIndex { get; set; }
        [JsonProperty("paragraph_count")] public int ParagraphCount { get; set; }
    }

    /// <summary>insert_caption 请求。</summary>
    public sealed class InsertCaptionRequest
    {
        /// <summary>题注标签（图 / 表 / Figure / Table）。</summary>
        public string Label { get; set; } = "图";

        /// <summary>题注文字（可选）。</summary>
        public string Title { get; set; }

        /// <summary>above / below（默认 below）。</summary>
        public string Position { get; set; } = "below";

        /// <summary>题注插入到该段落之后（below）或之前（above）；缺省最后一个段落。</summary>
        public int? ParagraphIndex { get; set; }
    }

    /// <summary>题注结果。</summary>
    public sealed class CaptionResult
    {
        [JsonProperty("label")] public string Label { get; set; }
        [JsonProperty("title")] public string Title { get; set; }
        [JsonProperty("position")] public string Position { get; set; }
        [JsonProperty("paragraph_index")] public int ParagraphIndex { get; set; }
        [JsonProperty("paragraph_count")] public int ParagraphCount { get; set; }
    }

    /// <summary>insert_field 请求。</summary>
    public sealed class InsertFieldRequest
    {
        /// <summary>域代码，如 "PAGEREF bm \\h"、"SEQ 图 \\* ARABIC"、"DATE"。</summary>
        public string FieldCode { get; set; }

        /// <summary>插入到该段落末尾（1 基）；缺省追加到文档末尾。</summary>
        public int? ParagraphIndex { get; set; }

        /// <summary>插入后是否立即更新域（默认 true）。</summary>
        public bool Update { get; set; } = true;
    }

    /// <summary>域信息。</summary>
    public sealed class FieldInfoResult
    {
        [JsonProperty("index")] public int Index { get; set; }
        [JsonProperty("code")] public string Code { get; set; }
        [JsonProperty("result")] public string Result { get; set; }
        [JsonProperty("field_count")] public int FieldCount { get; set; }
    }

    /// <summary>update_hyperlink 请求。</summary>
    public sealed class UpdateHyperlinkRequest
    {
        public int Index { get; set; }
        public string Text { get; set; }
        public string Address { get; set; }

        /// <summary>书签锚点（内部链接）。</summary>
        public string Bookmark { get; set; }

        public string ScreenTip { get; set; }
    }

    // ============================== 服务契约 ==============================
    /// <summary>
    /// Word 对象模型服务。所有方法内部通过 OfficeThreadMarshaller 封送到 Word 主线程执行，
    /// 调用方（Kestrel 线程池线程）可以安全地直接调用。
    /// 失败时抛出 McpToolException（业务错误码见 McpErrorCodes）。
    /// </summary>
    public interface IWordService
    {
        // ---- 文档 ----
        DocumentInfo GetActiveDocumentInfo();
        List<DocumentSummary> ListDocuments();
        DocumentInfo NewDocument();
        DocumentInfo ActivateDocument(string name, int? index);
        DocumentInfo OpenDocument(string path, bool readOnly);
        DocumentInfo CloseDocument(string name, int? index, bool saveChanges);
        DocumentInfo SaveDocument(string saveAsPath, string format);
        BackupResult BackupDocument(string backupPath);
        ExportResult ExportToPdf(string outputPath, bool openAfterExport);
        DocumentText GetDocumentText(int maxChars);
        DocumentOutline GetDocumentOutline(int maxItems);
        DocumentStatistics GetDocumentStatistics();

        // ---- 段落与文本 ----
        ParagraphInfo AddParagraph(AddParagraphRequest request);
        ParagraphInfo AddHeading(string text, int level, int? position);
        ParagraphInfo SetParagraphFormat(SetParagraphFormatRequest request);
        TextStyleResult SetTextStyle(SetTextStyleRequest request);
        ParagraphInfo InsertPageBreak(int? beforeParagraph, string text);
        ParagraphInfo InsertSectionBreak(int? beforeParagraph, string text);
        StyleApplyResult ApplyStyle(ApplyStyleRequest request);
        List<StyleInfo> ListStyles();

        // ---- 表格 ----
        TableInfo InsertTable(InsertTableRequest request);
        TableData GetTableData(int tableIndex);
        TableData UpdateTableData(int tableIndex, List<List<string>> data);
        TableInfo InsertTableRows(InsertTableRowsRequest request);
        TableInfo DeleteTableRow(int tableIndex, int rowIndex);
        TableInfo MergeTableCells(int tableIndex, int startRow, int startColumn, int endRow, int endColumn);
        TableInfo SetTableCellStyle(SetTableCellStyleRequest request);
        TableInfo SetTableLayout(SetTableLayoutRequest request);

        // ---- 查找与链接 ----
        FindReplaceResult FindReplaceText(FindReplaceRequest request);
        FindResult FindText(FindTextRequest request);
        BookmarkInfo InsertBookmark(string name, string text);
        BookmarkInfo GetBookmarkText(string name);
        void DeleteBookmark(string name);
        HyperlinkInfo InsertHyperlink(InsertHyperlinkRequest request);

        // ---- 插入与批注 ----
        ImageInfo InsertImage(InsertImageRequest request);
        TocInfo InsertToc(InsertTocRequest request);
        TocInfo UpdateToc();
        HeaderFooterInfo SetHeaderFooter(SetHeaderFooterRequest request);
        HeaderFooterInfo GetHeaderFooter(int sectionIndex);
        CommentInfo InsertComment(string text, string anchorText, int? paragraphIndex);
        List<CommentInfo> ListComments();
        void DeleteComment(int index);

        // ---- 页面设置 ----
        PageSetupInfo GetPageSetup(int sectionIndex);
        PageSetupInfo SetPageSetup(SetPageSetupRequest request);

        // ---- 光标与撤销 ----
        ParagraphInfo InsertTextAtCursor(string text, string style);
        int UndoLast(int count);

        // ---- 段落回读与编辑（P0） ----
        ParagraphDetail GetParagraph(int paragraphIndex);
        DeleteResult DeleteParagraph(int paragraphIndex);
        DeleteResult DeleteText(string text, int? occurrence, bool matchCase, bool matchWholeWord);
        ParagraphInfo AddParagraphRuns(AddParagraphRunsRequest request);
        ParagraphInfo SetListFormat(SetListFormatRequest request);
        ParagraphInfo SetParagraphShading(SetParagraphShadingRequest request);

        // ---- 表格闭环（P0） ----
        TableInfo DeleteTable(int tableIndex);
        TableInfo InsertTableColumns(InsertTableColumnsRequest request);
        TableInfo DeleteTableColumn(int tableIndex, int columnIndex);
        TableInfo SplitTableCell(int tableIndex, int row, int column, int rows, int columns);
        TableInfo SetRowHeight(SetRowHeightRequest request);
        TableInfo SetTableStyle(int tableIndex, string style);

        // ---- 域与文档属性（P0） ----
        FieldUpdateResult UpdateFields();
        DocumentPropertySet GetDocumentProperties();
        DocumentPropertySet SetDocumentProperties(SetDocumentPropertiesRequest request);

        // ---- 样式管理（P0） ----
        StyleInfo CreateStyle(StyleFormatRequest request);
        StyleInfo ModifyStyle(StyleFormatRequest request);
        void DeleteStyle(string name);

        // ---- 水印 / 修订 / 脚注（P0） ----
        WatermarkResult AddWatermark(AddWatermarkRequest request);
        WatermarkResult RemoveWatermark();
        RevisionResult SetTrackChanges(bool enabled);
        RevisionResult AcceptRevisions();
        RevisionResult RejectRevisions();
        NoteInfo InsertFootnote(string text, int? paragraphIndex, string anchorText);

        // ---- 论文增强：注释 / 超链接 / 交叉引用 / 题注 / 域 ----
        NoteInfo InsertEndnote(string text, int? paragraphIndex, string anchorText);
        List<NoteInfo> ListNotes(string noteType);
        NoteInfo GetNote(string noteType, int index);
        void DeleteNote(string noteType, int index);
        NoteOptionsInfo SetNoteOptions(SetNoteOptionsRequest request);
        List<HyperlinkInfo> ListHyperlinks();
        HyperlinkInfo UpdateHyperlink(UpdateHyperlinkRequest request);
        int DeleteHyperlinks(int? index, bool all);
        CrossReferenceResult InsertCrossReference(CrossReferenceRequest request);
        CaptionResult InsertCaption(InsertCaptionRequest request);
        FieldInfoResult InsertField(InsertFieldRequest request);
    }
}
