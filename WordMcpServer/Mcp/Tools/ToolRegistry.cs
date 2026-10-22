using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>
    /// 工具注册表：持有全部 MCP 工具实例，
    /// 供 tools/list（返回定义）与 tools/call（按名执行）使用。
    /// 在插件启动时一次性构造，运行期间不变（线程安全）。
    /// </summary>
    public sealed class ToolRegistry
    {
        /// <summary>只读工具集合（MCP annotations.readOnlyHint = true）。</summary>
        private static readonly HashSet<string> ReadOnlyTools = new HashSet<string>
        {
            "get_active_document_info",
            "list_documents",
            "get_document_text",
            "get_document_outline",
            "get_document_statistics",
            "list_styles",
            "get_table_data",
            "find_text",
            "get_bookmark_text",
            "get_header_footer",
            "list_comments",
            "get_page_setup",
            "get_paragraph",
            "get_document_properties",
            "list_notes",
            "get_note",
            "list_hyperlinks",
        };

        private readonly Dictionary<string, ITool> _tools;
        private readonly JArray _definitions;

        public ToolRegistry(IWordService wordService, IMcpLogger logger)
        {
            var tools = new ITool[]
            {
                // 文档（12）
                new GetActiveDocumentInfoTool(wordService, logger),
                new ListDocumentsTool(wordService, logger),
                new NewDocumentTool(wordService, logger),
                new OpenDocumentTool(wordService, logger),
                new ActivateDocumentTool(wordService, logger),
                new CloseDocumentTool(wordService, logger),
                new SaveDocumentTool(wordService, logger),
                new BackupDocumentTool(wordService, logger),
                new ExportPdfTool(wordService, logger),
                new GetDocumentTextTool(wordService, logger),
                new GetDocumentOutlineTool(wordService, logger),
                new GetDocumentStatisticsTool(wordService, logger),

                // 段落与文本（10 + 6）
                new AddParagraphTool(wordService, logger),
                new AddHeadingTool(wordService, logger),
                new SetParagraphFormatTool(wordService, logger),
                new SetTextStyleTool(wordService, logger),
                new InsertPageBreakTool(wordService, logger),
                new InsertSectionBreakTool(wordService, logger),
                new ApplyStyleTool(wordService, logger),
                new ListStylesTool(wordService, logger),
                new InsertTextAtCursorTool(wordService, logger),
                new UndoLastTool(wordService, logger),
                new GetParagraphTool(wordService, logger),
                new DeleteParagraphTool(wordService, logger),
                new DeleteTextTool(wordService, logger),
                new AddParagraphRunsTool(wordService, logger),
                new SetListFormatTool(wordService, logger),
                new SetParagraphShadingTool(wordService, logger),

                // 表格（8 + 6）
                new InsertTableTool(wordService, logger),
                new GetTableDataTool(wordService, logger),
                new UpdateTableDataTool(wordService, logger),
                new InsertTableRowsTool(wordService, logger),
                new DeleteTableRowTool(wordService, logger),
                new MergeTableCellsTool(wordService, logger),
                new SetTableCellStyleTool(wordService, logger),
                new SetTableLayoutTool(wordService, logger),
                new DeleteTableTool(wordService, logger),
                new InsertTableColumnsTool(wordService, logger),
                new DeleteTableColumnTool(wordService, logger),
                new SplitTableCellTool(wordService, logger),
                new SetRowHeightTool(wordService, logger),
                new SetTableStyleTool(wordService, logger),

                // 查找与链接（6）
                new FindReplaceTextTool(wordService, logger),
                new FindTextTool(wordService, logger),
                new InsertBookmarkTool(wordService, logger),
                new GetBookmarkTextTool(wordService, logger),
                new DeleteBookmarkTool(wordService, logger),
                new InsertHyperlinkTool(wordService, logger),

                // 插入与批注（8 + 3）
                new InsertImageTool(wordService, logger),
                new InsertTocTool(wordService, logger),
                new UpdateTocTool(wordService, logger),
                new SetHeaderFooterTool(wordService, logger),
                new GetHeaderFooterTool(wordService, logger),
                new InsertCommentTool(wordService, logger),
                new ListCommentsTool(wordService, logger),
                new DeleteCommentTool(wordService, logger),
                new AddWatermarkTool(wordService, logger),
                new RemoveWatermarkTool(wordService, logger),
                new InsertFootnoteTool(wordService, logger),

                // 论文增强：注释 / 超链接 / 交叉引用 / 题注 / 域（11）
                new InsertEndnoteTool(wordService, logger),
                new ListNotesTool(wordService, logger),
                new GetNoteTool(wordService, logger),
                new DeleteNoteTool(wordService, logger),
                new SetNoteOptionsTool(wordService, logger),
                new ListHyperlinksTool(wordService, logger),
                new UpdateHyperlinkTool(wordService, logger),
                new DeleteHyperlinkTool(wordService, logger),
                new InsertCrossReferenceTool(wordService, logger),
                new InsertCaptionTool(wordService, logger),
                new InsertFieldTool(wordService, logger),

                // 页面设置（2）
                new GetPageSetupTool(wordService, logger),
                new SetPageSetupTool(wordService, logger),

                // 域与文档属性（6）
                new UpdateFieldsTool(wordService, logger),
                new GetDocumentPropertiesTool(wordService, logger),
                new SetDocumentPropertiesTool(wordService, logger),
                new SetTrackChangesTool(wordService, logger),
                new AcceptRevisionsTool(wordService, logger),
                new RejectRevisionsTool(wordService, logger),

                // 样式管理（3）
                new CreateStyleTool(wordService, logger),
                new ModifyStyleTool(wordService, logger),
                new DeleteStyleTool(wordService, logger),
            };

            _tools = tools.ToDictionary(t => t.Name, t => t);
            _definitions = new JArray();
            foreach (var tool in tools)
            {
                var definition = new JObject
                {
                    ["name"] = tool.Name,
                    ["description"] = tool.Description,
                    ["inputSchema"] = tool.InputSchema,
                };

                // MCP 规范：只读工具标注 readOnlyHint，便于客户端区分读写操作
                if (ReadOnlyTools.Contains(tool.Name))
                {
                    definition["annotations"] = new JObject { ["readOnlyHint"] = true };
                }

                _definitions.Add(definition);
            }
        }

        /// <summary>工具总数。</summary>
        public int Count => _tools.Count;

        /// <summary>按名称查找工具；未找到返回 null。</summary>
        public ITool Find(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return null;
            }

            ITool tool;
            return _tools.TryGetValue(name, out tool) ? tool : null;
        }

        /// <summary>返回 MCP tools/list 所需的工具定义数组（深拷贝，防止外部修改内部缓存）。</summary>
        public JArray ToDefinitions()
        {
            var result = new JArray();
            foreach (var definition in _definitions)
            {
                result.Add(definition.DeepClone());
            }

            return result;
        }
    }
}
