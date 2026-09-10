namespace WordMcpServer.Mcp
{
    /// <summary>
    /// 业务错误码常量。工具执行失败时随 JSON-RPC 响应一起返回，
    /// 供 AI 客户端区分错误类型并给出可读提示。
    /// </summary>
    public static class McpErrorCodes
    {
        /// <summary>当前没有打开的文档。</summary>
        public const string DocumentNotOpen = "DOC_NOT_OPEN";

        /// <summary>段落索引超出范围。</summary>
        public const string ParagraphOutOfRange = "PARAGRAPH_OUT_OF_RANGE";

        /// <summary>找不到指定的表格。</summary>
        public const string TableNotFound = "TABLE_NOT_FOUND";

        /// <summary>表格行/列索引超出范围。</summary>
        public const string TableCellOutOfRange = "TABLE_CELL_OUT_OF_RANGE";

        /// <summary>找不到指定的书签。</summary>
        public const string BookmarkNotFound = "BOOKMARK_NOT_FOUND";

        /// <summary>找不到指定的批注。</summary>
        public const string CommentNotFound = "COMMENT_NOT_FOUND";

        /// <summary>找不到指定的样式。</summary>
        public const string StyleNotFound = "STYLE_NOT_FOUND";

        /// <summary>找不到指定的文本（查找/替换锚点）。</summary>
        public const string TextNotFound = "TEXT_NOT_FOUND";

        /// <summary>找不到指定的文档。</summary>
        public const string DocumentNotFound = "DOCUMENT_NOT_FOUND";

        /// <summary>参数缺失或取值非法。</summary>
        public const string InvalidParams = "INVALID_PARAMS";

        /// <summary>文件不存在或不可访问。</summary>
        public const string FileNotFound = "FILE_NOT_FOUND";

        /// <summary>不支持的类型（格式、纸张等）。</summary>
        public const string UnsupportedType = "UNSUPPORTED_TYPE";

        /// <summary>Word 对象模型操作失败（COM 异常等）。</summary>
        public const string OperationFailed = "OPERATION_FAILED";

        /// <summary>服务器内部未预期错误。</summary>
        public const string InternalError = "INTERNAL_ERROR";
    }
}
