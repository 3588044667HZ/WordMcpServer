using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using WordMcpServer.Logging;
using WordMcpServer.Mcp;
using WordMcpServer.Mcp.Threading;
using WordMcpServer.Mcp.Tools;
using OfficeCore = Microsoft.Office.Core;

namespace WordMcpServer.Services
{
    /// <summary>
    /// Word 对象模型服务实现。
    ///
    /// 线程安全设计：
    /// 1. 所有公开方法先用 _marshaller.Invoke(...) 把 Lambda 封送到 Word 主线程执行；
    ///    Invoke 内部用 SynchronizationContext.Send 阻塞调用方线程直到 UI 线程完成。
    /// 2. COM 对象释放策略：只释放本方法内新建的对象（try/finally），
    ///    对 ActiveDocument 这类由 Word 长期持有的对象不做 FinalReleaseComObject
    ///    （RCW 释放会使其后续访问失败；进程内插件由 GC 兜底回收）。
    /// 3. 业务错误统一抛 McpToolException，由协议层转为 MCP isError 结果。
    ///
    /// PIA 注意事项（已反射/编译实测）：
    /// - Range.Style / Paragraph.Style / Table.Style 是"参数化属性"（setter 为 ref object），
    ///   C# 不能直接赋值，必须用 get_Style() / set_Style(ref object)。
    /// - Font.Bold 等是 Int32（-1/0），不是 bool。
    /// - 行高/表头等属性（KeepWithNext、HeadingFormat、AllowBreakAcrossPages）是 Int32。
    /// </summary>
    public sealed class WordService : IWordService
    {
        /// <summary>水印 WordArt 形状的名称（用于查找并删除水印）。</summary>
        private const string WatermarkShapeName = "MCPWatermark";

        /// <summary>Word 的 wdUndefined 值（混合格式等场景）。</summary>
        private const int WdUndefinedValue = 9999999;

        private readonly Application _app;
        private readonly OfficeThreadMarshaller _marshaller;
        private readonly IMcpLogger _logger;

        public WordService(Application app, OfficeThreadMarshaller marshaller, IMcpLogger logger)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _marshaller = marshaller ?? throw new ArgumentNullException(nameof(marshaller));
            _logger = logger;
        }

        // ============================== 文档 ==============================

        public DocumentInfo GetActiveDocumentInfo()
        {
            return Run("读取文档信息", () => DescribeDocumentInfo(GetActiveDocumentOrThrow()));
        }

        public List<DocumentSummary> ListDocuments()
        {
            return Run("列出文档", () =>
            {
                var result = new List<DocumentSummary>();
                var count = _app.Documents.Count;
                if (count == 0)
                {
                    return result;
                }

                string activeName = SafeRead(() => _app.ActiveDocument?.Name);
                for (var i = 1; i <= count; i++)
                {
                    Document doc = null;
                    try
                    {
                        doc = _app.Documents[i];
                        result.Add(new DocumentSummary
                        {
                            Index = i,
                            Name = SafeRead(() => doc.Name),
                            Path = SafeRead(() => doc.Path),
                            Saved = SafeRead(() => doc.Saved),
                            ReadOnly = SafeRead(() => doc.ReadOnly),
                            Active = string.Equals(SafeRead(() => doc.Name), activeName, StringComparison.Ordinal),
                        });
                    }
                    finally
                    {
                        TryRelease(doc);
                    }
                }

                return result;
            });
        }

        public DocumentInfo NewDocument()
        {
            return Mutate("新建文档", () =>
            {
                var doc = _app.Documents.Add();
                return DescribeDocumentInfo(doc);
            });
        }

        public DocumentInfo ActivateDocument(string name, int? index)
        {
            return Mutate("激活文档", () =>
            {
                var doc = FindDocumentOrThrow(name, index);
                doc.Activate();
                return DescribeDocumentInfo(doc);
            });
        }

        public DocumentInfo OpenDocument(string path, bool readOnly)
        {
            return Mutate("打开文档", () =>
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：path");
                }

                if (!File.Exists(path))
                {
                    throw new McpToolException(McpErrorCodes.FileNotFound, $"文件不存在：{path}");
                }

                object fileName = path;
                object missing = Type.Missing;
                object readOnlyObj = readOnly;
                var doc = _app.Documents.Open(ref fileName, ref missing, ref readOnlyObj, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                return DescribeDocumentInfo(doc);
            });
        }

        public DocumentInfo CloseDocument(string name, int? index, bool saveChanges)
        {
            return Mutate("关闭文档", () =>
            {
                var doc = FindDocumentOrThrow(name, index);
                var info = DescribeDocumentInfo(doc);
                object saveOptions = saveChanges ? WdSaveOptions.wdSaveChanges : WdSaveOptions.wdDoNotSaveChanges;
                object missing = Type.Missing;
                doc.Close(ref saveOptions, ref missing, ref missing);
                return info;
            });
        }

        public DocumentInfo SaveDocument(string saveAsPath, string format)
        {
            return Mutate("保存文档", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (!string.IsNullOrWhiteSpace(saveAsPath))
                {
                    var fullPath = Path.GetFullPath(saveAsPath);
                    var dir = Path.GetDirectoryName(fullPath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    {
                        throw new McpToolException(McpErrorCodes.FileNotFound, $"目录不存在：{dir}");
                    }

                    object fileName = fullPath;
                    object fileFormat = ParseSaveFormat(format, fullPath);
                    object missing = Type.Missing;
                    doc.SaveAs2(ref fileName, ref fileFormat, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                }
                else
                {
                    doc.Save();
                }

                return DescribeDocumentInfo(doc);
            });
        }

        public BackupResult BackupDocument(string backupPath)
        {
            return Mutate("备份文档", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var sourcePath = SafeRead(() => doc.FullName);
                if (string.IsNullOrWhiteSpace(backupPath))
                {
                    var dir = SafeRead(() => doc.Path);
                    if (string.IsNullOrEmpty(dir))
                    {
                        dir = Path.GetTempPath();
                    }

                    var name = Path.GetFileNameWithoutExtension(SafeRead(() => doc.Name));
                    var ext = Path.GetExtension(SafeRead(() => doc.Name));
                    if (string.IsNullOrEmpty(ext))
                    {
                        ext = ".docx";
                    }

                    backupPath = Path.Combine(dir, $"{name}_backup_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");
                }

                backupPath = Path.GetFullPath(backupPath);
                var backupDir = Path.GetDirectoryName(backupPath);
                if (!string.IsNullOrEmpty(backupDir) && !Directory.Exists(backupDir))
                {
                    throw new McpToolException(McpErrorCodes.FileNotFound, $"备份目录不存在：{backupDir}");
                }

                string method;
                if (!string.IsNullOrEmpty(sourcePath) && File.Exists(sourcePath))
                {
                    // 已保存文档：直接文件复制（Word 没有 SaveCopyAs）
                    File.Copy(sourcePath, backupPath, true);
                    method = "file_copy";
                }
                else
                {
                    // 未保存文档：SaveAs2 到备份路径（会改变文档关联，之后可再另存回目标位置）
                    object fileName = backupPath;
                    object missing = Type.Missing;
                    doc.SaveAs2(ref fileName, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    method = "save_as";
                }

                var fi = new FileInfo(backupPath);
                return new BackupResult
                {
                    BackupPath = backupPath,
                    SizeBytes = fi.Exists ? fi.Length : 0,
                    SourcePath = sourcePath,
                    Method = method,
                };
            });
        }

        public ExportResult ExportToPdf(ExportPdfRequest request)
        {
            return Mutate("导出 PDF", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (request == null || string.IsNullOrWhiteSpace(request.OutputPath))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：output_path");
                }

                var outputPath = request.OutputPath;
                if (!outputPath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "output_path 必须以 .pdf 结尾");
                }

                outputPath = Path.GetFullPath(outputPath);
                var dir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    throw new McpToolException(McpErrorCodes.FileNotFound, $"目录不存在：{dir}");
                }

                var range = (request.Range ?? "all").Trim().ToLowerInvariant();
                WdExportRange exportRange;
                var fromPage = 0;
                var toPage = 0;
                switch (range)
                {
                    case "all":
                        exportRange = WdExportRange.wdExportAllDocument;
                        break;
                    case "selection":
                        exportRange = WdExportRange.wdExportSelection;
                        break;
                    case "current-page":
                        exportRange = WdExportRange.wdExportCurrentPage;
                        break;
                    case "from-to":
                        if (!request.FromPage.HasValue || !request.ToPage.HasValue)
                        {
                            throw new McpToolException(McpErrorCodes.InvalidParams,
                                "range=from-to 时必须提供 from_page 和 to_page");
                        }

                        if (request.FromPage.Value < 1 || request.ToPage.Value < request.FromPage.Value)
                        {
                            throw new McpToolException(McpErrorCodes.InvalidParams,
                                "from_page 必须 >= 1，且 to_page >= from_page");
                        }

                        exportRange = WdExportRange.wdExportFromTo;
                        fromPage = request.FromPage.Value;
                        toPage = request.ToPage.Value;
                        break;
                    default:
                        throw new McpToolException(McpErrorCodes.UnsupportedType,
                            $"不支持的导出范围：{request.Range}（支持 all / selection / current-page / from-to）");
                }

                object missing = Type.Missing;
                doc.ExportAsFixedFormat(outputPath, WdExportFormat.wdExportFormatPDF, request.OpenAfterExport,
                    WdExportOptimizeFor.wdExportOptimizeForPrint, exportRange, fromPage, toPage,
                    WdExportItem.wdExportDocumentContent, false, true,
                    request.CreateBookmarks
                        ? WdExportCreateBookmarks.wdExportCreateHeadingBookmarks
                        : WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                    true, false, request.PdfA, ref missing);

                var fi = new FileInfo(outputPath);
                return new ExportResult
                {
                    OutputPath = outputPath,
                    SizeBytes = fi.Exists ? fi.Length : 0,
                    PageCount = SafeRead(() => doc.ComputeStatistics(WdStatistic.wdStatisticPages, ref missing)),
                };
            });
        }

        public DocumentText GetDocumentText(int maxChars)
        {
            return Run("读取文档文本", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var raw = doc.Content.Text ?? string.Empty;
                var text = raw.Replace("\r", "\n").Replace("\v", "\n").TrimEnd('\n', '\a');
                var total = text.Length;
                var truncated = total > maxChars;
                if (truncated)
                {
                    text = text.Substring(0, maxChars);
                }

                return new DocumentText
                {
                    Text = text,
                    CharacterCount = total,
                    ParagraphCount = doc.Paragraphs.Count,
                    Truncated = truncated,
                };
            });
        }

        public DocumentOutline GetDocumentOutline(int maxItems)
        {
            return Run("读取文档大纲", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var outline = new DocumentOutline();
                var count = doc.Paragraphs.Count;
                var scanLimit = Math.Min(count, 20000);
                if (count > scanLimit)
                {
                    outline.Truncated = true;
                }

                for (var i = 1; i <= scanLimit; i++)
                {
                    Paragraph p = null;
                    try
                    {
                        p = doc.Paragraphs[i];
                        var level = GetOutlineLevel(p);
                        if (level < 1 || level > 9)
                        {
                            continue;
                        }

                        if (outline.Items.Count >= maxItems)
                        {
                            outline.Truncated = true;
                            break;
                        }

                        outline.Items.Add(new OutlineItem
                        {
                            ParagraphIndex = i,
                            Level = level,
                            Text = TrimParagraphText(SafeRead(() => p.Range.Text)),
                            Style = SafeRead(() => ((Style)p.Range.get_Style())?.NameLocal),
                        });
                    }
                    finally
                    {
                        TryRelease(p);
                    }
                }

                outline.HeadingCount = outline.Items.Count;
                return outline;
            });
        }

        public DocumentStatistics GetDocumentStatistics()
        {
            return Run("读取文档统计", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                object missing = Type.Missing;
                return new DocumentStatistics
                {
                    Pages = SafeRead(() => doc.ComputeStatistics(WdStatistic.wdStatisticPages, ref missing)),
                    Words = SafeRead(() => doc.ComputeStatistics(WdStatistic.wdStatisticWords, ref missing)),
                    Characters = SafeRead(() => doc.ComputeStatistics(WdStatistic.wdStatisticCharacters, ref missing)),
                    CharactersWithSpaces = SafeRead(() => doc.ComputeStatistics(WdStatistic.wdStatisticCharactersWithSpaces, ref missing)),
                    Paragraphs = SafeRead(() => doc.ComputeStatistics(WdStatistic.wdStatisticParagraphs, ref missing)),
                    Lines = SafeRead(() => doc.ComputeStatistics(WdStatistic.wdStatisticLines, ref missing)),
                };
            });
        }

        // ============================== 段落与文本 ==============================

        public ParagraphInfo AddParagraph(AddParagraphRequest request)
        {
            return Mutate("添加段落", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var index = AddParagraphCore(doc, request.Text, request.Position);
                if (!string.IsNullOrEmpty(request.Style))
                {
                    object styleObj = ResolveStyleValue(doc, request.Style);
                    doc.Paragraphs[index].Range.set_Style(ref styleObj);
                }

                return DescribeParagraph(doc, index);
            });
        }

        public ParagraphInfo AddHeading(string text, int level, int? position)
        {
            return Mutate("添加标题", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (level < 1 || level > 9)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "标题级别必须在 1-9 之间");
                }

                var index = AddParagraphCore(doc, text, position);
                object styleObj = HeadingStyle(level);
                doc.Paragraphs[index].Range.set_Style(ref styleObj);
                return DescribeParagraph(doc, index);
            });
        }

        public ParagraphInfo SetParagraphFormat(SetParagraphFormatRequest request)
        {
            return Mutate("设置段落格式", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (request.ParagraphIndex.HasValue)
                {
                    var p = GetParagraphOrThrow(doc, request.ParagraphIndex.Value);
                    ApplyParagraphFormat(p.Range.ParagraphFormat, request);
                    return DescribeParagraph(doc, request.ParagraphIndex.Value);
                }

                if (!string.IsNullOrEmpty(request.FindText))
                {
                    var index = FindParagraphIndexByText(doc, request.FindText);
                    if (index < 1)
                    {
                        throw new McpToolException(McpErrorCodes.TextNotFound, $"未找到包含文本 \"{request.FindText}\" 的段落");
                    }

                    ApplyParagraphFormat(doc.Paragraphs[index].Range.ParagraphFormat, request);
                    return DescribeParagraph(doc, index);
                }

                throw new McpToolException(McpErrorCodes.InvalidParams, "必须提供 paragraph_index 或 find_text 之一");
            });
        }

        public TextStyleResult SetTextStyle(SetTextStyleRequest request)
        {
            return Mutate("设置文本样式", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (request.ParagraphIndex.HasValue)
                {
                    var p = GetParagraphOrThrow(doc, request.ParagraphIndex.Value);
                    ApplyFont(p.Range, request);
                    return new TextStyleResult
                    {
                        StyledCount = 1,
                        ParagraphIndex = request.ParagraphIndex,
                        Text = TrimParagraphText(SafeRead(() => p.Range.Text)),
                        ParagraphCount = doc.Paragraphs.Count,
                    };
                }

                if (!string.IsNullOrEmpty(request.FindText))
                {
                    var count = ApplyFontToMatches(doc, request.FindText, request);
                    if (count == 0)
                    {
                        throw new McpToolException(McpErrorCodes.TextNotFound, $"未找到文本：{request.FindText}");
                    }

                    return new TextStyleResult
                    {
                        StyledCount = count,
                        Text = request.FindText,
                        ParagraphCount = doc.Paragraphs.Count,
                    };
                }

                throw new McpToolException(McpErrorCodes.InvalidParams, "必须提供 paragraph_index 或 find_text 之一");
            });
        }

        public ParagraphInfo InsertPageBreak(int? beforeParagraph, string text)
        {
            return InsertBreakCore("插入分页符", beforeParagraph, text, WdBreakType.wdPageBreak);
        }

        public ParagraphInfo InsertSectionBreak(int? beforeParagraph, string text)
        {
            return InsertBreakCore("插入分节符", beforeParagraph, text, WdBreakType.wdSectionBreakNextPage);
        }

        public StyleApplyResult ApplyStyle(ApplyStyleRequest request)
        {
            return Mutate("应用样式", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(request.Style))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：style");
                }

                object styleObj = ResolveStyleValue(doc, request.Style);

                if (request.All == true)
                {
                    var content = doc.Content;
                    content.set_Style(ref styleObj);
                    return new StyleApplyResult
                    {
                        Style = request.Style,
                        AppliedCount = doc.Paragraphs.Count,
                        ParagraphCount = doc.Paragraphs.Count,
                    };
                }

                if (request.ParagraphIndex.HasValue)
                {
                    var p = GetParagraphOrThrow(doc, request.ParagraphIndex.Value);
                    var range = p.Range;
                    range.set_Style(ref styleObj);
                    return new StyleApplyResult
                    {
                        Style = request.Style,
                        AppliedCount = 1,
                        ParagraphIndex = request.ParagraphIndex,
                        ParagraphCount = doc.Paragraphs.Count,
                    };
                }

                if (!string.IsNullOrEmpty(request.FindText))
                {
                    var count = ApplyStyleToMatches(doc, request.FindText, styleObj);
                    if (count == 0)
                    {
                        throw new McpToolException(McpErrorCodes.TextNotFound, $"未找到文本：{request.FindText}");
                    }

                    return new StyleApplyResult
                    {
                        Style = request.Style,
                        AppliedCount = count,
                        ParagraphCount = doc.Paragraphs.Count,
                    };
                }

                throw new McpToolException(McpErrorCodes.InvalidParams, "必须提供 paragraph_index、find_text 或 all=true 之一");
            });
        }

        public List<StyleInfo> ListStyles()
        {
            return Run("列出样式", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var result = new List<StyleInfo>();
                var count = doc.Styles.Count;
                for (var i = 1; i <= count && result.Count < 500; i++)
                {
                    Style style = null;
                    try
                    {
                        style = doc.Styles[i];
                        result.Add(new StyleInfo
                        {
                            Name = SafeRead(() => style.NameLocal),
                            Type = SafeRead(() => DescribeStyleType(style.Type)),
                            BuiltIn = SafeRead(() => style.BuiltIn),
                            InUse = SafeRead(() => style.InUse),
                        });
                    }
                    catch
                    {
                        // 单个样式读取失败不影响整体
                    }
                    finally
                    {
                        TryRelease(style);
                    }
                }

                return result;
            });
        }

        // ============================== 表格 ==============================

        public TableInfo InsertTable(InsertTableRequest request)
        {
            return Mutate("插入表格", () => WithScreenUpdating(() =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (request.Rows < 1 || request.Rows > 200)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "rows 必须在 1-200 之间");
                }

                if (request.Columns < 1 || request.Columns > 64)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "columns 必须在 1-64 之间");
                }

                var anchor = GetBlockInsertRange(doc, request.Position);
                var table = doc.Tables.Add(anchor, request.Rows, request.Columns);
                try
                {
                    // 表格插入点位于标题段落时，单元格会继承标题样式（行高/间距异常），统一改回正文样式
                    try
                    {
                        var tableRange = table.Range;
                        object normalStyle = WdBuiltinStyle.wdStyleNormal;
                        tableRange.set_Style(ref normalStyle);
                    }
                    catch
                    {
                        // 模板缺少“正文”样式时保持原样
                    }

                    if (request.Data != null && request.Data.Count > 0)
                    {
                        FillTable(table, request.Data);
                    }

                    if (!string.IsNullOrEmpty(request.Style))
                    {
                        ApplyTableStyle(doc, table, request.Style);
                    }

                    if (request.AutoFit)
                    {
                        table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                    }

                    var index = GetTableIndex(doc, table);
                    return DescribeTable(doc, index);
                }
                finally
                {
                    TryRelease(table);
                }
            }));
        }

        public TableData GetTableData(int tableIndex)
        {
            return Run("读取表格数据", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, tableIndex);
                return ReadTableData(table, tableIndex);
            });
        }

        public TableData UpdateTableData(int tableIndex, List<List<string>> data)
        {
            return Mutate("更新表格数据", () => WithScreenUpdating(() =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, tableIndex);
                if (data == null || data.Count == 0)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：data（二维数组）");
                }

                for (var r = 0; r < data.Count; r++)
                {
                    var row = data[r];
                    if (row == null)
                    {
                        continue;
                    }

                    for (var c = 0; c < row.Count; c++)
                    {
                        SetCellText(table, r + 1, c + 1, row[c]);
                    }
                }

                return ReadTableData(table, tableIndex);
            }));
        }

        public TableInfo InsertTableRows(InsertTableRowsRequest request)
        {
            return Mutate("插入表格行", () => WithScreenUpdating(() =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, request.TableIndex);
                if (request.Count < 1 || request.Count > 100)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "count 必须在 1-100 之间");
                }

                var rowCount = table.Rows.Count;
                int startIndex;
                if (request.BeforeRow.HasValue && request.BeforeRow.Value <= rowCount)
                {
                    var before = request.BeforeRow.Value;
                    if (before < 1)
                    {
                        throw new McpToolException(McpErrorCodes.TableCellOutOfRange, $"before_row {before} 超出范围（当前共 {rowCount} 行）");
                    }

                    object beforeObj = table.Rows[before];
                    for (var i = 0; i < request.Count; i++)
                    {
                        table.Rows.Add(ref beforeObj);
                    }

                    startIndex = before;
                }
                else
                {
                    if (request.BeforeRow.HasValue && request.BeforeRow.Value > rowCount + 1)
                    {
                        throw new McpToolException(McpErrorCodes.TableCellOutOfRange,
                            $"before_row {request.BeforeRow.Value} 超出范围（当前共 {rowCount} 行，允许 1-{rowCount + 1}）");
                    }

                    object missing = Type.Missing;
                    for (var i = 0; i < request.Count; i++)
                    {
                        table.Rows.Add(ref missing);
                    }

                    startIndex = rowCount + 1;
                }

                if (request.Data != null)
                {
                    for (var r = 0; r < request.Data.Count && r < request.Count; r++)
                    {
                        var row = request.Data[r];
                        if (row == null)
                        {
                            continue;
                        }

                        for (var c = 0; c < row.Count; c++)
                        {
                            SetCellText(table, startIndex + r, c + 1, row[c]);
                        }
                    }
                }

                var index = GetTableIndex(doc, table);
                return DescribeTable(doc, index);
            }));
        }

        public TableInfo DeleteTableRow(int tableIndex, int rowIndex)
        {
            return Mutate("删除表格行", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, tableIndex);
                var rowCount = table.Rows.Count;
                if (rowCount <= 1)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "表格只剩一行，无法删除（可删除整个表格内容）");
                }

                if (rowIndex < 1 || rowIndex > rowCount)
                {
                    throw new McpToolException(McpErrorCodes.TableCellOutOfRange,
                        $"行索引 {rowIndex} 超出范围（当前共 {rowCount} 行）");
                }

                table.Rows[rowIndex].Delete();
                return DescribeTable(doc, tableIndex);
            });
        }

        public TableInfo MergeTableCells(int tableIndex, int startRow, int startColumn, int endRow, int endColumn)
        {
            return Mutate("合并单元格", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, tableIndex);
                ValidateCellRange(table, startRow, startColumn, "start_row", "start_column");
                ValidateCellRange(table, endRow, endColumn, "end_row", "end_column");
                if (endRow < startRow || endColumn < startColumn)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "结束行列不能小于起始行列");
                }

                table.Cell(startRow, startColumn).Merge(table.Cell(endRow, endColumn));
                return DescribeTable(doc, tableIndex);
            });
        }

        public TableInfo SetTableCellStyle(SetTableCellStyleRequest request)
        {
            return Mutate("设置单元格样式", () => WithScreenUpdating(() =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, request.TableIndex);
                var rowStart = request.Row ?? 1;
                var rowEnd = request.Row ?? table.Rows.Count;
                var colStart = request.Column ?? 1;
                var colEnd = request.Column ?? table.Columns.Count;
                ValidateCellRange(table, rowStart, colStart, "row", "column");
                ValidateCellRange(table, rowEnd, colEnd, "row", "column");

                for (var r = rowStart; r <= rowEnd; r++)
                {
                    for (var c = colStart; c <= colEnd; c++)
                    {
                        Cell cell = null;
                        try
                        {
                            cell = table.Cell(r, c);
                            ApplyCellStyle(cell, request);
                        }
                        catch (McpToolException)
                        {
                            throw;
                        }
                        catch
                        {
                            // 合并单元格的覆盖位置访问会失败，跳过
                        }
                        finally
                        {
                            TryRelease(cell);
                        }
                    }
                }

                return DescribeTable(doc, request.TableIndex);
            }));
        }

        public TableInfo SetTableLayout(SetTableLayoutRequest request)
        {
            return Mutate("设置表格布局", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, request.TableIndex);

                if (!string.IsNullOrEmpty(request.Alignment))
                {
                    table.Rows.Alignment = ParseRowAlignment(request.Alignment);
                }

                if (!string.IsNullOrEmpty(request.AutoFit))
                {
                    table.AutoFitBehavior(ParseAutoFit(request.AutoFit));
                }

                if (request.ColumnWidths != null && request.ColumnWidths.Count > 0)
                {
                    for (var c = 0; c < request.ColumnWidths.Count && c < table.Columns.Count; c++)
                    {
                        var width = request.ColumnWidths[c];
                        if (width <= 0)
                        {
                            throw new McpToolException(McpErrorCodes.InvalidParams, $"column_widths[{c}] 必须大于 0");
                        }

                        var column = table.Columns[c + 1];
                        try
                        {
                            column.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
                            column.Width = (float)width;
                        }
                        finally
                        {
                            TryRelease(column);
                        }
                    }
                }

                if (request.RepeatHeaderRow.HasValue)
                {
                    table.Rows[1].HeadingFormat = request.RepeatHeaderRow.Value ? -1 : 0;
                }

                if (request.AllowBreakAcrossPages.HasValue)
                {
                    table.Rows.AllowBreakAcrossPages = request.AllowBreakAcrossPages.Value ? -1 : 0;
                }

                return DescribeTable(doc, request.TableIndex);
            });
        }

        // ============================== 查找与链接 ==============================

        public FindReplaceResult FindReplaceText(FindReplaceRequest request)
        {
            return Mutate("查找替换", () => WithScreenUpdating(() =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrEmpty(request.FindText))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：find_text");
                }

                if (request.Occurrence.HasValue && request.Occurrence.Value < 1)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "occurrence 必须 >= 1");
                }

                var ranges = ResolveFindRanges(doc, request.Scope);
                var occurrenceCounter = 0;
                var count = 0;
                foreach (var range in ranges)
                {
                    count += ReplaceInRange(range, request, ref occurrenceCounter);
                    if (request.Occurrence.HasValue && occurrenceCounter >= request.Occurrence.Value)
                    {
                        break;
                    }
                }

                return new FindReplaceResult
                {
                    FindText = request.FindText,
                    ReplaceText = request.ReplaceText,
                    ReplacementCount = count,
                    Message = count == 0 ? "未找到匹配文本" : null,
                };
            }));
        }

        public FindResult FindText(FindTextRequest request)
        {
            return Run("查找文本", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrEmpty(request.Text))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：text");
                }

                var maxResults = request.MaxResults < 1 ? 20 : Math.Min(request.MaxResults, 200);
                var result = new FindResult { FindText = request.Text };
                var scope = (request.Scope ?? "document").Trim().ToLowerInvariant();
                var ranges = ResolveFindRanges(doc, scope);
                var trackParagraph = scope != "headers-footers";

                foreach (var range in ranges)
                {
                    var find = range.Find;
                    find.ClearFormatting();
                    find.Text = request.Text;
                    find.MatchCase = request.MatchCase;
                    find.MatchWholeWord = request.MatchWholeWord;
                    find.MatchWildcards = request.MatchWildcards;
                    find.Forward = true;
                    find.Wrap = WdFindWrap.wdFindStop;

                    object missing = Type.Missing;
                    object findTextObj = request.Text;
                    while (result.Matches.Count < maxResults &&
                           find.Execute(ref findTextObj, ref missing, ref missing, ref missing, ref missing,
                               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                               ref missing, ref missing, ref missing, ref missing))
                    {
                        result.Matches.Add(new FindMatch
                        {
                            ParagraphIndex = trackParagraph ? GetParagraphIndexAt(doc, range.Start) : 0,
                            Text = SafeRead(() => range.Text),
                            Start = range.Start,
                            End = range.End,
                        });
                        range.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }

                    if (result.Matches.Count >= maxResults)
                    {
                        break;
                    }
                }

                result.MatchCount = result.Matches.Count;
                return result;
            });
        }

        public BookmarkInfo InsertBookmark(string name, string text, string findText, int? paragraphIndex, int? occurrence)
        {
            return Mutate("插入书签", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(name))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：name");
                }

                if (doc.Bookmarks.Exists(name))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, $"书签已存在：{name}");
                }

                Range range;
                if (!string.IsNullOrEmpty(findText))
                {
                    range = FindNthRange(doc, findText, occurrence ?? 1);
                    if (range == null)
                    {
                        throw new McpToolException(McpErrorCodes.TextNotFound,
                            $"未找到要加书签的文本：{findText}");
                    }
                }
                else if (paragraphIndex.HasValue)
                {
                    range = GetParagraphOrThrow(doc, paragraphIndex.Value).Range;
                    if (range.End > range.Start)
                    {
                        range.End = range.End - 1;
                    }
                }
                else
                {
                    var index = AddParagraphCore(doc, text, null);
                    range = doc.Paragraphs[index].Range;
                    if (range.End > range.Start)
                    {
                        range.End = range.End - 1;
                    }
                }

                object rangeObj = range;
                doc.Bookmarks.Add(name, ref rangeObj);
                return new BookmarkInfo
                {
                    Name = name,
                    Exists = true,
                    Text = TrimParagraphText(SafeRead(() => range.Text)),
                    Start = SafeRead(() => range.Start),
                    End = SafeRead(() => range.End),
                };
            });
        }

        public BookmarkInfo GetBookmarkText(string name)
        {
            return Run("读取书签", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(name) || !doc.Bookmarks.Exists(name))
                {
                    throw new McpToolException(McpErrorCodes.BookmarkNotFound, $"找不到书签：{name}");
                }

                Bookmark bookmark = null;
                try
                {
                    bookmark = doc.Bookmarks[name];
                    return new BookmarkInfo
                    {
                        Name = name,
                        Exists = true,
                        Text = TrimParagraphText(SafeRead(() => bookmark.Range.Text)),
                        Start = SafeRead(() => bookmark.Start),
                        End = SafeRead(() => bookmark.End),
                    };
                }
                finally
                {
                    TryRelease(bookmark);
                }
            });
        }

        public void DeleteBookmark(string name)
        {
            Mutate("删除书签", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(name) || !doc.Bookmarks.Exists(name))
                {
                    throw new McpToolException(McpErrorCodes.BookmarkNotFound, $"找不到书签：{name}");
                }

                doc.Bookmarks[name].Delete();
                return 0;
            });
        }

        public HyperlinkInfo InsertHyperlink(InsertHyperlinkRequest request)
        {
            return Mutate("插入超链接", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(request.Address) && string.IsNullOrWhiteSpace(request.Bookmark))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams,
                        "必须提供 address（外部链接）或 bookmark（内部锚点）之一");
                }

                if (!string.IsNullOrWhiteSpace(request.Bookmark) && !doc.Bookmarks.Exists(request.Bookmark))
                {
                    throw new McpToolException(McpErrorCodes.BookmarkNotFound,
                        $"找不到书签：{request.Bookmark}（请先用 insert_bookmark 创建）");
                }

                Range anchor;
                int index;
                if (request.ParagraphIndex.HasValue)
                {
                    var p = GetParagraphOrThrow(doc, request.ParagraphIndex.Value);
                    index = request.ParagraphIndex.Value;
                    anchor = p.Range;
                    anchor.Collapse(WdCollapseDirection.wdCollapseStart);
                }
                else
                {
                    index = AddParagraphCore(doc, null, null);
                    anchor = doc.Paragraphs[index].Range;
                    anchor.Collapse(WdCollapseDirection.wdCollapseStart);
                }

                object address = request.Address ?? string.Empty;
                object subAddress = request.Bookmark ?? string.Empty;
                object screenTip = request.ScreenTip ?? string.Empty;
                object textToDisplay = request.Text ?? request.Address ?? request.Bookmark;
                object target = string.Empty;
                var link = doc.Hyperlinks.Add(anchor, ref address, ref subAddress, ref screenTip, ref textToDisplay, ref target);

                return new HyperlinkInfo
                {
                    Index = FindHyperlinkIndex(doc, link),
                    Text = request.Text ?? request.Address ?? request.Bookmark,
                    Address = request.Address,
                    SubAddress = request.Bookmark,
                    ScreenTip = request.ScreenTip,
                    ParagraphIndex = index,
                    HyperlinkCount = doc.Hyperlinks.Count,
                };
            });
        }

        // ============================== 插入与批注 ==============================

        public ImageEntry InsertImage(InsertImageRequest request)
        {
            return Mutate("插入图片", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(request.Path))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：path");
                }

                var path = Path.GetFullPath(request.Path);
                if (!File.Exists(path))
                {
                    throw new McpToolException(McpErrorCodes.FileNotFound, $"图片文件不存在：{path}");
                }

                var layout = (request.Layout ?? "inline").Trim().ToLowerInvariant();
                if (layout != "inline" && layout != "floating")
                {
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的图片布局：{request.Layout}（支持 inline / floating）");
                }

                Range anchor;
                int index;
                if (request.ParagraphIndex.HasValue)
                {
                    index = request.ParagraphIndex.Value;
                    anchor = GetParagraphOrThrow(doc, index).Range;
                }
                else
                {
                    index = AddParagraphCore(doc, null, null);
                    anchor = doc.Paragraphs[index].Range;
                }

                if (layout == "inline")
                {
                    var inlineAnchor = anchor;
                    inlineAnchor.Collapse(WdCollapseDirection.wdCollapseStart);
                    object linkToFile = false;
                    object saveWithDocument = true;
                    object rangeObj = inlineAnchor;
                    var shape = doc.InlineShapes.AddPicture(path, ref linkToFile, ref saveWithDocument, ref rangeObj);
                    try
                    {
                        if (request.Width.HasValue && request.Height.HasValue)
                        {
                            shape.Width = (float)request.Width.Value;
                            shape.Height = (float)request.Height.Value;
                        }
                        else if (request.Width.HasValue)
                        {
                            var ratio = request.Width.Value / shape.Width;
                            shape.Width = (float)request.Width.Value;
                            shape.Height = (float)(shape.Height * ratio);
                        }
                        else if (request.Height.HasValue)
                        {
                            var ratio = request.Height.Value / shape.Height;
                            shape.Height = (float)request.Height.Value;
                            shape.Width = (float)(shape.Width * ratio);
                        }

                        if (!string.IsNullOrEmpty(request.Alignment))
                        {
                            shape.Range.ParagraphFormat.Alignment = ParseAlignment(request.Alignment);
                        }

                        if (!string.IsNullOrEmpty(request.AltText))
                        {
                            shape.AlternativeText = request.AltText;
                        }

                        return DescribeInlineImage(doc, FindInlineShapeIndex(doc, shape));
                    }
                    finally
                    {
                        TryRelease(shape);
                    }
                }

                // 浮动图片
                object floatingLink = false;
                object floatingSave = true;
                object left = request.Left ?? 0;
                object top = request.Top ?? 0;
                object width = request.Width.HasValue ? (object)(float)request.Width.Value : Type.Missing;
                object height = request.Height.HasValue ? (object)(float)request.Height.Value : Type.Missing;
                object anchorObj = anchor;
                var floating = doc.Shapes.AddPicture(path, ref floatingLink, ref floatingSave, ref left, ref top,
                    ref width, ref height, ref anchorObj);
                try
                {
                    if (!string.IsNullOrEmpty(request.Wrap))
                    {
                        floating.WrapFormat.Type = ParseWrapType(request.Wrap);
                    }

                    floating.WrapFormat.AllowOverlap = -1;
                    var relative = (request.RelativeTo ?? "page").Trim().ToLowerInvariant();
                    floating.RelativeHorizontalPosition = ParseRelativeHorizontal(relative);
                    floating.RelativeVerticalPosition = ParseRelativeVertical(relative);

                    if (request.Width.HasValue)
                    {
                        floating.Width = (float)request.Width.Value;
                    }

                    if (request.Height.HasValue)
                    {
                        floating.Height = (float)request.Height.Value;
                    }

                    if (!string.IsNullOrEmpty(request.AltText))
                    {
                        floating.AlternativeText = request.AltText;
                    }

                    return DescribeFloatingImage(doc, floating);
                }
                finally
                {
                    TryRelease(floating);
                }
            });
        }

        public TocInfo InsertToc(InsertTocRequest request)
        {
            return Mutate("插入目录", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var levels = Math.Max(1, Math.Min(9, request.Levels));

                if (!string.IsNullOrEmpty(request.Title))
                {
                    AddParagraphCore(doc, request.Title, null);
                }

                var index = AddParagraphCore(doc, null, null);
                var anchor = doc.Paragraphs[index].Range;
                anchor.Collapse(WdCollapseDirection.wdCollapseStart);

                object useHeadingStyles = true;
                object upperHeadingLevel = 1;
                object lowerHeadingLevel = levels;
                object useFields = false;
                object tableId = string.Empty;
                object rightAlignPageNumbers = true;
                object includePageNumbers = true;
                object addedStyles = string.Empty;
                object useHyperlinks = true;
                object hidePageNumbersInWeb = false;
                object useOutlineLevels = true;
                doc.TablesOfContents.Add(anchor, ref useHeadingStyles, ref upperHeadingLevel, ref lowerHeadingLevel,
                    ref useFields, ref tableId, ref rightAlignPageNumbers, ref includePageNumbers, ref addedStyles,
                    ref useHyperlinks, ref hidePageNumbersInWeb, ref useOutlineLevels);

                return new TocInfo
                {
                    TocCount = doc.TablesOfContents.Count,
                    UpperLevel = 1,
                    LowerLevel = levels,
                    Updated = true,
                };
            });
        }

        public TocInfo UpdateToc()
        {
            return Mutate("更新目录", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var count = doc.TablesOfContents.Count;
                for (var i = 1; i <= count; i++)
                {
                    TableOfContents toc = null;
                    try
                    {
                        toc = doc.TablesOfContents[i];
                        toc.Update();
                    }
                    finally
                    {
                        TryRelease(toc);
                    }
                }

                return new TocInfo
                {
                    TocCount = count,
                    Updated = count > 0,
                    UpperLevel = 0,
                    LowerLevel = 0,
                };
            });
        }

        public HeaderFooterInfo SetHeaderFooter(SetHeaderFooterRequest request)
        {
            return Mutate("设置页眉页脚", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var section = GetSectionOrThrow(doc, request.SectionIndex);

                if (request.DifferentFirstPage.HasValue)
                {
                    section.PageSetup.DifferentFirstPageHeaderFooter = request.DifferentFirstPage.Value ? -1 : 0;
                }

                if (request.OddEvenPages.HasValue)
                {
                    section.PageSetup.OddAndEvenPagesHeaderFooter = request.OddEvenPages.Value ? -1 : 0;
                }

                var headerIndex = ParseHeaderFooterType(request.HeaderType);
                var header = section.Headers[headerIndex];
                var footer = section.Footers[headerIndex];

                if (request.LinkToPrevious.HasValue)
                {
                    try
                    {
                        header.LinkToPrevious = request.LinkToPrevious.Value;
                        footer.LinkToPrevious = request.LinkToPrevious.Value;
                    }
                    catch
                    {
                        // 首页/奇偶页未启用时设置链接可能失败，忽略
                    }
                }

                if (request.HeaderText != null)
                {
                    header.Range.Text = NormalizeText(request.HeaderText) ?? string.Empty;
                }

                if (request.FooterText != null)
                {
                    footer.Range.Text = NormalizeText(request.FooterText) ?? string.Empty;
                }

                if (request.IncludePageNumber)
                {
                    var range = footer.Range;
                    if (TrimParagraphText(SafeRead(() => range.Text)).Length > 0)
                    {
                        range.Collapse(WdCollapseDirection.wdCollapseEnd);
                        range.InsertAfter("  ");
                    }
                    else
                    {
                        range.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }

                    object fieldType = WdFieldType.wdFieldPage;
                    object fieldText = string.Empty;
                    object preserveFormatting = false;
                    doc.Fields.Add(range, ref fieldType, ref fieldText, ref preserveFormatting);
                }

                return DescribeHeaderFooter(doc, request.SectionIndex, request.HeaderType);
            });
        }

        public HeaderFooterInfo GetHeaderFooter(int sectionIndex, string headerType)
        {
            return Run("读取页眉页脚", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                GetSectionOrThrow(doc, sectionIndex);
                return DescribeHeaderFooter(doc, sectionIndex, headerType);
            });
        }

        public CommentInfo InsertComment(string text, string anchorText, int? paragraphIndex)
        {
            return Mutate("插入批注", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrEmpty(text))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：text");
                }

                Range scope;
                if (paragraphIndex.HasValue)
                {
                    scope = GetParagraphOrThrow(doc, paragraphIndex.Value).Range;
                }
                else
                {
                    if (string.IsNullOrEmpty(anchorText))
                    {
                        throw new McpToolException(McpErrorCodes.InvalidParams, "必须提供 anchor_text 或 paragraph_index 之一");
                    }

                    scope = FindFirstRange(doc, anchorText);
                    if (scope == null)
                    {
                        throw new McpToolException(McpErrorCodes.TextNotFound, $"未找到批注锚点文本：{anchorText}");
                    }
                }

                object commentText = text;
                var comment = doc.Comments.Add(scope, ref commentText);
                var index = FindCommentIndex(doc, comment);
                return DescribeComment(doc, index);
            });
        }

        public List<CommentInfo> ListComments()
        {
            return Run("列出批注", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var result = new List<CommentInfo>();
                var count = doc.Comments.Count;
                for (var i = 1; i <= count && result.Count < 500; i++)
                {
                    result.Add(DescribeComment(doc, i));
                }

                return result;
            });
        }

        public void DeleteComment(int index)
        {
            Mutate("删除批注", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var count = doc.Comments.Count;
                if (index < 1 || index > count)
                {
                    throw new McpToolException(McpErrorCodes.CommentNotFound,
                        $"批注索引 {index} 超出范围（当前共 {count} 条批注）");
                }

                doc.Comments[index].Delete();
                return 0;
            });
        }

        // ============================== 页面设置 ==============================

        public PageSetupInfo GetPageSetup(int sectionIndex)
        {
            return Run("读取页面设置", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                GetSectionOrThrow(doc, sectionIndex);
                return DescribePageSetup(doc, sectionIndex);
            });
        }

        public PageSetupInfo SetPageSetup(SetPageSetupRequest request)
        {
            return Mutate("设置页面设置", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var section = GetSectionOrThrow(doc, request.SectionIndex);
                var setup = section.PageSetup;

                if (!string.IsNullOrEmpty(request.Orientation))
                {
                    setup.Orientation = ParseOrientation(request.Orientation);
                }

                if (!string.IsNullOrEmpty(request.PaperSize))
                {
                    setup.PaperSize = ParsePaperSize(request.PaperSize);
                }

                if (request.TopMargin.HasValue) setup.TopMargin = (float)request.TopMargin.Value;
                if (request.BottomMargin.HasValue) setup.BottomMargin = (float)request.BottomMargin.Value;
                if (request.LeftMargin.HasValue) setup.LeftMargin = (float)request.LeftMargin.Value;
                if (request.RightMargin.HasValue) setup.RightMargin = (float)request.RightMargin.Value;
                if (request.HeaderDistance.HasValue) setup.HeaderDistance = (float)request.HeaderDistance.Value;
                if (request.FooterDistance.HasValue) setup.FooterDistance = (float)request.FooterDistance.Value;
                if (request.Gutter.HasValue) setup.Gutter = (float)request.Gutter.Value;

                if (request.TextColumns.HasValue)
                {
                    if (request.TextColumns.Value < 1 || request.TextColumns.Value > 8)
                    {
                        throw new McpToolException(McpErrorCodes.InvalidParams, "text_columns 必须在 1-8 之间");
                    }

                    setup.TextColumns.SetCount(request.TextColumns.Value);
                }

                return DescribePageSetup(doc, request.SectionIndex);
            });
        }

        // ============================== 光标与撤销 ==============================

        public ParagraphInfo InsertTextAtCursor(string text, string style)
        {
            return Mutate("光标处插入文本", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var selection = _app.Selection;
                selection.TypeText(NormalizeText(text) ?? string.Empty);

                if (!string.IsNullOrEmpty(style))
                {
                    object styleObj = ResolveStyleValue(doc, style);
                    var range = selection.Range;
                    range.set_Style(ref styleObj);
                }

                var index = GetParagraphIndexAt(doc, selection.End);
                if (index < 1 || index > doc.Paragraphs.Count)
                {
                    index = doc.Paragraphs.Count;
                }

                return DescribeParagraph(doc, index);
            });
        }

        public int UndoLast(int count)
        {
            return Mutate("撤销操作", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (count < 1 || count > 100)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "count 必须在 1-100 之间");
                }

                object times = count;
                return doc.Undo(ref times) ? count : 0;
            });
        }

        // ============================== P0 扩展：段落回读与编辑 ==============================

        public ParagraphDetail GetParagraph(int paragraphIndex)
        {
            return Run("读取段落详情", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var p = GetParagraphOrThrow(doc, paragraphIndex);
                return DescribeParagraphDetail(doc, p, paragraphIndex);
            });
        }

        public DeleteResult DeleteParagraph(int paragraphIndex)
        {
            return Mutate("删除段落", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var p = GetParagraphOrThrow(doc, paragraphIndex);
                var inTable = SafeRead(() => (bool)p.Range.Information[WdInformation.wdWithInTable]);
                p.Range.Delete();
                return new DeleteResult
                {
                    Deleted = true,
                    DeletedCount = 1,
                    ParagraphIndex = paragraphIndex,
                    ParagraphCount = doc.Paragraphs.Count,
                    Message = inTable ? "已删除表格单元格内的段落内容" : null,
                };
            });
        }

        public DeleteResult DeleteText(string text, int? occurrence, bool matchCase, bool matchWholeWord)
        {
            return Mutate("删除文本", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrEmpty(text))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：text");
                }

                var limit = occurrence ?? 10000;
                if (limit < 1)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "occurrence 必须 >= 1");
                }

                var range = doc.Content;
                var find = range.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = text;
                find.Replacement.Text = string.Empty;
                find.MatchCase = matchCase;
                find.MatchWholeWord = matchWholeWord;
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;

                object missing = Type.Missing;
                object findTextObj = text;
                object replaceOne = WdReplace.wdReplaceOne;
                var count = 0;
                while (count < limit && find.Execute(ref findTextObj, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceOne,
                    ref missing, ref missing, ref missing, ref missing))
                {
                    count++;
                    range.Collapse(WdCollapseDirection.wdCollapseEnd);
                }

                return new DeleteResult
                {
                    Deleted = count > 0,
                    DeletedCount = count,
                    ParagraphCount = doc.Paragraphs.Count,
                    Message = count == 0 ? "未找到匹配文本" : null,
                };
            });
        }

        public ParagraphInfo AddParagraphRuns(AddParagraphRunsRequest request)
        {
            return Mutate("添加富文本段落", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (request.Runs == null || request.Runs.Count == 0)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：runs");
                }

                var index = AddParagraphCore(doc, null, request.Position);
                var paragraph = doc.Paragraphs[index];
                if (!string.IsNullOrEmpty(request.Style))
                {
                    object styleObj = ResolveStyleValue(doc, request.Style);
                    paragraph.Range.set_Style(ref styleObj);
                }

                var range = paragraph.Range;
                range.Collapse(WdCollapseDirection.wdCollapseStart);
                foreach (var run in request.Runs)
                {
                    if (run == null)
                    {
                        continue;
                    }

                    var text = NormalizeText(run.Text) ?? string.Empty;
                    if (text.Length == 0)
                    {
                        continue;
                    }

                    range.InsertAfter(text);
                    ApplyRunFont(range, run);
                    range.Collapse(WdCollapseDirection.wdCollapseEnd);
                }

                return DescribeParagraph(doc, index);
            });
        }

        public ParagraphInfo SetListFormat(SetListFormatRequest request)
        {
            return Mutate("设置列表格式", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var range = ResolveListRange(doc, request);
                var listFormat = range.ListFormat;
                object missing = Type.Missing;
                switch ((request.Type ?? string.Empty).Trim().ToLowerInvariant())
                {
                    case "bullet":
                        listFormat.ApplyBulletDefault(ref missing);
                        break;
                    case "number":
                        listFormat.ApplyNumberDefault(ref missing);
                        break;
                    case "outline":
                        var template = _app.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
                        object continuePrevious = true;
                        object applyTo = missing;
                        object behavior = WdDefaultListBehavior.wdWord10ListBehavior;
                        listFormat.ApplyListTemplate(template, ref continuePrevious, ref applyTo, ref behavior);
                        break;
                    case "none":
                        listFormat.RemoveNumbers(ref missing);
                        break;
                    default:
                        throw new McpToolException(McpErrorCodes.UnsupportedType,
                            $"不支持的列表类型：{request.Type}（支持 bullet / number / outline / none）");
                }

                if (request.Level.HasValue)
                {
                    if (request.Level.Value < 1 || request.Level.Value > 9)
                    {
                        throw new McpToolException(McpErrorCodes.InvalidParams, "level 必须在 1-9 之间");
                    }

                    listFormat.ListLevelNumber = request.Level.Value;
                }

                var index = GetParagraphIndexAt(doc, range.Start);
                if (index < 1 || index > doc.Paragraphs.Count)
                {
                    index = doc.Paragraphs.Count;
                }

                return DescribeParagraph(doc, index);
            });
        }

        public ParagraphInfo SetParagraphShading(SetParagraphShadingRequest request)
        {
            return Mutate("设置段落底纹/边框", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (!request.ShadingArgb.HasValue && !request.BorderArgb.HasValue &&
                    string.IsNullOrEmpty(request.BorderStyle))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams,
                        "至少提供 shading、border_color 或 border_style 之一");
                }

                int index;
                Paragraph p;
                if (request.ParagraphIndex.HasValue)
                {
                    index = request.ParagraphIndex.Value;
                    p = GetParagraphOrThrow(doc, index);
                }
                else if (!string.IsNullOrEmpty(request.FindText))
                {
                    index = FindParagraphIndexByText(doc, request.FindText);
                    if (index < 1)
                    {
                        throw new McpToolException(McpErrorCodes.TextNotFound,
                            $"未找到包含文本 \"{request.FindText}\" 的段落");
                    }

                    p = doc.Paragraphs[index];
                }
                else
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "必须提供 paragraph_index 或 find_text 之一");
                }

                var range = p.Range;
                if (request.ShadingArgb.HasValue)
                {
                    range.Shading.BackgroundPatternColor = ArgbToWdColor(request.ShadingArgb.Value);
                }

                var borderStyle = request.BorderStyle ?? (request.BorderArgb.HasValue ? "single" : null);
                if (!string.IsNullOrEmpty(borderStyle))
                {
                    var lineStyle = ParseBorderStyle(borderStyle);
                    foreach (var borderType in new[]
                             {
                                 WdBorderType.wdBorderTop, WdBorderType.wdBorderLeft,
                                 WdBorderType.wdBorderBottom, WdBorderType.wdBorderRight,
                             })
                    {
                        Border border = null;
                        try
                        {
                            border = range.Borders[borderType];
                            border.LineStyle = lineStyle;
                            if (request.BorderArgb.HasValue)
                            {
                                border.Color = ArgbToWdColor(request.BorderArgb.Value);
                            }

                            if (!string.IsNullOrEmpty(request.BorderWidth))
                            {
                                border.LineWidth = ParseBorderWidth(request.BorderWidth);
                            }
                        }
                        catch
                        {
                            // 单条边框设置失败不影响其他边框
                        }
                        finally
                        {
                            TryRelease(border);
                        }
                    }
                }

                return DescribeParagraph(doc, index);
            });
        }

        // ============================== P0 扩展：表格闭环 ==============================

        public TableInfo DeleteTable(int tableIndex)
        {
            return Mutate("删除表格", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, tableIndex);
                var info = DescribeTable(doc, tableIndex);
                table.Delete();
                info.TableIndex = tableIndex;
                info.TableCount = doc.Tables.Count;
                return info;
            });
        }

        public TableInfo InsertTableColumns(InsertTableColumnsRequest request)
        {
            return Mutate("插入表格列", () => WithScreenUpdating(() =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, request.TableIndex);
                if (request.Count < 1 || request.Count > 20)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "count 必须在 1-20 之间");
                }

                var columnCount = table.Columns.Count;
                int startColumn;
                if (request.BeforeColumn.HasValue && request.BeforeColumn.Value <= columnCount)
                {
                    var before = request.BeforeColumn.Value;
                    if (before < 1)
                    {
                        throw new McpToolException(McpErrorCodes.TableCellOutOfRange,
                            $"before_column {before} 超出范围（当前共 {columnCount} 列）");
                    }

                    object beforeObj = table.Columns[before];
                    for (var i = 0; i < request.Count; i++)
                    {
                        table.Columns.Add(ref beforeObj);
                    }

                    startColumn = before;
                }
                else
                {
                    if (request.BeforeColumn.HasValue && request.BeforeColumn.Value > columnCount + 1)
                    {
                        throw new McpToolException(McpErrorCodes.TableCellOutOfRange,
                            $"before_column {request.BeforeColumn.Value} 超出范围（当前共 {columnCount} 列，允许 1-{columnCount + 1}）");
                    }

                    object missing = Type.Missing;
                    for (var i = 0; i < request.Count; i++)
                    {
                        table.Columns.Add(ref missing);
                    }

                    startColumn = columnCount + 1;
                }

                if (request.Data != null)
                {
                    for (var r = 0; r < request.Data.Count && r < table.Rows.Count; r++)
                    {
                        var row = request.Data[r];
                        if (row == null)
                        {
                            continue;
                        }

                        for (var c = 0; c < row.Count && c < request.Count; c++)
                        {
                            SetCellText(table, r + 1, startColumn + c, row[c]);
                        }
                    }
                }

                return DescribeTable(doc, request.TableIndex);
            }));
        }

        public TableInfo DeleteTableColumn(int tableIndex, int columnIndex)
        {
            return Mutate("删除表格列", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, tableIndex);
                var columnCount = table.Columns.Count;
                if (columnCount <= 1)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "表格只剩一列，无法删除");
                }

                if (columnIndex < 1 || columnIndex > columnCount)
                {
                    throw new McpToolException(McpErrorCodes.TableCellOutOfRange,
                        $"列索引 {columnIndex} 超出范围（当前共 {columnCount} 列）");
                }

                table.Columns[columnIndex].Delete();
                return DescribeTable(doc, tableIndex);
            });
        }

        public TableInfo SplitTableCell(int tableIndex, int row, int column, int rows, int columns)
        {
            return Mutate("拆分单元格", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, tableIndex);
                ValidateCellRange(table, row, column, "row", "column");
                if (rows < 1 || rows > 50 || columns < 1 || columns > 50)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "拆分后的行数/列数必须在 1-50 之间");
                }

                object rowsObj = rows;
                object columnsObj = columns;
                try
                {
                    table.Cell(row, column).Split(ref rowsObj, ref columnsObj);
                }
                catch (Exception ex) when (!(ex is McpToolException))
                {
                    throw new McpToolException(McpErrorCodes.OperationFailed, $"拆分单元格失败：{ex.Message}");
                }

                return DescribeTable(doc, tableIndex);
            });
        }

        public TableInfo SetRowHeight(SetRowHeightRequest request)
        {
            return Mutate("设置行高", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, request.TableIndex);
                var rule = (request.Rule ?? "at-least").Trim().ToLowerInvariant();
                WdRowHeightRule heightRule;
                switch (rule)
                {
                    case "auto":
                        heightRule = WdRowHeightRule.wdRowHeightAuto;
                        break;
                    case "at-least":
                    case "atleast":
                        heightRule = WdRowHeightRule.wdRowHeightAtLeast;
                        break;
                    case "exactly":
                    case "exact":
                        heightRule = WdRowHeightRule.wdRowHeightExactly;
                        break;
                    default:
                        throw new McpToolException(McpErrorCodes.UnsupportedType,
                            $"不支持的行高规则：{request.Rule}（支持 auto / at-least / exactly）");
                }

                if (heightRule != WdRowHeightRule.wdRowHeightAuto &&
                    (!request.Height.HasValue || request.Height.Value <= 0))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams,
                        "rule 为 at-least/exactly 时必须提供 height（磅）");
                }

                if (request.Row.HasValue)
                {
                    var rowCount = table.Rows.Count;
                    if (request.Row.Value < 1 || request.Row.Value > rowCount)
                    {
                        throw new McpToolException(McpErrorCodes.TableCellOutOfRange,
                            $"行索引 {request.Row.Value} 超出范围（当前共 {rowCount} 行）");
                    }

                    var row = table.Rows[request.Row.Value];
                    row.HeightRule = heightRule;
                    if (request.Height.HasValue)
                    {
                        row.Height = (float)request.Height.Value;
                    }
                }
                else
                {
                    table.Rows.HeightRule = heightRule;
                    if (request.Height.HasValue)
                    {
                        table.Rows.Height = (float)request.Height.Value;
                    }
                }

                return DescribeTable(doc, request.TableIndex);
            });
        }

        public TableInfo SetTableStyle(int tableIndex, string style)
        {
            return Mutate("设置表格样式", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, tableIndex);
                if (string.IsNullOrWhiteSpace(style))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：style");
                }

                ApplyTableStyle(doc, table, style);
                return DescribeTable(doc, tableIndex);
            });
        }

        // ============================== P0 扩展：域与文档属性 ==============================

        public FieldUpdateResult UpdateFields()
        {
            return Mutate("更新域", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                doc.Fields.Update();

                var sectionCount = doc.Sections.Count;
                for (var s = 1; s <= sectionCount; s++)
                {
                    Section section = null;
                    try
                    {
                        section = doc.Sections[s];
                        foreach (var headerIndex in new[]
                                 {
                                     WdHeaderFooterIndex.wdHeaderFooterPrimary,
                                     WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                                     WdHeaderFooterIndex.wdHeaderFooterEvenPages,
                                 })
                        {
                            UpdateHeaderFooterFields(section.Headers[headerIndex]);
                            UpdateHeaderFooterFields(section.Footers[headerIndex]);
                        }
                    }
                    catch
                    {
                        // 单节更新失败不影响其他节
                    }
                    finally
                    {
                        TryRelease(section);
                    }
                }

                var tocCount = doc.TablesOfContents.Count;
                for (var i = 1; i <= tocCount; i++)
                {
                    TableOfContents toc = null;
                    try
                    {
                        toc = doc.TablesOfContents[i];
                        toc.Update();
                    }
                    finally
                    {
                        TryRelease(toc);
                    }
                }

                return new FieldUpdateResult
                {
                    FieldsUpdated = true,
                    TocCount = tocCount,
                    SectionCount = sectionCount,
                    Message = $"已更新正文域、{sectionCount} 个节的页眉页脚域与 {tocCount} 个目录",
                };
            });
        }

        public DocumentPropertySet GetDocumentProperties()
        {
            return Run("读取文档属性", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                return ReadDocumentProperties(doc);
            });
        }

        public DocumentPropertySet SetDocumentProperties(SetDocumentPropertiesRequest request)
        {
            return Mutate("设置文档属性", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var props = (OfficeCore.DocumentProperties)doc.BuiltInDocumentProperties;
                WriteDocumentProperty(props, "Title", request.Title);
                WriteDocumentProperty(props, "Author", request.Author);
                WriteDocumentProperty(props, "Subject", request.Subject);
                WriteDocumentProperty(props, "Keywords", request.Keywords);
                WriteDocumentProperty(props, "Comments", request.Comments);
                WriteDocumentProperty(props, "Category", request.Category);
                WriteDocumentProperty(props, "Company", request.Company);
                WriteDocumentProperty(props, "Manager", request.Manager);

                if (request.Custom != null && request.Custom.Count > 0)
                {
                    var custom = (OfficeCore.DocumentProperties)doc.CustomDocumentProperties;
                    foreach (var pair in request.Custom)
                    {
                        WriteCustomDocumentProperty(custom, pair.Key, pair.Value);
                    }
                }

                return ReadDocumentProperties(doc);
            });
        }

        // ============================== P0 扩展：样式管理 ==============================

        public StyleInfo CreateStyle(StyleFormatRequest request)
        {
            return Mutate("创建样式", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(request.Name))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：name");
                }

                var existing = FindStyleOrNull(doc, request.Name);
                if (existing != null)
                {
                    TryRelease(existing);
                    throw new McpToolException(McpErrorCodes.InvalidParams,
                        $"样式已存在：{request.Name}（可用 modify_style 修改）");
                }

                object styleType = ParseStyleType(request.Type);
                Style style = null;
                try
                {
                    style = doc.Styles.Add(request.Name, ref styleType);
                    if (!string.IsNullOrEmpty(request.BaseStyle))
                    {
                        object baseStyle = request.BaseStyle;
                        style.set_BaseStyle(ref baseStyle);
                    }

                    ApplyStyleFormatting(style, request);
                    return DescribeStyle(style);
                }
                finally
                {
                    TryRelease(style);
                }
            });
        }

        public StyleInfo ModifyStyle(StyleFormatRequest request)
        {
            return Mutate("修改样式", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(request.Name))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：name");
                }

                var style = FindStyleOrNull(doc, request.Name);
                if (style == null)
                {
                    throw new McpToolException(McpErrorCodes.StyleNotFound, $"找不到样式：{request.Name}");
                }

                try
                {
                    if (!string.IsNullOrEmpty(request.BaseStyle))
                    {
                        object baseStyle = request.BaseStyle;
                        style.set_BaseStyle(ref baseStyle);
                    }

                    ApplyStyleFormatting(style, request);
                    return DescribeStyle(style);
                }
                finally
                {
                    TryRelease(style);
                }
            });
        }

        public void DeleteStyle(string name)
        {
            Mutate("删除样式", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var style = FindStyleOrNull(doc, name);
                if (style == null)
                {
                    throw new McpToolException(McpErrorCodes.StyleNotFound, $"找不到样式：{name}");
                }

                try
                {
                    if (style.BuiltIn)
                    {
                        throw new McpToolException(McpErrorCodes.InvalidParams, $"内置样式不能删除：{name}");
                    }

                    style.Delete();
                    return 0;
                }
                finally
                {
                    TryRelease(style);
                }
            });
        }

        // ============================== P0 扩展：水印 / 修订 / 脚注 ==============================

        public WatermarkResult AddWatermark(AddWatermarkRequest request)
        {
            return Mutate("添加水印", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(request.Text))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：text");
                }

                var layout = (request.Layout ?? "diagonal").Trim().ToLowerInvariant();
                if (layout != "diagonal" && layout != "horizontal")
                {
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的水印布局：{request.Layout}（支持 diagonal / horizontal）");
                }

                var transparency = request.Transparency ?? 0.5;
                if (transparency < 0 || transparency > 1)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "transparency 必须在 0-1 之间");
                }

                var text = NormalizeText(request.Text) ?? string.Empty;
                var fontName = string.IsNullOrWhiteSpace(request.FontName) ? "微软雅黑" : request.FontName;
                var fontSize = request.FontSize ?? Math.Max(24, Math.Min(72, 480.0 / Math.Max(1, text.Length)));
                if (fontSize < 1 || fontSize > 300)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "font_size 必须在 1-300 之间");
                }

                var color = ArgbToOleRgb(request.ColorArgb ?? unchecked((int)0xFFC0C0C0));
                var rotation = layout == "diagonal" ? 315f : 0f;

                // 幂等：先移除旧水印
                RemoveWatermarkCore(doc);

                var sectionCount = doc.Sections.Count;
                var appliedSections = 0;
                for (var s = 1; s <= sectionCount; s++)
                {
                    Section section = null;
                    try
                    {
                        section = doc.Sections[s];
                        var header = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];

                        // 链接到前一节的页眉会继承前一节的水印，直接写入会 E_ACCESSDENIED
                        if (s > 1 && SafeRead(() => header.LinkToPrevious))
                        {
                            continue;
                        }

                        var shape = header.Shapes.AddTextEffect(OfficeCore.MsoPresetTextEffect.msoTextEffect1,
                            text, fontName, (float)fontSize, OfficeCore.MsoTriState.msoFalse,
                            OfficeCore.MsoTriState.msoFalse, 0, 0);
                        try
                        {
                            shape.Name = WatermarkShapeName;
                            try
                            {
                                shape.TextEffect.NormalizedHeight = OfficeCore.MsoTriState.msoFalse;
                            }
                            catch
                            {
                                // 部分版本不支持，忽略
                            }

                            shape.Line.Visible = OfficeCore.MsoTriState.msoFalse;
                            shape.Fill.Visible = OfficeCore.MsoTriState.msoTrue;
                            shape.Fill.ForeColor.RGB = color;
                            shape.Fill.Transparency = (float)transparency;
                            shape.Rotation = rotation;
                            shape.WrapFormat.Type = WdWrapType.wdWrapNone;
                            shape.WrapFormat.AllowOverlap = -1;
                            shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                            shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                            shape.Left = (float)WdShapePosition.wdShapeCenter;
                            shape.Top = (float)WdShapePosition.wdShapeCenter;
                            shape.LockAspectRatio = OfficeCore.MsoTriState.msoTrue;
                        }
                        finally
                        {
                            TryRelease(shape);
                        }

                        appliedSections++;
                    }
                    catch
                    {
                        // 单节失败继续处理其他节
                    }
                    finally
                    {
                        TryRelease(section);
                    }
                }

                if (appliedSections == 0)
                {
                    throw new McpToolException(McpErrorCodes.OperationFailed, "添加水印失败：无法写入任何节的页眉");
                }

                return new WatermarkResult
                {
                    Applied = true,
                    Text = request.Text,
                    Layout = layout,
                    SectionCount = appliedSections,
                };
            });
        }

        public WatermarkResult RemoveWatermark()
        {
            return Mutate("删除水印", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var removed = RemoveWatermarkCore(doc);
                return new WatermarkResult
                {
                    Applied = false,
                    SectionCount = doc.Sections.Count,
                    RemovedCount = removed,
                };
            });
        }

        public RevisionResult SetTrackChanges(bool enabled)
        {
            return Mutate("设置修订跟踪", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                doc.TrackRevisions = enabled;
                return new RevisionResult
                {
                    TrackChanges = enabled,
                    RevisionCount = doc.Revisions.Count,
                    Message = enabled ? "已开启修订跟踪（后续修改将保留痕迹）" : "已关闭修订跟踪",
                };
            });
        }

        public RevisionResult AcceptRevisions()
        {
            return Mutate("接受所有修订", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var count = doc.Revisions.Count;
                doc.AcceptAllRevisions();
                return new RevisionResult
                {
                    TrackChanges = doc.TrackRevisions,
                    RevisionCount = 0,
                    Message = count == 0 ? "没有待处理的修订" : $"已接受 {count} 处修订",
                };
            });
        }

        public RevisionResult RejectRevisions()
        {
            return Mutate("拒绝所有修订", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var count = doc.Revisions.Count;
                doc.RejectAllRevisions();
                return new RevisionResult
                {
                    TrackChanges = doc.TrackRevisions,
                    RevisionCount = 0,
                    Message = count == 0 ? "没有待处理的修订" : $"已拒绝 {count} 处修订",
                };
            });
        }

        public NoteInfo InsertFootnote(string text, int? paragraphIndex, string anchorText)
        {
            return InsertNoteCore("footnote", text, paragraphIndex, anchorText);
        }

        // ============================== 论文增强：注释 / 超链接 / 交叉引用 / 题注 / 域 ==============================

        public NoteInfo InsertEndnote(string text, int? paragraphIndex, string anchorText)
        {
            return InsertNoteCore("endnote", text, paragraphIndex, anchorText);
        }

        private NoteInfo InsertNoteCore(string noteType, string text, int? paragraphIndex, string anchorText)
        {
            var type = NormalizeNoteType(noteType);
            return Mutate(type == "footnote" ? "插入脚注" : "插入尾注", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrEmpty(text))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：text");
                }

                Range referenceRange;
                if (!string.IsNullOrEmpty(anchorText))
                {
                    referenceRange = FindFirstRange(doc, anchorText);
                    if (referenceRange == null)
                    {
                        throw new McpToolException(McpErrorCodes.TextNotFound, $"未找到角标锚点文本：{anchorText}");
                    }

                    referenceRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
                else
                {
                    var index = paragraphIndex ?? doc.Paragraphs.Count;
                    referenceRange = GetParagraphEndRange(doc, index);
                }

                object noteText = NormalizeText(text) ?? string.Empty;
                object referenceMark = Type.Missing;
                int noteIndex;
                int referenceStart;
                string noteContent;
                int noteCount;

                if (type == "footnote")
                {
                    Footnote note = null;
                    try
                    {
                        // PIA 参数顺序为 (Range, Reference, Text)：Reference 是自定义引用标记，Text 是注释内容
                        note = doc.Footnotes.Add(referenceRange, ref referenceMark, ref noteText);
                        noteIndex = note.Index;
                        referenceStart = SafeRead(() => note.Reference.Start);
                        noteContent = TrimParagraphText(SafeRead(() => note.Range.Text));
                        noteCount = doc.Footnotes.Count;
                    }
                    finally
                    {
                        TryRelease(note);
                    }
                }
                else
                {
                    Endnote note = null;
                    try
                    {
                        note = doc.Endnotes.Add(referenceRange, ref referenceMark, ref noteText);
                        noteIndex = note.Index;
                        referenceStart = SafeRead(() => note.Reference.Start);
                        noteContent = TrimParagraphText(SafeRead(() => note.Range.Text));
                        noteCount = doc.Endnotes.Count;
                    }
                    finally
                    {
                        TryRelease(note);
                    }
                }

                var referenceParagraph = GetParagraphIndexAt(doc, referenceStart);
                return new NoteInfo
                {
                    NoteType = type,
                    Index = noteIndex,
                    Text = noteContent,
                    ReferenceParagraphIndex = referenceParagraph,
                    ReferenceText = GetParagraphPreview(doc, referenceParagraph),
                    NoteCount = noteCount,
                };
            });
        }

        public List<NoteInfo> ListNotes(string noteType)
        {
            return Run("列出注释", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var type = NormalizeNoteType(noteType);
                var result = new List<NoteInfo>();
                var count = type == "footnote" ? doc.Footnotes.Count : doc.Endnotes.Count;
                for (var i = 1; i <= count && result.Count < 1000; i++)
                {
                    result.Add(DescribeNote(doc, type, i));
                }

                return result;
            });
        }

        public NoteInfo GetNote(string noteType, int index)
        {
            return Run("读取注释", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var type = NormalizeNoteType(noteType);
                ValidateNoteIndex(doc, type, index);
                return DescribeNote(doc, type, index);
            });
        }

        public void DeleteNote(string noteType, int index)
        {
            Mutate("删除注释", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var type = NormalizeNoteType(noteType);
                ValidateNoteIndex(doc, type, index);
                if (type == "footnote")
                {
                    doc.Footnotes[index].Delete();
                }
                else
                {
                    doc.Endnotes[index].Delete();
                }

                return 0;
            });
        }

        public NoteOptionsInfo SetNoteOptions(SetNoteOptionsRequest request)
        {
            return Mutate("设置注释编号", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var type = NormalizeNoteType(request.NoteType);

                if (type == "footnote")
                {
                    var notes = doc.Footnotes;
                    if (!string.IsNullOrEmpty(request.NumberStyle))
                    {
                        notes.NumberStyle = ParseNoteNumberStyle(request.NumberStyle);
                    }

                    if (!string.IsNullOrEmpty(request.NumberingRule))
                    {
                        notes.NumberingRule = ParseNumberingRule(request.NumberingRule);
                    }

                    if (request.StartingNumber.HasValue)
                    {
                        ValidateStartingNumber(request.StartingNumber.Value);
                        notes.StartingNumber = request.StartingNumber.Value;
                    }

                    if (!string.IsNullOrEmpty(request.Location))
                    {
                        notes.Location = ParseFootnoteLocation(request.Location);
                    }
                }
                else
                {
                    var notes = doc.Endnotes;
                    if (!string.IsNullOrEmpty(request.NumberStyle))
                    {
                        notes.NumberStyle = ParseNoteNumberStyle(request.NumberStyle);
                    }

                    if (!string.IsNullOrEmpty(request.NumberingRule))
                    {
                        notes.NumberingRule = ParseNumberingRule(request.NumberingRule);
                    }

                    if (request.StartingNumber.HasValue)
                    {
                        ValidateStartingNumber(request.StartingNumber.Value);
                        notes.StartingNumber = request.StartingNumber.Value;
                    }
                }

                return DescribeNoteOptions(doc, type);
            });
        }

        public List<HyperlinkInfo> ListHyperlinks()
        {
            return Run("列出超链接", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var result = new List<HyperlinkInfo>();
                var count = doc.Hyperlinks.Count;
                for (var i = 1; i <= count && result.Count < 1000; i++)
                {
                    result.Add(DescribeHyperlink(doc, i));
                }

                return result;
            });
        }

        public HyperlinkInfo UpdateHyperlink(UpdateHyperlinkRequest request)
        {
            return Mutate("修改超链接", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var count = doc.Hyperlinks.Count;
                if (request.Index < 1 || request.Index > count)
                {
                    throw new McpToolException(McpErrorCodes.HyperlinkNotFound,
                        $"超链接索引 {request.Index} 超出范围（当前共 {count} 个）");
                }

                Hyperlink link = null;
                try
                {
                    link = doc.Hyperlinks[request.Index];
                    if (request.Text != null) link.TextToDisplay = request.Text;
                    if (request.Address != null) link.Address = request.Address;
                    if (request.Bookmark != null) link.SubAddress = request.Bookmark;
                    if (request.ScreenTip != null) link.ScreenTip = request.ScreenTip;
                    return DescribeHyperlink(doc, request.Index);
                }
                finally
                {
                    TryRelease(link);
                }
            });
        }

        public int DeleteHyperlinks(int? index, bool all)
        {
            return Mutate("删除超链接", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var count = doc.Hyperlinks.Count;
                if (all)
                {
                    for (var i = count; i >= 1; i--)
                    {
                        doc.Hyperlinks[i].Delete();
                    }

                    return count;
                }

                if (!index.HasValue)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "必须提供 index 或 all=true");
                }

                if (index.Value < 1 || index.Value > count)
                {
                    throw new McpToolException(McpErrorCodes.HyperlinkNotFound,
                        $"超链接索引 {index.Value} 超出范围（当前共 {count} 个）");
                }

                doc.Hyperlinks[index.Value].Delete();
                return 1;
            });
        }

        public CrossReferenceResult InsertCrossReference(CrossReferenceRequest request)
        {
            return Mutate("插入交叉引用", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var type = (request.ReferenceType ?? string.Empty).Trim().ToLowerInvariant();
                object referenceType;
                switch (type)
                {
                    case "heading":
                        referenceType = WdReferenceType.wdRefTypeHeading;
                        break;
                    case "bookmark":
                        referenceType = WdReferenceType.wdRefTypeBookmark;
                        break;
                    case "footnote":
                        referenceType = WdReferenceType.wdRefTypeFootnote;
                        break;
                    case "endnote":
                        referenceType = WdReferenceType.wdRefTypeEndnote;
                        break;
                    case "caption":
                        if (string.IsNullOrWhiteSpace(request.CaptionLabel))
                        {
                            throw new McpToolException(McpErrorCodes.InvalidParams,
                                "caption 类型需要 caption_label（如 图 / 表）");
                        }

                        referenceType = ResolveCaptionLabel(request.CaptionLabel);
                        break;
                    default:
                        throw new McpToolException(McpErrorCodes.UnsupportedType,
                            $"不支持的引用类型：{request.ReferenceType}（支持 heading / bookmark / footnote / endnote / caption）");
                }

                if (string.IsNullOrWhiteSpace(request.ReferenceItem))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：reference_item");
                }

                object referenceItem;
                if (type == "footnote" || type == "endnote" || type == "caption")
                {
                    if (!int.TryParse(request.ReferenceItem.Trim(), out var ordinal) || ordinal < 1)
                    {
                        throw new McpToolException(McpErrorCodes.InvalidParams,
                            $"{type} 的 reference_item 必须是正整数编号");
                    }

                    referenceItem = ordinal;
                }
                else
                {
                    referenceItem = request.ReferenceItem.Trim();
                }

                var kind = ParseReferenceKind(request.ReferenceKind, type);
                Range target;
                int paragraphIndex;
                if (request.ParagraphIndex.HasValue)
                {
                    paragraphIndex = request.ParagraphIndex.Value;
                    target = GetParagraphEndRange(doc, paragraphIndex);
                }
                else
                {
                    paragraphIndex = AddParagraphCore(doc, null, null);
                    target = doc.Paragraphs[paragraphIndex].Range;
                    target.Collapse(WdCollapseDirection.wdCollapseStart);
                }

                object asHyperlink = request.InsertAsHyperlink;
                object includePosition = false;
                object separateNumbers = false;
                object separator = ", ";

                try
                {
                    target.InsertCrossReference(ref referenceType, kind, ref referenceItem, ref asHyperlink,
                        ref includePosition, ref separateNumbers, ref separator);
                }
                catch (Exception ex) when (!(ex is McpToolException))
                {
                    throw new McpToolException(McpErrorCodes.OperationFailed,
                        $"插入交叉引用失败：{ex.Message}（请确认引用目标存在且类型匹配）");
                }

                return new CrossReferenceResult
                {
                    Inserted = true,
                    ReferenceType = type,
                    ReferenceKind = DescribeReferenceKind(kind),
                    ParagraphIndex = paragraphIndex,
                    ParagraphCount = doc.Paragraphs.Count,
                };
            });
        }

        public CaptionResult InsertCaption(InsertCaptionRequest request)
        {
            return Mutate("插入题注", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var label = ResolveCaptionLabel(request.Label);
                var position = (request.Position ?? "below").Trim().ToLowerInvariant();
                if (position != "above" && position != "below")
                {
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的题注位置：{request.Position}（支持 above / below）");
                }

                var targetIndex = request.ParagraphIndex ?? doc.Paragraphs.Count;
                var paragraph = GetParagraphOrThrow(doc, targetIndex);
                var range = paragraph.Range;

                var title = NormalizeText(request.Title) ?? string.Empty;
                if (title.Length > 0 && !char.IsWhiteSpace(title[0]))
                {
                    title = " " + title;
                }

                object labelObj = label;
                object titleObj = title;
                object titleAutoText = Type.Missing;
                object positionObj = position == "above"
                    ? WdCaptionPosition.wdCaptionPositionAbove
                    : WdCaptionPosition.wdCaptionPositionBelow;
                object excludeLabel = false;

                try
                {
                    range.InsertCaption(ref labelObj, ref titleObj, ref titleAutoText, ref positionObj, ref excludeLabel);
                }
                catch (Exception ex) when (!(ex is McpToolException))
                {
                    throw new McpToolException(McpErrorCodes.OperationFailed, $"插入题注失败：{ex.Message}");
                }

                // 定位刚插入的题注段落（InsertCaption 在目标段落前/后插入新段落）
                var captionIndex = -1;
                var searchFrom = Math.Max(1, targetIndex - 1);
                var searchTo = Math.Min(doc.Paragraphs.Count, targetIndex + 2);
                for (var i = searchFrom; i <= searchTo; i++)
                {
                    var captionText = SafeRead(() => doc.Paragraphs[i].Range.Text) ?? string.Empty;
                    if (captionText.Contains(label) &&
                        (string.IsNullOrEmpty(request.Title) || captionText.Contains(request.Title)))
                    {
                        captionIndex = i;
                        break;
                    }
                }

                if (captionIndex < 0)
                {
                    captionIndex = position == "above"
                        ? targetIndex
                        : Math.Min(targetIndex + 1, doc.Paragraphs.Count);
                }

                return new CaptionResult
                {
                    Label = label,
                    Title = request.Title,
                    Position = position,
                    ParagraphIndex = captionIndex,
                    ParagraphCount = doc.Paragraphs.Count,
                };
            });
        }

        public FieldInfoResult InsertField(InsertFieldRequest request)
        {
            return Mutate("插入域", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(request.FieldCode))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：field_code");
                }

                Range target;
                if (request.ParagraphIndex.HasValue)
                {
                    target = GetParagraphEndRange(doc, request.ParagraphIndex.Value);
                }
                else
                {
                    var index = AddParagraphCore(doc, null, null);
                    target = doc.Paragraphs[index].Range;
                    target.Collapse(WdCollapseDirection.wdCollapseStart);
                }

                object fieldType = WdFieldType.wdFieldEmpty;
                object code = request.FieldCode.Trim();
                object preserveFormatting = false;
                Field field = null;
                try
                {
                    field = doc.Fields.Add(target, ref fieldType, ref code, ref preserveFormatting);
                    if (request.Update)
                    {
                        try
                        {
                            field.Update();
                        }
                        catch
                        {
                            // 域暂时无法更新（如引用目标尚不存在）时保留原始代码
                        }
                    }

                    return new FieldInfoResult
                    {
                        Index = FindFieldIndex(doc, field),
                        Code = SafeRead(() => field.Code.Text)?.Trim(),
                        Result = SafeRead(() => field.Result.Text)?.Trim(),
                        FieldCount = doc.Fields.Count,
                    };
                }
                finally
                {
                    TryRelease(field);
                }
            });
        }


        // ============================== P0 增强：选区与内容移动 ==============================

        public SelectionInfo GetSelection()
        {
            return Run("读取选区", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                return DescribeSelection(doc);
            });
        }

        public ParagraphInfo ReplaceSelection(string text, string style)
        {
            return Mutate("替换选区", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var selection = _app.Selection;
                selection.TypeText(NormalizeText(text) ?? string.Empty);
                if (!string.IsNullOrEmpty(style))
                {
                    object styleObj = ResolveStyleValue(doc, style);
                    var range = selection.Range;
                    range.set_Style(ref styleObj);
                }

                var index = GetParagraphIndexAt(doc, selection.End);
                if (index < 1 || index > doc.Paragraphs.Count)
                {
                    index = doc.Paragraphs.Count;
                }

                return DescribeParagraph(doc, index);
            });
        }

        public SelectionInfo GotoPage(int page)
        {
            return Mutate("跳转到页", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (page < 1)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "page 必须 >= 1");
                }

                object what = WdGoToItem.wdGoToPage;
                object which = WdGoToDirection.wdGoToAbsolute;
                object count = page;
                object name = Type.Missing;
                try
                {
                    _app.Selection.GoTo(ref what, ref which, ref count, ref name);
                }
                catch (Exception ex) when (!(ex is McpToolException))
                {
                    throw new McpToolException(McpErrorCodes.OperationFailed,
                        $"跳转到第 {page} 页失败：{ex.Message}");
                }

                return DescribeSelection(doc);
            });
        }

        public ParagraphInfo MoveParagraph(int paragraphIndex, int offset)
        {
            return Mutate("移动段落", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (offset == 0)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "offset 不能为 0");
                }

                var count = doc.Paragraphs.Count;
                var paragraph = GetParagraphOrThrow(doc, paragraphIndex);
                if (SafeRead(() => (bool)paragraph.Range.Information[WdInformation.wdWithInTable]))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "表格内的段落不支持移动");
                }

                var target = paragraphIndex + offset;
                if (target < 1 || target > count)
                {
                    throw new McpToolException(McpErrorCodes.ParagraphOutOfRange,
                        $"移动后的位置 {target} 超出范围（当前共 {count} 段）");
                }

                var source = doc.Paragraphs[paragraphIndex].Range;

                // 插入位置（基于移动前的文档结构）
                var appendAtEnd = offset > 0 && target >= count;
                if (appendAtEnd)
                {
                    // 文末插入需保留段落标记与格式，单独处理
                    AppendParagraphRange(doc, source);
                }
                else
                {
                    Range insertAt;
                    if (offset < 0)
                    {
                        insertAt = doc.Paragraphs[target].Range;
                        insertAt.Collapse(WdCollapseDirection.wdCollapseStart);
                    }
                    else
                    {
                        insertAt = doc.Paragraphs[target + 1].Range;
                        insertAt.Collapse(WdCollapseDirection.wdCollapseStart);
                    }

                    var formatted = source.FormattedText;
                    insertAt.FormattedText = formatted;
                }

                // 插入到源段落之前时，源段落索引后移 1
                var deleteIndex = offset < 0 ? paragraphIndex + 1 : paragraphIndex;
                doc.Paragraphs[deleteIndex].Range.Delete();

                var finalIndex = offset < 0
                    ? target
                    : (appendAtEnd ? doc.Paragraphs.Count : target);
                return DescribeParagraph(doc, finalIndex);
            });
        }

        public ParagraphInfo CopyRange(CopyRangeRequest request)
        {
            return Mutate("复制内容", () =>
            {
                var doc = GetActiveDocumentOrThrow();

                Range source;
                if (!string.IsNullOrEmpty(request.FindText))
                {
                    source = FindFirstRange(doc, request.FindText);
                    if (source == null)
                    {
                        throw new McpToolException(McpErrorCodes.TextNotFound,
                            $"未找到要复制的内容：{request.FindText}");
                    }
                }
                else
                {
                    var startIndex = request.SourceStartParagraph ?? 1;
                    var endIndex = request.SourceEndParagraph ?? startIndex;
                    if (endIndex < startIndex)
                    {
                        var temp = startIndex;
                        startIndex = endIndex;
                        endIndex = temp;
                    }

                    GetParagraphOrThrow(doc, startIndex);
                    GetParagraphOrThrow(doc, endIndex);
                    object startObj = doc.Paragraphs[startIndex].Range.Start;
                    object endObj = doc.Paragraphs[endIndex].Range.End;
                    source = doc.Range(ref startObj, ref endObj);
                }

                int start;
                if (request.TargetParagraph.HasValue)
                {
                    var insertAt = GetParagraphOrThrow(doc, request.TargetParagraph.Value).Range;
                    insertAt.Collapse(WdCollapseDirection.wdCollapseStart);
                    start = insertAt.Start;
                    insertAt.FormattedText = source.FormattedText;
                }
                else
                {
                    // 文末追加需保留段落标记与格式，单独处理
                    start = doc.Content.End;
                    AppendParagraphRange(doc, source);
                }

                var index = GetParagraphIndexAt(doc, start);
                if (index < 1 || index > doc.Paragraphs.Count)
                {
                    index = doc.Paragraphs.Count;
                }

                return DescribeParagraph(doc, index);
            });
        }

        // ============================== P0 增强：图片 ==============================

        public List<ImageEntry> ListImages()
        {
            return Run("列出图片", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var result = new List<ImageEntry>();

                for (var i = 1; i <= doc.InlineShapes.Count; i++)
                {
                    InlineShape shape = null;
                    try
                    {
                        shape = doc.InlineShapes[i];
                        var type = SafeRead(() => shape.Type);
                        if (type != WdInlineShapeType.wdInlineShapePicture &&
                            type != WdInlineShapeType.wdInlineShapeLinkedPicture)
                        {
                            continue;
                        }

                        var rangeStart = SafeRead(() => shape.Range.Start);
                        result.Add(new ImageEntry
                        {
                            Kind = "inline",
                            Index = i,
                            Width = SafeRead(() => (double)shape.Width),
                            Height = SafeRead(() => (double)shape.Height),
                            AltText = SafeRead(() => shape.AlternativeText),
                            ParagraphIndex = GetParagraphIndexAt(doc, rangeStart),
                        });
                    }
                    finally
                    {
                        TryRelease(shape);
                    }
                }

                for (var i = 1; i <= doc.Shapes.Count; i++)
                {
                    Shape shape = null;
                    try
                    {
                        shape = doc.Shapes[i];
                        var type = SafeRead(() => shape.Type);
                        if (type != OfficeCore.MsoShapeType.msoPicture &&
                            type != OfficeCore.MsoShapeType.msoLinkedPicture)
                        {
                            continue;
                        }

                        var anchorStart = SafeRead(() => shape.Anchor.Start);
                        result.Add(new ImageEntry
                        {
                            Kind = "floating",
                            Index = i,
                            Name = SafeRead(() => shape.Name),
                            Width = SafeRead(() => (double)shape.Width),
                            Height = SafeRead(() => (double)shape.Height),
                            Left = SafeRead(() => (double?)shape.Left),
                            Top = SafeRead(() => (double?)shape.Top),
                            Wrap = SafeRead(() => DescribeWrapType(shape.WrapFormat.Type)),
                            AltText = SafeRead(() => shape.AlternativeText),
                            ParagraphIndex = GetParagraphIndexAt(doc, anchorStart),
                        });
                    }
                    finally
                    {
                        TryRelease(shape);
                    }
                }

                return result;
            });
        }

        public ImageEntry UpdateImage(UpdateImageRequest request)
        {
            return Mutate("修改图片", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var kind = (request.Kind ?? "inline").Trim().ToLowerInvariant();

                if (kind == "inline")
                {
                    if (!request.Index.HasValue)
                    {
                        throw new McpToolException(McpErrorCodes.InvalidParams, "inline 图片需要提供 index");
                    }

                    var index = request.Index.Value;
                    if (index < 1 || index > doc.InlineShapes.Count)
                    {
                        throw new McpToolException(McpErrorCodes.ImageNotFound,
                            $"内嵌图片索引 {index} 超出范围（当前共 {doc.InlineShapes.Count} 个）");
                    }

                    InlineShape shape = null;
                    try
                    {
                        shape = doc.InlineShapes[index];
                        if (request.Width.HasValue) shape.Width = (float)request.Width.Value;
                        if (request.Height.HasValue) shape.Height = (float)request.Height.Value;
                        if (request.ScaleWidthPercent.HasValue) shape.ScaleWidth = (float)request.ScaleWidthPercent.Value;
                        if (request.ScaleHeightPercent.HasValue) shape.ScaleHeight = (float)request.ScaleHeightPercent.Value;

                        var picture = shape.PictureFormat;
                        if (request.CropTop.HasValue) picture.CropTop = (float)request.CropTop.Value;
                        if (request.CropBottom.HasValue) picture.CropBottom = (float)request.CropBottom.Value;
                        if (request.CropLeft.HasValue) picture.CropLeft = (float)request.CropLeft.Value;
                        if (request.CropRight.HasValue) picture.CropRight = (float)request.CropRight.Value;

                        if (request.AltText != null) shape.AlternativeText = request.AltText;
                        if (request.Title != null) shape.Title = request.Title;
                        return DescribeInlineImage(doc, index);
                    }
                    finally
                    {
                        TryRelease(shape);
                    }
                }

                if (kind != "floating")
                {
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的图片类型：{request.Kind}（支持 inline / floating）");
                }

                Shape floating = FindShapeOrNull(doc, request.Name, request.Index);
                if (floating == null)
                {
                    throw new McpToolException(McpErrorCodes.ImageNotFound,
                        $"找不到浮动图片：{(string.IsNullOrEmpty(request.Name) ? "#" + request.Index : request.Name)}");
                }

                try
                {
                    if (request.Width.HasValue) floating.Width = (float)request.Width.Value;
                    if (request.Height.HasValue) floating.Height = (float)request.Height.Value;
                    if (request.AltText != null) floating.AlternativeText = request.AltText;
                    if (request.Title != null) floating.Title = request.Title;
                    return DescribeFloatingImage(doc, floating);
                }
                finally
                {
                    TryRelease(floating);
                }
            });
        }

        public void DeleteImage(string kind, int? index, string name)
        {
            Mutate("删除图片", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var type = (kind ?? "inline").Trim().ToLowerInvariant();
                if (type == "inline")
                {
                    if (!index.HasValue)
                    {
                        throw new McpToolException(McpErrorCodes.InvalidParams, "inline 图片需要提供 index");
                    }

                    if (index.Value < 1 || index.Value > doc.InlineShapes.Count)
                    {
                        throw new McpToolException(McpErrorCodes.ImageNotFound,
                            $"内嵌图片索引 {index.Value} 超出范围（当前共 {doc.InlineShapes.Count} 个）");
                    }

                    doc.InlineShapes[index.Value].Delete();
                    return 0;
                }

                if (type != "floating")
                {
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的图片类型：{kind}（支持 inline / floating）");
                }

                var shape = FindShapeOrNull(doc, name, index);
                if (shape == null)
                {
                    throw new McpToolException(McpErrorCodes.ImageNotFound,
                        $"找不到浮动图片：{(string.IsNullOrEmpty(name) ? "#" + index : name)}");
                }

                try
                {
                    shape.Delete();
                    return 0;
                }
                finally
                {
                    TryRelease(shape);
                }
            });
        }

        // ============================== P0 增强：页码 ==============================

        public PageNumbersInfo SetPageNumbers(SetPageNumbersRequest request)
        {
            return Mutate("设置页码格式", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var section = GetSectionOrThrow(doc, request.SectionIndex);
                var numbers = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers;

                if (!string.IsNullOrEmpty(request.NumberStyle))
                {
                    numbers.NumberStyle = ParsePageNumberStyle(request.NumberStyle);
                }

                // 先设置是否重编号，再设置起始编号（顺序反了起始值会被重置为 0）
                if (request.RestartAtSection.HasValue)
                {
                    numbers.RestartNumberingAtSection = request.RestartAtSection.Value;
                }

                if (request.StartingNumber.HasValue)
                {
                    ValidateStartingNumber(request.StartingNumber.Value);
                    numbers.StartingNumber = request.StartingNumber.Value;
                }

                return DescribePageNumbers(doc, request.SectionIndex);
            });
        }

        // ============================== P0 增强：修订 / 批注 ==============================

        public List<RevisionInfo> ListRevisions()
        {
            return Run("列出修订", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var result = new List<RevisionInfo>();
                for (var i = 1; i <= doc.Revisions.Count && result.Count < 1000; i++)
                {
                    result.Add(DescribeRevision(doc, i));
                }

                return result;
            });
        }

        public RevisionResult AcceptRevision(int index)
        {
            return Mutate("接受修订", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                ValidateRevisionIndex(doc, index);
                var info = DescribeRevision(doc, index);
                doc.Revisions[index].Accept();
                return new RevisionResult
                {
                    TrackChanges = doc.TrackRevisions,
                    RevisionCount = doc.Revisions.Count,
                    Message = $"已接受修订 #{index}（{info.Type}，作者 {info.Author}）",
                };
            });
        }

        public RevisionResult RejectRevision(int index)
        {
            return Mutate("拒绝修订", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                ValidateRevisionIndex(doc, index);
                var info = DescribeRevision(doc, index);
                doc.Revisions[index].Reject();
                return new RevisionResult
                {
                    TrackChanges = doc.TrackRevisions,
                    RevisionCount = doc.Revisions.Count,
                    Message = $"已拒绝修订 #{index}（{info.Type}，作者 {info.Author}）",
                };
            });
        }

        public CommentInfo GetComment(int index)
        {
            return Run("读取批注", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                ValidateCommentIndex(doc, index);
                return DescribeComment(doc, index);
            });
        }

        public CommentInfo ReplyComment(int parentIndex, string text)
        {
            return Mutate("回复批注", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrEmpty(text))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：text");
                }

                ValidateCommentIndex(doc, parentIndex);
                Comment parent = null;
                try
                {
                    parent = doc.Comments[parentIndex];
                    object commentText = NormalizeText(text) ?? string.Empty;
                    var reply = parent.Replies.Add(parent.Range, ref commentText);
                    return DescribeComment(doc, reply.Index);
                }
                finally
                {
                    TryRelease(parent);
                }
            });
        }

        public CommentInfo ResolveComment(int index, bool done)
        {
            return Mutate("标记批注", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                ValidateCommentIndex(doc, index);
                Comment comment = null;
                try
                {
                    comment = doc.Comments[index];
                    comment.Done = done;
                    return DescribeComment(doc, index);
                }
                finally
                {
                    TryRelease(comment);
                }
            });
        }

        // ============================== P0 增强：表格 ==============================

        public TableInfo SetTableBorders(SetTableBordersRequest request)
        {
            return Mutate("设置表格边框", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, request.TableIndex);
                var lineStyle = ParseBorderStyle(request.Style ?? "single");
                var inside = request.Inside ?? true;
                var outside = request.Outside ?? true;

                var types = new List<WdBorderType>();
                if (outside)
                {
                    types.Add(WdBorderType.wdBorderTop);
                    types.Add(WdBorderType.wdBorderBottom);
                    types.Add(WdBorderType.wdBorderLeft);
                    types.Add(WdBorderType.wdBorderRight);
                }

                if (inside)
                {
                    types.Add(WdBorderType.wdBorderHorizontal);
                    types.Add(WdBorderType.wdBorderVertical);
                }

                var borders = table.Borders;
                foreach (var borderType in types)
                {
                    Border border = null;
                    try
                    {
                        border = borders[borderType];
                        border.LineStyle = lineStyle;
                        if (!string.IsNullOrEmpty(request.Width))
                        {
                            border.LineWidth = ParseBorderWidth(request.Width);
                        }

                        if (request.ColorArgb.HasValue)
                        {
                            border.Color = ArgbToWdColor(request.ColorArgb.Value);
                        }
                    }
                    finally
                    {
                        TryRelease(border);
                    }
                }

                return DescribeTable(doc, request.TableIndex);
            });
        }

        public TableInfo SetCellPadding(CellPaddingRequest request)
        {
            return Mutate("设置单元格内边距", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, request.TableIndex);
                if (!request.Top.HasValue && !request.Bottom.HasValue &&
                    !request.Left.HasValue && !request.Right.HasValue)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams,
                        "至少提供 top / bottom / left / right 之一");
                }

                var rowStart = request.Row ?? 1;
                var rowEnd = request.Row ?? table.Rows.Count;
                var columnStart = request.Column ?? 1;
                var columnEnd = request.Column ?? table.Columns.Count;
                ValidateCellRange(table, rowStart, columnStart, "row", "column");
                ValidateCellRange(table, rowEnd, columnEnd, "row", "column");

                for (var r = rowStart; r <= rowEnd; r++)
                {
                    for (var c = columnStart; c <= columnEnd; c++)
                    {
                        Cell cell = null;
                        try
                        {
                            cell = table.Cell(r, c);
                            if (request.Top.HasValue) cell.TopPadding = (float)request.Top.Value;
                            if (request.Bottom.HasValue) cell.BottomPadding = (float)request.Bottom.Value;
                            if (request.Left.HasValue) cell.LeftPadding = (float)request.Left.Value;
                            if (request.Right.HasValue) cell.RightPadding = (float)request.Right.Value;
                        }
                        catch
                        {
                            // 合并单元格的覆盖位置访问会失败，跳过
                        }
                        finally
                        {
                            TryRelease(cell);
                        }
                    }
                }

                return DescribeTable(doc, request.TableIndex);
            });
        }

        public TableInfo SortTable(SortTableRequest request)
        {
            return Mutate("表格排序", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, request.TableIndex);
                if (request.KeyColumn < 1 || request.KeyColumn > table.Columns.Count)
                {
                    throw new McpToolException(McpErrorCodes.TableCellOutOfRange,
                        $"排序列 {request.KeyColumn} 超出范围（当前共 {table.Columns.Count} 列）");
                }

                object excludeHeader = request.HasHeader;
                object fieldNumber = request.KeyColumn;
                object sortFieldType = ParseSortFieldType(request.FieldType);
                object sortOrder = request.Ascending
                    ? WdSortOrder.wdSortOrderAscending
                    : WdSortOrder.wdSortOrderDescending;
                object missing = Type.Missing;

                table.Sort(ref excludeHeader, ref fieldNumber, ref sortFieldType, ref sortOrder,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                return DescribeTable(doc, request.TableIndex);
            });
        }

        public TableInfo ConvertTableToText(int tableIndex, string separator)
        {
            return Mutate("表格转文本", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var table = GetTableOrThrow(doc, tableIndex);
                var info = DescribeTable(doc, tableIndex);
                object separatorObj = string.IsNullOrEmpty(separator) ? "\t" : separator;
                object nestedTables = false;
                table.ConvertToText(ref separatorObj, ref nestedTables);
                info.TableCount = doc.Tables.Count;
                info.Message = "表格已转换为文本";
                return info;
            });
        }

        public TableInfo ConvertTextToTable(int startParagraph, int endParagraph, int columns, string separator)
        {
            return Mutate("文本转表格", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                GetParagraphOrThrow(doc, startParagraph);
                GetParagraphOrThrow(doc, endParagraph);
                if (endParagraph < startParagraph)
                {
                    var temp = startParagraph;
                    startParagraph = endParagraph;
                    endParagraph = temp;
                }

                object startObj = doc.Paragraphs[startParagraph].Range.Start;
                object endObj = doc.Paragraphs[endParagraph].Range.End;
                var range = doc.Range(ref startObj, ref endObj);

                object separatorObj = string.IsNullOrEmpty(separator) ? "\t" : separator;
                object missing = Type.Missing;
                object numColumns = columns > 0 ? (object)columns : missing;

                Table table = null;
                try
                {
                    table = range.ConvertToTable(ref separatorObj, ref missing, ref numColumns,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    return DescribeTable(doc, GetTableIndex(doc, table));
                }
                finally
                {
                    TryRelease(table);
                }
            });
        }

        // ============================== P0 增强：保护 / 封面 ==============================

        public ProtectionInfo ProtectDocument(string protectionType, string password)
        {
            return Mutate("保护文档", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                var type = ParseProtectionType(protectionType);
                object noReset = false;
                object passwordObj = string.IsNullOrEmpty(password) ? Type.Missing : (object)password;
                object useIRM = false;
                object enforceStyleLock = false;
                doc.Protect(type, ref noReset, ref passwordObj, ref useIRM, ref enforceStyleLock);
                return DescribeProtection(doc);
            });
        }

        public ProtectionInfo UnprotectDocument(string password)
        {
            return Mutate("取消文档保护", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                object passwordObj = string.IsNullOrEmpty(password) ? Type.Missing : (object)password;
                doc.Unprotect(ref passwordObj);
                return DescribeProtection(doc);
            });
        }

        public ProtectionInfo MarkAsFinal(bool final)
        {
            return Mutate("标记为最终状态", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (final && string.IsNullOrEmpty(SafeRead(() => doc.Path)))
                {
                    // 未保存文档设置 Final 会弹出“另存为”模态框并阻塞自动化调用
                    throw new McpToolException(McpErrorCodes.OperationFailed,
                        "文档尚未保存，无法标记为最终状态；请先使用 save_document 保存后再标记");
                }

                doc.Final = final;
                return DescribeProtection(doc);
            });
        }

        public CoverPageResult InsertCoverPage(CoverPageRequest request)
        {
            return Mutate("插入封面页", () =>
            {
                var doc = GetActiveDocumentOrThrow();
                if (string.IsNullOrWhiteSpace(request.Title))
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：title");
                }

                // 文档以表格开头时，位置 0 位于首个单元格内，无法直接插入内容；
                // 先插入分节符创建文首段落，封面写入后再删除分节符，避免影响原有节设置
                var startsWithTable = doc.Paragraphs.Count > 0 &&
                    SafeRead(() => (bool)doc.Paragraphs[1].Range.Information[WdInformation.wdWithInTable]);
                var sectionBreakAdded = false;
                if (startsWithTable)
                {
                    object breakStart = 0;
                    object breakEnd = 0;
                    var breakRange = doc.Range(ref breakStart, ref breakEnd);
                    object breakType = WdBreakType.wdSectionBreakContinuous;
                    breakRange.InsertBreak(ref breakType);
                    sectionBreakAdded = true;
                }

                var lines = new List<Tuple<string, string>>
                {
                    Tuple.Create(request.Title, "title"),
                };
                if (!string.IsNullOrWhiteSpace(request.Subtitle))
                {
                    lines.Add(Tuple.Create(request.Subtitle, "subtitle"));
                }

                if (!string.IsNullOrWhiteSpace(request.Author))
                {
                    lines.Add(Tuple.Create(request.Author, "normal"));
                }

                if (!string.IsNullOrWhiteSpace(request.Institution))
                {
                    lines.Add(Tuple.Create(request.Institution, "normal"));
                }

                var dateText = string.IsNullOrWhiteSpace(request.Date)
                    ? DateTime.Now.ToString("yyyy年M月d日")
                    : request.Date;
                lines.Add(Tuple.Create(dateText, "normal"));

                // 在文档最开头（位置 0）插入封面文本；不能按“段落 1”插入，文档可能以表格开头
                object zero = 0;
                object zeroEnd = 0;
                var start = doc.Range(ref zero, ref zeroEnd);
                var coverText = string.Join("\r", lines.ConvertAll(line => line.Item1)) + "\r";
                start.InsertAfter(coverText);

                // 套用样式（标题/副标题/正文）
                for (var i = 1; i <= lines.Count; i++)
                {
                    object styleObj = ResolveStyleValue(doc,
                        lines[i - 1].Item2 == "title" ? "title" : lines[i - 1].Item2 == "subtitle" ? "subtitle" : "normal");
                    doc.Paragraphs[i].Range.set_Style(ref styleObj);
                }

                // 统一居中与间距
                for (var i = 1; i <= lines.Count; i++)
                {
                    var paragraph = doc.Paragraphs[i];
                    var format = paragraph.Range.ParagraphFormat;
                    format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    format.SpaceBefore = 0;
                    format.SpaceAfter = 12;
                    if (i == 1)
                    {
                        format.SpaceBefore = 120;
                        paragraph.Range.Font.Size = 26;
                    }
                    else if (i == 2)
                    {
                        paragraph.Range.Font.Size = 16;
                    }
                }

                if (sectionBreakAdded)
                {
                    // 删除临时分节符：其字符位于第 1 节末尾（封面最后一段之后）
                    var sectionEnd = doc.Sections[1].Range.End;
                    object breakStart = sectionEnd - 1;
                    object breakEnd = sectionEnd;
                    var breakRange = doc.Range(ref breakStart, ref breakEnd);
                    breakRange.Delete();
                }

                if (request.PageBreak)
                {
                    // 在封面最后一段的段落标记之前插入分页符，正文（可能是表格）从新页开始
                    var breakRange = GetParagraphEndRange(doc, lines.Count);
                    object breakType = WdBreakType.wdPageBreak;
                    breakRange.InsertBreak(ref breakType);
                }

                return new CoverPageResult
                {
                    Title = request.Title,
                    CoverParagraphs = lines.Count,
                    ParagraphCount = doc.Paragraphs.Count,
                };
            });
        }

        // ============================== 私有实现 ==============================

        /// <summary>只读操作：封送 UI 线程并统一包装异常。</summary>
        private T Run<T>(string operation, Func<T> action)
        {
            return _marshaller.Invoke(() =>
            {
                try
                {
                    return action();
                }
                catch (McpToolException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    _logger?.Warn($"{operation}失败：{ex.Message}");
                    throw new McpToolException(McpErrorCodes.OperationFailed, $"{operation}失败：{ex.Message}");
                }
            });
        }

        /// <summary>写操作：在只读包装基础上加自定义撤销单元与 ScreenUpdating 抑制。</summary>
        private T Mutate<T>(string operation, Func<T> action)
        {
            return Run(operation, () => WithUndoRecord(operation, () => action()));
        }

        private T WithScreenUpdating<T>(Func<T> action)
        {
            bool original = true;
            try
            {
                original = _app.ScreenUpdating;
                _app.ScreenUpdating = false;
            }
            catch
            {
                // 部分状态下不可写，忽略
            }

            try
            {
                return action();
            }
            finally
            {
                try
                {
                    _app.ScreenUpdating = original;
                }
                catch
                {
                    // 忽略恢复失败
                }
            }
        }

        private T WithUndoRecord<T>(string name, Func<T> action)
        {
            UndoRecord undo = null;
            var started = false;
            try
            {
                try
                {
                    undo = _app.UndoRecord;
                    if (undo != null && !undo.IsRecordingCustomRecord)
                    {
                        undo.StartCustomRecord("MCP: " + name);
                        started = true;
                    }
                }
                catch
                {
                    started = false;
                }

                return action();
            }
            finally
            {
                if (started)
                {
                    try
                    {
                        undo.EndCustomRecord();
                    }
                    catch
                    {
                        // 忽略
                    }
                }
            }
        }

        /// <summary>获取活动文档；没有打开的文档时抛出友好的业务错误。</summary>
        private Document GetActiveDocumentOrThrow()
        {
            // 边界防御：Word 启动初期偶发 Documents 尚未就绪，短暂等待后重试一次
            if (_app.Documents.Count == 0)
            {
                System.Threading.Thread.Sleep(250);
            }

            if (_app.Documents.Count == 0)
            {
                throw new McpToolException(McpErrorCodes.DocumentNotOpen,
                    "当前没有打开的 Word 文档，请先在 Word 中新建或打开一个文档");
            }

            try
            {
                var doc = _app.ActiveDocument;
                if (doc == null)
                {
                    throw new McpToolException(McpErrorCodes.DocumentNotOpen, "当前没有活动文档窗口");
                }

                return doc;
            }
            catch (McpToolException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new McpToolException(McpErrorCodes.DocumentNotOpen, "无法获取当前活动文档：" + ex.Message);
            }
        }

        /// <summary>按名称或索引查找文档（都为空时返回活动文档）。</summary>
        private Document FindDocumentOrThrow(string name, int? index)
        {
            var count = _app.Documents.Count;
            if (count == 0)
            {
                throw new McpToolException(McpErrorCodes.DocumentNotOpen, "当前没有打开的 Word 文档");
            }

            if (index.HasValue)
            {
                if (index.Value < 1 || index.Value > count)
                {
                    throw new McpToolException(McpErrorCodes.DocumentNotFound,
                        $"文档索引 {index.Value} 超出范围（当前共 {count} 个文档）");
                }

                return _app.Documents[index.Value];
            }

            if (!string.IsNullOrWhiteSpace(name))
            {
                for (var i = 1; i <= count; i++)
                {
                    Document doc = null;
                    try
                    {
                        doc = _app.Documents[i];
                        if (string.Equals(doc.Name, name, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(doc.FullName, name, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(Path.GetFileNameWithoutExtension(doc.Name), name, StringComparison.OrdinalIgnoreCase))
                        {
                            return doc;
                        }
                    }
                    catch
                    {
                        // 忽略单个文档读取失败
                    }
                }

                throw new McpToolException(McpErrorCodes.DocumentNotFound, $"找不到文档：{name}");
            }

            return GetActiveDocumentOrThrow();
        }

        private static Paragraph GetParagraphOrThrow(Document doc, int paragraphIndex)
        {
            var count = doc.Paragraphs.Count;
            if (paragraphIndex < 1 || paragraphIndex > count)
            {
                throw new McpToolException(McpErrorCodes.ParagraphOutOfRange,
                    $"段落索引 {paragraphIndex} 超出范围（当前共 {count} 段）");
            }

            return doc.Paragraphs[paragraphIndex];
        }

        /// <summary>
        /// 将源范围（可含多段）追加到文档末尾，保留段落与字符格式。
        /// 注意：不能先新增空段再赋值完整 FormattedText（会多出空段落），
        /// 也不能直接向文档末尾赋值（会丢失段落标记导致段落合并）。
        /// </summary>
        private static void AppendParagraphRange(Document doc, Range source)
        {
            // 必须在插入前确定边界与格式：InsertParagraphAfter 会扩展文档末段的 Range，
            // 使 source 把新增的段落标记也包含进来
            var text = SafeRead(() => source.Text) ?? string.Empty;
            var contentStart = source.Start;
            var contentEnd = text.EndsWith("\r", StringComparison.Ordinal) ? source.End - 1 : source.End;

            object styleObj = null;
            ParagraphFormat sourceFormat = null;
            Paragraph sourceLast = null;
            try
            {
                var paragraphs = source.Paragraphs;
                sourceLast = paragraphs[paragraphs.Count];
                var sourceLastRange = sourceLast.Range;
                styleObj = sourceLastRange.get_Style();
                sourceFormat = sourceLastRange.ParagraphFormat;
            }
            catch (Exception)
            {
                // 段落格式读取失败时保持目标段落默认格式
            }
            finally
            {
                TryRelease(sourceLast);
            }

            var end = doc.Content;
            try
            {
                end.Collapse(WdCollapseDirection.wdCollapseEnd);
                end.InsertParagraphAfter();

                var target = doc.Paragraphs[doc.Paragraphs.Count].Range;
                try
                {
                    // 先复制样式与段落格式（随后写入内容不会覆盖段落格式）
                    if (styleObj != null)
                    {
                        target.set_Style(ref styleObj);
                    }

                    if (sourceFormat != null)
                    {
                        target.ParagraphFormat = sourceFormat;
                    }

                    // 去掉源范围末尾的段落标记，写入目标段落自带的段落标记，避免多出空段落
                    object contentStartObj = contentStart;
                    object contentEndObj = contentEnd;
                    var content = doc.Range(ref contentStartObj, ref contentEndObj);
                    try
                    {
                        target.FormattedText = content.FormattedText;
                    }
                    finally
                    {
                        TryRelease(content);
                    }
                }
                finally
                {
                    TryRelease(target);
                }
            }
            finally
            {
                TryRelease(end);
            }
        }

        /// <summary>
        /// 返回段落末尾（段落标记之前）的折叠范围。
        /// 注意：Paragraph.Range.Collapse(end) 会落在段落标记之后（下一段开头），
        /// 向段落末尾追加内容必须用本方法定位。
        /// </summary>
        private static Range GetParagraphEndRange(Document doc, int paragraphIndex)
        {
            var paragraph = GetParagraphOrThrow(doc, paragraphIndex);
            var paragraphRange = paragraph.Range;
            var position = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            object start = position;
            object end = position;
            return doc.Range(ref start, ref end);
        }

        private static Section GetSectionOrThrow(Document doc, int sectionIndex)
        {
            var count = doc.Sections.Count;
            if (sectionIndex < 1 || sectionIndex > count)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams,
                    $"节索引 {sectionIndex} 超出范围（当前共 {count} 节）");
            }

            return doc.Sections[sectionIndex];
        }

        private static Table GetTableOrThrow(Document doc, int tableIndex)
        {
            var count = doc.Tables.Count;
            if (count == 0)
            {
                throw new McpToolException(McpErrorCodes.TableNotFound, "当前文档没有表格");
            }

            if (tableIndex < 1 || tableIndex > count)
            {
                throw new McpToolException(McpErrorCodes.TableNotFound,
                    $"表格索引 {tableIndex} 超出范围（当前共 {count} 个表格）");
            }

            return doc.Tables[tableIndex];
        }

        /// <summary>在文档末尾追加段落并返回其索引（末尾空段落会被复用，避免产生连续空段）。</summary>
        private static int AddParagraphCore(Document doc, string text, int? position)
        {
            var normalized = NormalizeText(text) ?? string.Empty;
            if (position.HasValue)
            {
                var count = doc.Paragraphs.Count;
                if (position.Value < 1 || position.Value > count + 1)
                {
                    throw new McpToolException(McpErrorCodes.ParagraphOutOfRange,
                        $"插入位置 {position.Value} 超出范围（当前共 {count} 段，允许 1-{count + 1}）");
                }

                if (position.Value <= count)
                {
                    doc.Paragraphs[position.Value].Range.InsertBefore(normalized + "\r");
                    return position.Value;
                }
            }

            return AppendParagraph(doc, normalized);
        }

        /// <summary>追加段落：复用末尾空段落，否则在文档末尾新增空段落再写入。</summary>
        private static int AppendParagraph(Document doc, string text)
        {
            var last = doc.Paragraphs[doc.Paragraphs.Count];
            if (!IsParagraphEmpty(last))
            {
                var end = doc.Content;
                end.Collapse(WdCollapseDirection.wdCollapseEnd);
                end.InsertParagraphAfter();
            }

            var target = doc.Paragraphs[doc.Paragraphs.Count];
            if (!string.IsNullOrEmpty(text))
            {
                target.Range.InsertBefore(text);
            }

            return doc.Paragraphs.Count;
        }

        private static bool IsParagraphEmpty(Paragraph paragraph)
        {
            try
            {
                return TrimParagraphText(paragraph.Range.Text).Length == 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>获取"块级对象"插入位置：指定段落段首，或文档末尾的空段落段首。</summary>
        private static Range GetBlockInsertRange(Document doc, int? position)
        {
            if (position.HasValue)
            {
                var p = GetParagraphOrThrow(doc, position.Value);
                var range = p.Range;
                range.Collapse(WdCollapseDirection.wdCollapseStart);
                return range;
            }

            var index = AppendParagraph(doc, null);
            var anchor = doc.Paragraphs[index].Range;
            anchor.Collapse(WdCollapseDirection.wdCollapseStart);
            return anchor;
        }

        private ParagraphInfo InsertBreakCore(string operation, int? beforeParagraph, string text, WdBreakType breakType)
        {
            return Mutate(operation, () =>
            {
                var doc = GetActiveDocumentOrThrow();
                int index;
                Range target;
                if (beforeParagraph.HasValue)
                {
                    var p = GetParagraphOrThrow(doc, beforeParagraph.Value);
                    index = beforeParagraph.Value;
                    target = p.Range;
                    target.Collapse(WdCollapseDirection.wdCollapseStart);
                }
                else
                {
                    index = AppendParagraph(doc, null);
                    target = doc.Paragraphs[index].Range;
                    target.Collapse(WdCollapseDirection.wdCollapseStart);
                }

                object breakObj = breakType;
                target.InsertBreak(ref breakObj);
                if (!string.IsNullOrEmpty(text))
                {
                    target.InsertAfter(NormalizeText(text));
                }

                return DescribeParagraph(doc, index);
            });
        }

        private static ParagraphInfo DescribeParagraph(Document doc, int paragraphIndex)
        {
            var p = doc.Paragraphs[paragraphIndex];
            return new ParagraphInfo
            {
                ParagraphIndex = paragraphIndex,
                Text = TrimParagraphText(SafeRead(() => p.Range.Text)),
                Style = SafeRead(() => ((Style)p.Range.get_Style())?.NameLocal),
                Alignment = SafeRead(() => DescribeAlignment(p.Alignment)),
                ParagraphCount = doc.Paragraphs.Count,
            };
        }

        private static void ApplyParagraphFormat(ParagraphFormat format, SetParagraphFormatRequest request)
        {
            if (!string.IsNullOrEmpty(request.Alignment))
            {
                format.Alignment = ParseAlignment(request.Alignment);
            }

            if (request.LeftIndent.HasValue) format.LeftIndent = (float)request.LeftIndent.Value;
            if (request.RightIndent.HasValue) format.RightIndent = (float)request.RightIndent.Value;
            if (request.FirstLineIndent.HasValue) format.FirstLineIndent = (float)request.FirstLineIndent.Value;
            if (request.SpaceBefore.HasValue) format.SpaceBefore = (float)request.SpaceBefore.Value;
            if (request.SpaceAfter.HasValue) format.SpaceAfter = (float)request.SpaceAfter.Value;

            if (!string.IsNullOrEmpty(request.LineSpacing))
            {
                switch (request.LineSpacing.Trim().ToLowerInvariant())
                {
                    case "single":
                        format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                        break;
                    case "1.5":
                    case "one-and-half":
                    case "1-5":
                        format.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
                        break;
                    case "double":
                        format.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
                        break;
                    case "at-least":
                        RequireLineSpacingValue(request);
                        format.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                        format.LineSpacing = (float)request.LineSpacingValue.Value;
                        break;
                    case "exactly":
                        RequireLineSpacingValue(request);
                        format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                        format.LineSpacing = (float)request.LineSpacingValue.Value;
                        break;
                    case "multiple":
                        RequireLineSpacingValue(request);
                        format.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                        format.LineSpacing = (float)request.LineSpacingValue.Value;
                        break;
                    default:
                        throw new McpToolException(McpErrorCodes.UnsupportedType,
                            $"不支持的行距模式：{request.LineSpacing}（支持 single / 1.5 / double / at-least / exactly / multiple）");
                }
            }

            if (request.KeepWithNext.HasValue) format.KeepWithNext = request.KeepWithNext.Value ? -1 : 0;
            if (request.KeepTogether.HasValue) format.KeepTogether = request.KeepTogether.Value ? -1 : 0;
            if (request.PageBreakBefore.HasValue) format.PageBreakBefore = request.PageBreakBefore.Value ? -1 : 0;

            if (!string.IsNullOrEmpty(request.OutlineLevel))
            {
                format.OutlineLevel = ParseOutlineLevel(request.OutlineLevel);
            }
        }

        private static void RequireLineSpacingValue(SetParagraphFormatRequest request)
        {
            if (!request.LineSpacingValue.HasValue)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams,
                    $"行距模式 {request.LineSpacing} 需要提供 line_spacing_value");
            }
        }

        private static void ApplyFont(Range range, SetTextStyleRequest request)
        {
            var font = range.Font;
            if (!string.IsNullOrEmpty(request.FontName))
            {
                font.Name = request.FontName;
                font.NameAscii = request.FontName;
                font.NameOther = request.FontName;
                try
                {
                    // NameFarEast 只接受中文字体名（传 Arial 等西文字体会抛 DISP_E_BADPARAMCOUNT）
                    font.NameFarEast = request.FontName;
                }
                catch
                {
                    // 非中文字体忽略（Name/NameAscii 已生效）
                }
            }

            if (request.FontSize.HasValue)
            {
                if (request.FontSize.Value < 1 || request.FontSize.Value > 400)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "font_size 必须在 1-400 之间");
                }

                font.Size = (float)request.FontSize.Value;
            }

            if (request.Bold.HasValue) font.Bold = request.Bold.Value ? -1 : 0;
            if (request.Italic.HasValue) font.Italic = request.Italic.Value ? -1 : 0;
            if (request.Underline.HasValue)
            {
                font.Underline = request.Underline.Value ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;
            }

            if (request.ColorArgb.HasValue) font.Color = ArgbToWdColor(request.ColorArgb.Value);
            if (!string.IsNullOrEmpty(request.Highlight)) range.HighlightColorIndex = ParseHighlight(request.Highlight);
            if (request.AllCaps.HasValue) font.AllCaps = request.AllCaps.Value ? -1 : 0;
            if (request.Subscript.HasValue) font.Subscript = request.Subscript.Value ? -1 : 0;
            if (request.Superscript.HasValue) font.Superscript = request.Superscript.Value ? -1 : 0;
        }

        private static int ApplyFontToMatches(Document doc, string findText, SetTextStyleRequest request)
        {
            var range = doc.Content;
            var find = range.Find;
            find.ClearFormatting();
            find.Text = findText;
            find.MatchCase = false;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;

            object missing = Type.Missing;
            object findTextObj = findText;
            var count = 0;
            while (count < 1000 && find.Execute(ref findTextObj, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing))
            {
                ApplyFont(range, request);
                count++;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
            }

            return count;
        }

        private static int ApplyStyleToMatches(Document doc, string findText, object styleObj)
        {
            var range = doc.Content;
            var find = range.Find;
            find.ClearFormatting();
            find.Text = findText;
            find.MatchCase = false;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;

            object missing = Type.Missing;
            object findTextObj = findText;
            var count = 0;
            while (count < 1000 && find.Execute(ref findTextObj, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing))
            {
                range.set_Style(ref styleObj);
                count++;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
            }

            return count;
        }

        private static int FindParagraphIndexByText(Document doc, string findText)
        {
            var count = doc.Paragraphs.Count;
            for (var i = 1; i <= count; i++)
            {
                Paragraph p = null;
                try
                {
                    p = doc.Paragraphs[i];
                    var text = p.Range.Text;
                    if (text != null && text.IndexOf(findText, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return i;
                    }
                }
                catch
                {
                    // 忽略单个段落读取失败
                }
                finally
                {
                    TryRelease(p);
                }
            }

            return -1;
        }

        private static Range FindFirstRange(Document doc, string text)
        {
            var range = doc.Content;
            var find = range.Find;
            find.ClearFormatting();
            find.Text = text;
            find.MatchCase = false;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;

            object missing = Type.Missing;
            object findTextObj = text;
            return find.Execute(ref findTextObj, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing)
                ? range
                : null;
        }

        private static int GetParagraphIndexAt(Document doc, int position)
        {
            try
            {
                if (position <= 0)
                {
                    return 1;
                }

                var contentEnd = doc.Content.End;
                if (position >= contentEnd)
                {
                    return doc.Paragraphs.Count;
                }

                object start = position;
                object end = Math.Min(position + 1, contentEnd);
                var range = doc.Range(ref start, ref end);
                var paragraphStart = range.Paragraphs[1].Range.Start;
                if (paragraphStart <= 0)
                {
                    return 1;
                }

                object zero = 0;
                object before = paragraphStart - 1;
                return doc.Range(ref zero, ref before).Paragraphs.Count + 1;
            }
            catch
            {
                return 0;
            }
        }

        private static int GetOutlineLevel(Paragraph paragraph)
        {
            try
            {
                switch (paragraph.OutlineLevel)
                {
                    case WdOutlineLevel.wdOutlineLevel1: return 1;
                    case WdOutlineLevel.wdOutlineLevel2: return 2;
                    case WdOutlineLevel.wdOutlineLevel3: return 3;
                    case WdOutlineLevel.wdOutlineLevel4: return 4;
                    case WdOutlineLevel.wdOutlineLevel5: return 5;
                    case WdOutlineLevel.wdOutlineLevel6: return 6;
                    case WdOutlineLevel.wdOutlineLevel7: return 7;
                    case WdOutlineLevel.wdOutlineLevel8: return 8;
                    case WdOutlineLevel.wdOutlineLevel9: return 9;
                    default: return 0;
                }
            }
            catch
            {
                return 0;
            }
        }

        private static DocumentInfo DescribeDocumentInfo(Document doc)
        {
            object missing = Type.Missing;
            return new DocumentInfo
            {
                Name = SafeRead(() => doc.Name),
                Path = SafeRead(() => doc.Path),
                FullName = SafeRead(() => doc.FullName),
                Saved = SafeRead(() => doc.Saved),
                ReadOnly = SafeRead(() => doc.ReadOnly),
                ParagraphCount = SafeRead(() => doc.Paragraphs.Count),
                TableCount = SafeRead(() => doc.Tables.Count),
                SectionCount = SafeRead(() => doc.Sections.Count),
                PageCount = SafeRead(() => doc.ComputeStatistics(WdStatistic.wdStatisticPages, ref missing)),
                WordCount = SafeRead(() => doc.ComputeStatistics(WdStatistic.wdStatisticWords, ref missing)),
                CharacterCount = SafeRead(() => doc.ComputeStatistics(WdStatistic.wdStatisticCharacters, ref missing)),
            };
        }

        private static HeaderFooterInfo DescribeHeaderFooter(Document doc, int sectionIndex, string headerType)
        {
            var section = doc.Sections[sectionIndex];
            var headerIndex = ParseHeaderFooterType(headerType);
            var header = section.Headers[headerIndex];
            var footer = section.Footers[headerIndex];
            return new HeaderFooterInfo
            {
                SectionIndex = sectionIndex,
                HeaderType = NormalizeHeaderTypeName(headerType),
                HeaderText = TrimParagraphText(SafeRead(() => header.Range.Text)),
                FooterText = TrimParagraphText(SafeRead(() => footer.Range.Text)),
                LinkToPrevious = SafeRead(() => (bool?)header.LinkToPrevious),
                DifferentFirstPage = SafeRead(() => section.PageSetup.DifferentFirstPageHeaderFooter == -1),
                OddEvenPages = SafeRead(() => section.PageSetup.OddAndEvenPagesHeaderFooter == -1),
                SectionCount = doc.Sections.Count,
            };
        }

        private static PageSetupInfo DescribePageSetup(Document doc, int sectionIndex)
        {
            var setup = doc.Sections[sectionIndex].PageSetup;
            return new PageSetupInfo
            {
                SectionIndex = sectionIndex,
                PageWidth = SafeRead(() => (double)setup.PageWidth),
                PageHeight = SafeRead(() => (double)setup.PageHeight),
                Orientation = SafeRead(() => DescribeOrientation(setup.Orientation)),
                PaperSize = SafeRead(() => DescribePaperSize(setup.PaperSize)),
                TopMargin = SafeRead(() => (double)setup.TopMargin),
                BottomMargin = SafeRead(() => (double)setup.BottomMargin),
                LeftMargin = SafeRead(() => (double)setup.LeftMargin),
                RightMargin = SafeRead(() => (double)setup.RightMargin),
                HeaderDistance = SafeRead(() => (double)setup.HeaderDistance),
                FooterDistance = SafeRead(() => (double)setup.FooterDistance),
                Gutter = SafeRead(() => (double)setup.Gutter),
                TextColumns = SafeRead(() => setup.TextColumns.Count),
                SectionCount = doc.Sections.Count,
            };
        }

        private static TableInfo DescribeTable(Document doc, int tableIndex)
        {
            var table = doc.Tables[tableIndex];
            var widths = new List<double>();
            try
            {
                for (var c = 1; c <= table.Columns.Count; c++)
                {
                    Column column = null;
                    try
                    {
                        column = table.Columns[c];
                        widths.Add((double)column.Width);
                    }
                    catch
                    {
                        widths.Add(0);
                    }
                    finally
                    {
                        TryRelease(column);
                    }
                }
            }
            catch
            {
                // 列宽读取失败时忽略
            }

            return new TableInfo
            {
                TableIndex = tableIndex,
                Rows = SafeRead(() => table.Rows.Count),
                Columns = SafeRead(() => table.Columns.Count),
                Style = SafeRead(() => ((Style)table.get_Style())?.NameLocal),
                Alignment = SafeRead(() => DescribeRowAlignment(table.Rows.Alignment)),
                TableCount = doc.Tables.Count,
                ColumnWidths = widths,
            };
        }

        private static TableData ReadTableData(Table table, int tableIndex)
        {
            var data = new TableData
            {
                TableIndex = tableIndex,
                Rows = table.Rows.Count,
                Columns = table.Columns.Count,
            };

            for (var r = 1; r <= data.Rows; r++)
            {
                var rowData = new List<string>();
                for (var c = 1; c <= data.Columns; c++)
                {
                    Cell cell = null;
                    try
                    {
                        cell = table.Cell(r, c);
                        rowData.Add(TrimParagraphText(cell.Range.Text));
                    }
                    catch
                    {
                        // 合并单元格的覆盖位置无法访问，返回空串
                        rowData.Add(string.Empty);
                    }
                    finally
                    {
                        TryRelease(cell);
                    }
                }

                data.Data.Add(rowData);
            }

            return data;
        }

        private static void FillTable(Table table, List<List<string>> data)
        {
            for (var r = 0; r < data.Count && r < table.Rows.Count; r++)
            {
                var row = data[r];
                if (row == null)
                {
                    continue;
                }

                for (var c = 0; c < row.Count && c < table.Columns.Count; c++)
                {
                    SetCellText(table, r + 1, c + 1, row[c]);
                }
            }
        }

        private static void SetCellText(Table table, int row, int column, string value)
        {
            try
            {
                table.Cell(row, column).Range.Text = NormalizeText(value) ?? string.Empty;
            }
            catch (Exception ex) when (!(ex is McpToolException))
            {
                throw new McpToolException(McpErrorCodes.TableCellOutOfRange,
                    $"写入单元格失败（第 {row} 行第 {column} 列）：{ex.Message}");
            }
        }

        private static void ApplyCellStyle(Cell cell, SetTableCellStyleRequest request)
        {
            var range = cell.Range;
            var font = range.Font;
            if (!string.IsNullOrEmpty(request.FontName))
            {
                font.Name = request.FontName;
                font.NameAscii = request.FontName;
                font.NameOther = request.FontName;
                try
                {
                    // NameFarEast 只接受中文字体名（传 Arial 等西文字体会抛 DISP_E_BADPARAMCOUNT）
                    font.NameFarEast = request.FontName;
                }
                catch
                {
                    // 非中文字体忽略（Name/NameAscii 已生效）
                }
            }

            if (request.FontSize.HasValue) font.Size = (float)request.FontSize.Value;
            if (request.Bold.HasValue) font.Bold = request.Bold.Value ? -1 : 0;
            if (request.Italic.HasValue) font.Italic = request.Italic.Value ? -1 : 0;
            if (request.ColorArgb.HasValue) font.Color = ArgbToWdColor(request.ColorArgb.Value);

            if (!string.IsNullOrEmpty(request.Alignment))
            {
                range.ParagraphFormat.Alignment = ParseAlignment(request.Alignment);
            }

            if (request.ShadingArgb.HasValue)
            {
                cell.Shading.BackgroundPatternColor = ArgbToWdColor(request.ShadingArgb.Value);
            }

            if (!string.IsNullOrEmpty(request.VerticalAlignment))
            {
                cell.VerticalAlignment = ParseCellVerticalAlignment(request.VerticalAlignment);
            }
        }

        private static void ValidateCellRange(Table table, int row, int column, string rowName, string columnName)
        {
            var rowCount = table.Rows.Count;
            var columnCount = table.Columns.Count;
            if (row < 1 || row > rowCount)
            {
                throw new McpToolException(McpErrorCodes.TableCellOutOfRange,
                    $"{rowName}={row} 超出范围（当前共 {rowCount} 行）");
            }

            if (column < 1 || column > columnCount)
            {
                throw new McpToolException(McpErrorCodes.TableCellOutOfRange,
                    $"{columnName}={column} 超出范围（当前共 {columnCount} 列）");
            }
        }

        private static int GetTableIndex(Document doc, Table table)
        {
            try
            {
                var start = table.Range.Start;
                for (var i = 1; i <= doc.Tables.Count; i++)
                {
                    Table candidate = null;
                    try
                    {
                        candidate = doc.Tables[i];
                        if (candidate.Range.Start == start)
                        {
                            return i;
                        }
                    }
                    finally
                    {
                        TryRelease(candidate);
                    }
                }
            }
            catch
            {
                // 忽略
            }

            return doc.Tables.Count;
        }

        private static void ApplyTableStyle(Document doc, Table table, string styleName)
        {
            object styleObj = ResolveStyleValue(doc, styleName);
            try
            {
                table.set_Style(ref styleObj);
            }
            catch (Exception ex)
            {
                throw new McpToolException(McpErrorCodes.StyleNotFound, $"应用表格样式失败：{ex.Message}");
            }
        }

        private static int FindCommentIndex(Document doc, Comment comment)
        {
            try
            {
                var start = comment.Scope.Start;
                for (var i = 1; i <= doc.Comments.Count; i++)
                {
                    if (doc.Comments[i].Scope.Start == start)
                    {
                        return i;
                    }
                }
            }
            catch
            {
                // 忽略
            }

            return doc.Comments.Count;
        }

        private static CommentInfo DescribeComment(Document doc, int index)
        {
            var comment = doc.Comments[index];
            return new CommentInfo
            {
                Index = index,
                Author = SafeRead(() => comment.Author),
                Initial = SafeRead(() => comment.Initial),
                Text = SafeRead(() => TrimParagraphText(comment.Range.Text)),
                ScopeText = SafeRead(() => TrimParagraphText(comment.Scope.Text)),
                Done = SafeRead(() => comment.Done),
                ReplyCount = SafeRead(() => comment.Replies.Count),
            };
        }

        /// <summary>
        /// 解析样式：内置别名映射到 WdBuiltinStyle（与界面语言无关），
        /// 其他值按样式名（支持本地化名称）查找。
        /// </summary>
        private static object ResolveStyleValue(Document doc, string style)
        {
            var key = style.Trim().ToLowerInvariant().Replace("_", "-").Replace(" ", "-");
            switch (key)
            {
                case "normal":
                case "body":
                case "正文": return WdBuiltinStyle.wdStyleNormal;
                case "heading1":
                case "heading-1":
                case "h1":
                case "标题1":
                case "标题-1": return WdBuiltinStyle.wdStyleHeading1;
                case "heading2":
                case "heading-2":
                case "h2":
                case "标题2":
                case "标题-2": return WdBuiltinStyle.wdStyleHeading2;
                case "heading3":
                case "heading-3":
                case "h3":
                case "标题3":
                case "标题-3": return WdBuiltinStyle.wdStyleHeading3;
                case "heading4":
                case "heading-4":
                case "h4":
                case "标题4":
                case "标题-4": return WdBuiltinStyle.wdStyleHeading4;
                case "heading5":
                case "heading-5":
                case "h5":
                case "标题5":
                case "标题-5": return WdBuiltinStyle.wdStyleHeading5;
                case "heading6":
                case "heading-6":
                case "h6":
                case "标题6":
                case "标题-6": return WdBuiltinStyle.wdStyleHeading6;
                case "heading7":
                case "heading-7":
                case "h7":
                case "标题7":
                case "标题-7": return WdBuiltinStyle.wdStyleHeading7;
                case "heading8":
                case "heading-8":
                case "h8":
                case "标题8":
                case "标题-8": return WdBuiltinStyle.wdStyleHeading8;
                case "heading9":
                case "heading-9":
                case "h9":
                case "标题9":
                case "标题-9": return WdBuiltinStyle.wdStyleHeading9;
                case "title":
                case "标题": return WdBuiltinStyle.wdStyleTitle;
                case "subtitle":
                case "副标题": return WdBuiltinStyle.wdStyleSubtitle;
                case "quote":
                case "引用": return WdBuiltinStyle.wdStyleQuote;
                case "intense-quote":
                case "intensequote":
                case "明显引用": return WdBuiltinStyle.wdStyleIntenseQuote;
                case "list-paragraph":
                case "listparagraph":
                case "列表段落": return WdBuiltinStyle.wdStyleListParagraph;
                case "caption":
                case "题注": return WdBuiltinStyle.wdStyleCaption;
                case "strong":
                case "加粗": return WdBuiltinStyle.wdStyleStrong;
                case "emphasis":
                case "强调": return WdBuiltinStyle.wdStyleEmphasis;
            }

            // 按名称查找（先精确匹配，再忽略大小写匹配本地化名称）
            try
            {
                var styleObject = doc.Styles[style];
                if (styleObject != null)
                {
                    return styleObject;
                }
            }
            catch
            {
                // 继续模糊匹配
            }

            for (var i = 1; i <= doc.Styles.Count; i++)
            {
                Style candidate = null;
                try
                {
                    candidate = doc.Styles[i];
                    if (string.Equals(candidate.NameLocal, style, StringComparison.OrdinalIgnoreCase))
                    {
                        return candidate;
                    }
                }
                catch
                {
                    // 忽略单个样式读取失败
                }
            }

            throw new McpToolException(McpErrorCodes.StyleNotFound,
                $"找不到样式：{style}（可用 list_styles 查看当前文档样式）");
        }

        private static WdBuiltinStyle HeadingStyle(int level)
        {
            switch (level)
            {
                case 1: return WdBuiltinStyle.wdStyleHeading1;
                case 2: return WdBuiltinStyle.wdStyleHeading2;
                case 3: return WdBuiltinStyle.wdStyleHeading3;
                case 4: return WdBuiltinStyle.wdStyleHeading4;
                case 5: return WdBuiltinStyle.wdStyleHeading5;
                case 6: return WdBuiltinStyle.wdStyleHeading6;
                case 7: return WdBuiltinStyle.wdStyleHeading7;
                case 8: return WdBuiltinStyle.wdStyleHeading8;
                default: return WdBuiltinStyle.wdStyleHeading9;
            }
        }

        private static WdParagraphAlignment ParseAlignment(string alignment)
        {
            switch (alignment.Trim().ToLowerInvariant())
            {
                case "left": return WdParagraphAlignment.wdAlignParagraphLeft;
                case "center":
                case "centre": return WdParagraphAlignment.wdAlignParagraphCenter;
                case "right": return WdParagraphAlignment.wdAlignParagraphRight;
                case "justify":
                case "both": return WdParagraphAlignment.wdAlignParagraphJustify;
                case "distribute": return WdParagraphAlignment.wdAlignParagraphDistribute;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的对齐方式：{alignment}（支持 left / center / right / justify / distribute）");
            }
        }

        private static string DescribeAlignment(WdParagraphAlignment alignment)
        {
            switch (alignment)
            {
                case WdParagraphAlignment.wdAlignParagraphLeft: return "left";
                case WdParagraphAlignment.wdAlignParagraphCenter: return "center";
                case WdParagraphAlignment.wdAlignParagraphRight: return "right";
                case WdParagraphAlignment.wdAlignParagraphJustify: return "justify";
                case WdParagraphAlignment.wdAlignParagraphDistribute: return "distribute";
                default: return alignment.ToString();
            }
        }

        private static WdRowAlignment ParseRowAlignment(string alignment)
        {
            switch (alignment.Trim().ToLowerInvariant())
            {
                case "left": return WdRowAlignment.wdAlignRowLeft;
                case "center":
                case "centre": return WdRowAlignment.wdAlignRowCenter;
                case "right": return WdRowAlignment.wdAlignRowRight;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的表格对齐方式：{alignment}（支持 left / center / right）");
            }
        }

        private static string DescribeRowAlignment(WdRowAlignment alignment)
        {
            switch (alignment)
            {
                case WdRowAlignment.wdAlignRowLeft: return "left";
                case WdRowAlignment.wdAlignRowCenter: return "center";
                case WdRowAlignment.wdAlignRowRight: return "right";
                default: return alignment.ToString();
            }
        }

        private static WdCellVerticalAlignment ParseCellVerticalAlignment(string alignment)
        {
            switch (alignment.Trim().ToLowerInvariant())
            {
                case "top": return WdCellVerticalAlignment.wdCellAlignVerticalTop;
                case "center":
                case "centre":
                case "middle": return WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                case "bottom": return WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的垂直对齐方式：{alignment}（支持 top / center / bottom）");
            }
        }

        private static WdAutoFitBehavior ParseAutoFit(string autoFit)
        {
            switch (autoFit.Trim().ToLowerInvariant())
            {
                case "fixed": return WdAutoFitBehavior.wdAutoFitFixed;
                case "content": return WdAutoFitBehavior.wdAutoFitContent;
                case "window": return WdAutoFitBehavior.wdAutoFitWindow;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的自动调整方式：{autoFit}（支持 fixed / content / window）");
            }
        }

        private static WdOutlineLevel ParseOutlineLevel(string outlineLevel)
        {
            switch (outlineLevel.Trim().ToLowerInvariant())
            {
                case "body":
                case "body-text":
                case "text": return WdOutlineLevel.wdOutlineLevelBodyText;
                case "level-1":
                case "1": return WdOutlineLevel.wdOutlineLevel1;
                case "level-2":
                case "2": return WdOutlineLevel.wdOutlineLevel2;
                case "level-3":
                case "3": return WdOutlineLevel.wdOutlineLevel3;
                case "level-4":
                case "4": return WdOutlineLevel.wdOutlineLevel4;
                case "level-5":
                case "5": return WdOutlineLevel.wdOutlineLevel5;
                case "level-6":
                case "6": return WdOutlineLevel.wdOutlineLevel6;
                case "level-7":
                case "7": return WdOutlineLevel.wdOutlineLevel7;
                case "level-8":
                case "8": return WdOutlineLevel.wdOutlineLevel8;
                case "level-9":
                case "9": return WdOutlineLevel.wdOutlineLevel9;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的大纲级别：{outlineLevel}（支持 body-text / level-1 .. level-9）");
            }
        }

        private static WdOrientation ParseOrientation(string orientation)
        {
            switch (orientation.Trim().ToLowerInvariant())
            {
                case "portrait":
                case "纵向": return WdOrientation.wdOrientPortrait;
                case "landscape":
                case "横向": return WdOrientation.wdOrientLandscape;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的页面方向：{orientation}（支持 portrait / landscape）");
            }
        }

        private static string DescribeOrientation(WdOrientation orientation)
        {
            switch (orientation)
            {
                case WdOrientation.wdOrientPortrait: return "portrait";
                case WdOrientation.wdOrientLandscape: return "landscape";
                default: return orientation.ToString();
            }
        }

        private static WdPaperSize ParsePaperSize(string paperSize)
        {
            switch (paperSize.Trim().ToLowerInvariant())
            {
                case "a3": return WdPaperSize.wdPaperA3;
                case "a4": return WdPaperSize.wdPaperA4;
                case "a5": return WdPaperSize.wdPaperA5;
                case "b4": return WdPaperSize.wdPaperB4;
                case "b5": return WdPaperSize.wdPaperB5;
                case "letter": return WdPaperSize.wdPaperLetter;
                case "legal": return WdPaperSize.wdPaperLegal;
                case "executive": return WdPaperSize.wdPaperExecutive;
                case "tabloid": return WdPaperSize.wdPaperTabloid;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的纸张大小：{paperSize}（支持 A3 / A4 / A5 / B4 / B5 / Letter / Legal / Executive / Tabloid）");
            }
        }

        private static string DescribePaperSize(WdPaperSize paperSize)
        {
            switch (paperSize)
            {
                case WdPaperSize.wdPaperA3: return "A3";
                case WdPaperSize.wdPaperA4: return "A4";
                case WdPaperSize.wdPaperA5: return "A5";
                case WdPaperSize.wdPaperB4: return "B4";
                case WdPaperSize.wdPaperB5: return "B5";
                case WdPaperSize.wdPaperLetter: return "Letter";
                case WdPaperSize.wdPaperLegal: return "Legal";
                case WdPaperSize.wdPaperExecutive: return "Executive";
                case WdPaperSize.wdPaperTabloid: return "Tabloid";
                default: return paperSize.ToString();
            }
        }

        private static string DescribeStyleType(WdStyleType type)
        {
            switch (type)
            {
                case WdStyleType.wdStyleTypeParagraph: return "paragraph";
                case WdStyleType.wdStyleTypeCharacter: return "character";
                case WdStyleType.wdStyleTypeTable: return "table";
                case WdStyleType.wdStyleTypeList: return "list";
                default: return type.ToString();
            }
        }

        private static WdSaveFormat ParseSaveFormat(string format, string path)
        {
            var key = (format ?? Path.GetExtension(path) ?? string.Empty).TrimStart('.').ToLowerInvariant();
            switch (key)
            {
                case "":
                case "docx":
                case "documentdefault": return WdSaveFormat.wdFormatDocumentDefault;
                case "docm": return WdSaveFormat.wdFormatXMLDocumentMacroEnabled;
                case "xml":
                case "wordxml": return WdSaveFormat.wdFormatXML;
                case "doc":
                case "doc97": return WdSaveFormat.wdFormatDocument;
                case "rtf": return WdSaveFormat.wdFormatRTF;
                case "txt":
                case "text": return WdSaveFormat.wdFormatText;
                case "odt": return WdSaveFormat.wdFormatOpenDocumentText;
                case "html":
                case "htm": return WdSaveFormat.wdFormatHTML;
                case "mhtml":
                case "mht": return WdSaveFormat.wdFormatWebArchive;
                case "pdf": throw new McpToolException(McpErrorCodes.UnsupportedType, "保存为 PDF 请使用 export_to_pdf 工具");
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的保存格式：{key}（支持 docx / docm / doc / rtf / txt / odt / html / mhtml / xml）");
            }
        }

        /// <summary>ARGB（0xAARRGGBB）转 WdColor（BGR 顺序：0x00BBGGRR）。</summary>
        private static WdColor ArgbToWdColor(int argb) => (WdColor)ArgbToOleRgb(argb);

        /// <summary>ARGB（0xAARRGGBB）转 OLE 颜色整型（BGR 顺序：0x00BBGGRR）。</summary>
        private static int ArgbToOleRgb(int argb)
        {
            var r = (argb >> 16) & 0xFF;
            var g = (argb >> 8) & 0xFF;
            var b = argb & 0xFF;
            return r | (g << 8) | (b << 16);
        }

        /// <summary>WdColor 转 #RRGGBB；混合值（wdUndefined）返回 null。</summary>
        private static string WdColorToHex(WdColor color)
        {
            var value = (int)color;
            if (value == WdUndefinedValue)
            {
                return null;
            }

            var r = value & 0xFF;
            var g = (value >> 8) & 0xFF;
            var b = (value >> 16) & 0xFF;
            return $"#{r:X2}{g:X2}{b:X2}";
        }

        /// <summary>Word 三态 Int32（-1/0/9999999）转可空布尔。</summary>
        private static bool? ToNullableBool(int value)
        {
            if (value == -1)
            {
                return true;
            }

            if (value == 0)
            {
                return false;
            }

            return null;
        }

        private static ParagraphDetail DescribeParagraphDetail(Document doc, Paragraph p, int index)
        {
            var detail = new ParagraphDetail
            {
                ParagraphIndex = index,
                Text = TrimParagraphText(SafeRead(() => p.Range.Text)),
                Style = SafeRead(() => ((Style)p.Range.get_Style())?.NameLocal),
                Alignment = SafeRead(() => DescribeAlignment(p.Alignment)),
                OutlineLevel = SafeRead(() => DescribeOutlineLevel(p.OutlineLevel)),
                InTable = SafeRead(() => (bool)p.Range.Information[WdInformation.wdWithInTable]),
                ParagraphCount = doc.Paragraphs.Count,
            };

            var format = p.Range.ParagraphFormat;
            detail.LeftIndent = SafeRead(() => (double)format.LeftIndent);
            detail.RightIndent = SafeRead(() => (double)format.RightIndent);
            detail.FirstLineIndent = SafeRead(() => (double)format.FirstLineIndent);
            detail.SpaceBefore = SafeRead(() => (double)format.SpaceBefore);
            detail.SpaceAfter = SafeRead(() => (double)format.SpaceAfter);
            detail.LineSpacingRule = SafeRead(() => DescribeLineSpacingRule(format.LineSpacingRule));
            detail.LineSpacing = SafeRead(() => (double)format.LineSpacing);
            detail.KeepWithNext = SafeRead(() => (bool?)(p.KeepWithNext == -1));
            detail.KeepTogether = SafeRead(() => (bool?)(p.KeepTogether == -1));
            detail.PageBreakBefore = SafeRead(() => (bool?)(p.PageBreakBefore == -1));

            var font = p.Range.Font;
            detail.FontName = SafeRead(() => font.Name);
            detail.FontSize = SafeRead(() => (double)font.Size);
            detail.Bold = SafeRead(() => ToNullableBool(font.Bold));
            detail.Italic = SafeRead(() => ToNullableBool(font.Italic));
            detail.Underline = SafeRead(() => font.Underline != WdUnderline.wdUnderlineNone);
            detail.FontColor = SafeRead(() => WdColorToHex(font.Color));
            detail.Highlight = SafeRead(() => DescribeHighlight(p.Range.HighlightColorIndex));
            detail.ListType = SafeRead(() => DescribeListType(p.Range.ListFormat.ListType));
            detail.ListLevel = SafeRead(() => p.Range.ListFormat.ListLevelNumber);
            return detail;
        }

        private static string DescribeOutlineLevel(WdOutlineLevel level)
        {
            switch (level)
            {
                case WdOutlineLevel.wdOutlineLevel1: return "level-1";
                case WdOutlineLevel.wdOutlineLevel2: return "level-2";
                case WdOutlineLevel.wdOutlineLevel3: return "level-3";
                case WdOutlineLevel.wdOutlineLevel4: return "level-4";
                case WdOutlineLevel.wdOutlineLevel5: return "level-5";
                case WdOutlineLevel.wdOutlineLevel6: return "level-6";
                case WdOutlineLevel.wdOutlineLevel7: return "level-7";
                case WdOutlineLevel.wdOutlineLevel8: return "level-8";
                case WdOutlineLevel.wdOutlineLevel9: return "level-9";
                case WdOutlineLevel.wdOutlineLevelBodyText: return "body-text";
                default: return level.ToString();
            }
        }

        private static string DescribeLineSpacingRule(WdLineSpacing rule)
        {
            switch (rule)
            {
                case WdLineSpacing.wdLineSpaceSingle: return "single";
                case WdLineSpacing.wdLineSpace1pt5: return "1.5";
                case WdLineSpacing.wdLineSpaceDouble: return "double";
                case WdLineSpacing.wdLineSpaceAtLeast: return "at-least";
                case WdLineSpacing.wdLineSpaceExactly: return "exactly";
                case WdLineSpacing.wdLineSpaceMultiple: return "multiple";
                default: return rule.ToString();
            }
        }

        private static string DescribeHighlight(WdColorIndex index)
        {
            if ((int)index == WdUndefinedValue)
            {
                return "mixed";
            }

            switch (index)
            {
                case WdColorIndex.wdAuto: return "none";
                case WdColorIndex.wdYellow: return "yellow";
                case WdColorIndex.wdBrightGreen: return "bright-green";
                case WdColorIndex.wdTurquoise: return "turquoise";
                case WdColorIndex.wdPink: return "pink";
                case WdColorIndex.wdBlue: return "blue";
                case WdColorIndex.wdRed: return "red";
                case WdColorIndex.wdDarkBlue: return "dark-blue";
                case WdColorIndex.wdTeal: return "teal";
                case WdColorIndex.wdGreen: return "green";
                case WdColorIndex.wdViolet: return "violet";
                case WdColorIndex.wdDarkRed: return "dark-red";
                case WdColorIndex.wdDarkYellow: return "dark-yellow";
                case WdColorIndex.wdGray50: return "gray";
                case WdColorIndex.wdGray25: return "gray-25";
                case WdColorIndex.wdBlack: return "black";
                case WdColorIndex.wdWhite: return "white";
                default: return index.ToString();
            }
        }

        private static WdColorIndex ParseHighlight(string name)
        {
            switch (name.Trim().ToLowerInvariant())
            {
                case "none":
                case "auto": return WdColorIndex.wdAuto;
                case "yellow": return WdColorIndex.wdYellow;
                case "bright-green":
                case "green": return WdColorIndex.wdBrightGreen;
                case "turquoise":
                case "cyan": return WdColorIndex.wdTurquoise;
                case "pink": return WdColorIndex.wdPink;
                case "blue": return WdColorIndex.wdBlue;
                case "red": return WdColorIndex.wdRed;
                case "dark-blue": return WdColorIndex.wdDarkBlue;
                case "teal": return WdColorIndex.wdTeal;
                case "dark-green": return WdColorIndex.wdGreen;
                case "violet": return WdColorIndex.wdViolet;
                case "dark-red": return WdColorIndex.wdDarkRed;
                case "dark-yellow": return WdColorIndex.wdDarkYellow;
                case "gray":
                case "gray-50": return WdColorIndex.wdGray50;
                case "gray-25": return WdColorIndex.wdGray25;
                case "black": return WdColorIndex.wdBlack;
                case "white": return WdColorIndex.wdWhite;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的突出显示颜色：{name}（支持 yellow / bright-green / turquoise / pink / blue / red / dark-blue / teal / violet / dark-red / dark-yellow / gray / gray-25 / black / white / none）");
            }
        }

        private static string DescribeListType(WdListType type)
        {
            switch (type)
            {
                case WdListType.wdListNoNumbering: return "none";
                case WdListType.wdListBullet: return "bullet";
                case WdListType.wdListSimpleNumbering: return "number";
                case WdListType.wdListOutlineNumbering: return "outline";
                case WdListType.wdListMixedNumbering: return "mixed";
                case WdListType.wdListListNumOnly: return "list-num";
                case WdListType.wdListPictureBullet: return "picture-bullet";
                default: return type.ToString();
            }
        }

        /// <summary>解析列表目标范围：优先 start/end 段落区间，其次单段落，缺省最后一个段落。</summary>
        private static Range ResolveListRange(Document doc, SetListFormatRequest request)
        {
            if (request.StartParagraph.HasValue || request.EndParagraph.HasValue)
            {
                var start = request.StartParagraph ?? request.EndParagraph.Value;
                var end = request.EndParagraph ?? request.StartParagraph.Value;
                if (start > end)
                {
                    var temp = start;
                    start = end;
                    end = temp;
                }

                GetParagraphOrThrow(doc, start);
                GetParagraphOrThrow(doc, end);
                object startObj = doc.Paragraphs[start].Range.Start;
                object endObj = doc.Paragraphs[end].Range.End;
                return doc.Range(ref startObj, ref endObj);
            }

            var index = request.ParagraphIndex ?? doc.Paragraphs.Count;
            return GetParagraphOrThrow(doc, index).Range;
        }

        private static void ApplyRunFont(Range range, TextRunSpec run)
        {
            var font = range.Font;
            if (!string.IsNullOrEmpty(run.FontName))
            {
                font.Name = run.FontName;
                font.NameAscii = run.FontName;
                font.NameOther = run.FontName;
                try
                {
                    font.NameFarEast = run.FontName;
                }
                catch
                {
                    // 非中文字体忽略
                }
            }

            if (run.FontSize.HasValue)
            {
                if (run.FontSize.Value < 1 || run.FontSize.Value > 400)
                {
                    throw new McpToolException(McpErrorCodes.InvalidParams, "runs[].font_size 必须在 1-400 之间");
                }

                font.Size = (float)run.FontSize.Value;
            }

            if (run.Bold.HasValue) font.Bold = run.Bold.Value ? -1 : 0;
            if (run.Italic.HasValue) font.Italic = run.Italic.Value ? -1 : 0;
            if (run.Underline.HasValue)
            {
                font.Underline = run.Underline.Value ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;
            }

            if (run.ColorArgb.HasValue) font.Color = ArgbToWdColor(run.ColorArgb.Value);
            if (!string.IsNullOrEmpty(run.Highlight)) range.HighlightColorIndex = ParseHighlight(run.Highlight);
            if (run.Superscript.HasValue) font.Superscript = run.Superscript.Value ? -1 : 0;
            if (run.Subscript.HasValue) font.Subscript = run.Subscript.Value ? -1 : 0;
        }

        private static WdStyleType ParseStyleType(string type)
        {
            switch ((type ?? "paragraph").Trim().ToLowerInvariant())
            {
                case "paragraph": return WdStyleType.wdStyleTypeParagraph;
                case "character": return WdStyleType.wdStyleTypeCharacter;
                case "table": return WdStyleType.wdStyleTypeTable;
                case "list": return WdStyleType.wdStyleTypeList;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的样式类型：{type}（支持 paragraph / character / table / list）");
            }
        }

        private static void ApplyStyleFormatting(Style style, StyleFormatRequest request)
        {
            var font = style.Font;
            if (!string.IsNullOrEmpty(request.FontName))
            {
                font.Name = request.FontName;
                font.NameAscii = request.FontName;
                font.NameOther = request.FontName;
                try
                {
                    font.NameFarEast = request.FontName;
                }
                catch
                {
                    // 非中文字体忽略
                }
            }

            if (request.FontSize.HasValue) font.Size = (float)request.FontSize.Value;
            if (request.Bold.HasValue) font.Bold = request.Bold.Value ? -1 : 0;
            if (request.Italic.HasValue) font.Italic = request.Italic.Value ? -1 : 0;
            if (request.ColorArgb.HasValue) font.Color = ArgbToWdColor(request.ColorArgb.Value);

            var format = style.ParagraphFormat;
            if (!string.IsNullOrEmpty(request.Alignment)) format.Alignment = ParseAlignment(request.Alignment);
            if (request.SpaceBefore.HasValue) format.SpaceBefore = (float)request.SpaceBefore.Value;
            if (request.SpaceAfter.HasValue) format.SpaceAfter = (float)request.SpaceAfter.Value;
            if (!string.IsNullOrEmpty(request.LineSpacing))
            {
                switch (request.LineSpacing.Trim().ToLowerInvariant())
                {
                    case "single": format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle; break;
                    case "1.5": format.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5; break;
                    case "double": format.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble; break;
                    default:
                        throw new McpToolException(McpErrorCodes.UnsupportedType,
                            $"不支持的行距：{request.LineSpacing}（支持 single / 1.5 / double）");
                }
            }
        }

        private static StyleInfo DescribeStyle(Style style)
        {
            return new StyleInfo
            {
                Name = SafeRead(() => style.NameLocal),
                Type = SafeRead(() => DescribeStyleType(style.Type)),
                BuiltIn = SafeRead(() => style.BuiltIn),
                InUse = SafeRead(() => style.InUse),
            };
        }

        /// <summary>按名称查找样式（先精确、再忽略大小写匹配本地化名称）；找不到返回 null。</summary>
        private static Style FindStyleOrNull(Document doc, string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            try
            {
                var style = doc.Styles[name];
                if (style != null)
                {
                    return style;
                }
            }
            catch
            {
                // 继续模糊匹配
            }

            for (var i = 1; i <= doc.Styles.Count; i++)
            {
                Style candidate = null;
                try
                {
                    candidate = doc.Styles[i];
                    if (string.Equals(candidate.NameLocal, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return candidate;
                    }
                }
                catch
                {
                    // 忽略单个样式读取失败
                }

                TryRelease(candidate);
            }

            return null;
        }

        private static DocumentPropertySet ReadDocumentProperties(Document doc)
        {
            var result = new DocumentPropertySet();
            try
            {
                var props = (OfficeCore.DocumentProperties)doc.BuiltInDocumentProperties;
                result.Title = ReadDocumentProperty(props, "Title");
                result.Author = ReadDocumentProperty(props, "Author");
                result.Subject = ReadDocumentProperty(props, "Subject");
                result.Keywords = ReadDocumentProperty(props, "Keywords");
                result.Comments = ReadDocumentProperty(props, "Comments");
                result.Category = ReadDocumentProperty(props, "Category");
                result.Company = ReadDocumentProperty(props, "Company");
                result.Manager = ReadDocumentProperty(props, "Manager");
                result.LastAuthor = ReadDocumentProperty(props, "Last Author");
                result.RevisionNumber = ReadDocumentProperty(props, "Revision Number");
            }
            catch
            {
                // 忽略内置属性读取失败
            }

            try
            {
                var custom = (OfficeCore.DocumentProperties)doc.CustomDocumentProperties;
                for (var i = 1; i <= custom.Count; i++)
                {
                    try
                    {
                        var property = custom[i];
                        result.Custom[property.Name] = property.Value?.ToString();
                    }
                    catch
                    {
                        // 忽略单个自定义属性
                    }
                }
            }
            catch
            {
                // 忽略自定义属性读取失败
            }

            return result;
        }

        private static string ReadDocumentProperty(OfficeCore.DocumentProperties props, string name)
        {
            try
            {
                var property = props[name];
                return property?.Value?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static void WriteDocumentProperty(OfficeCore.DocumentProperties props, string name, string value)
        {
            if (value == null)
            {
                return;
            }

            try
            {
                var property = props[name];
                if (property != null)
                {
                    property.Value = value;
                }
            }
            catch
            {
                // 单个属性写入失败忽略
            }
        }

        private static void WriteCustomDocumentProperty(OfficeCore.DocumentProperties custom, string name, string value)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            try
            {
                for (var i = 1; i <= custom.Count; i++)
                {
                    var existing = custom[i];
                    if (string.Equals(existing.Name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        existing.Value = value ?? string.Empty;
                        return;
                    }
                }

                custom.Add(name, false, OfficeCore.MsoDocProperties.msoPropertyTypeString, value ?? string.Empty);
            }
            catch
            {
                // 忽略
            }
        }

        private static WdLineStyle ParseBorderStyle(string style)
        {
            switch (style.Trim().ToLowerInvariant())
            {
                case "single": return WdLineStyle.wdLineStyleSingle;
                case "double": return WdLineStyle.wdLineStyleDouble;
                case "dotted":
                case "dot": return WdLineStyle.wdLineStyleDot;
                case "dashed":
                case "dash": return WdLineStyle.wdLineStyleDashSmallGap;
                case "none": return WdLineStyle.wdLineStyleNone;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的边框样式：{style}（支持 single / double / dotted / dashed / none）");
            }
        }

        private static WdLineWidth ParseBorderWidth(string width)
        {
            switch (width.Trim().ToLowerInvariant())
            {
                case "thin":
                case "0.5": return WdLineWidth.wdLineWidth050pt;
                case "medium":
                case "1": return WdLineWidth.wdLineWidth100pt;
                case "thick":
                case "1.5": return WdLineWidth.wdLineWidth150pt;
                case "2.25": return WdLineWidth.wdLineWidth225pt;
                case "3": return WdLineWidth.wdLineWidth300pt;
                default:
                    if (double.TryParse(width, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var points))
                    {
                        if (points <= 0.5) return WdLineWidth.wdLineWidth050pt;
                        if (points <= 0.75) return WdLineWidth.wdLineWidth075pt;
                        if (points <= 1) return WdLineWidth.wdLineWidth100pt;
                        if (points <= 1.5) return WdLineWidth.wdLineWidth150pt;
                        if (points <= 2.25) return WdLineWidth.wdLineWidth225pt;
                        if (points <= 3) return WdLineWidth.wdLineWidth300pt;
                        if (points <= 4.5) return WdLineWidth.wdLineWidth450pt;
                        return WdLineWidth.wdLineWidth600pt;
                    }

                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的边框宽度：{width}（支持 thin / medium / thick 或磅值）");
            }
        }

        private static void UpdateHeaderFooterFields(HeaderFooter headerFooter)
        {
            try
            {
                if (headerFooter != null && headerFooter.Exists)
                {
                    headerFooter.Range.Fields.Update();
                }
            }
            catch
            {
                // 忽略单个页眉页脚更新失败
            }
        }

        /// <summary>移除所有节页眉中的水印形状，返回删除数量。</summary>
        private static int RemoveWatermarkCore(Document doc)
        {
            var removed = 0;
            for (var s = 1; s <= doc.Sections.Count; s++)
            {
                Section section = null;
                try
                {
                    section = doc.Sections[s];
                    foreach (var headerIndex in new[]
                             {
                                 WdHeaderFooterIndex.wdHeaderFooterPrimary,
                                 WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                                 WdHeaderFooterIndex.wdHeaderFooterEvenPages,
                             })
                    {
                        Shapes shapes = null;
                        try
                        {
                            var header = section.Headers[headerIndex];

                            // 链接到前一节的页眉共享形状，跳过避免重复处理
                            if (s > 1 && SafeRead(() => header.LinkToPrevious))
                            {
                                continue;
                            }

                            shapes = header.Shapes;
                            for (var i = shapes.Count; i >= 1; i--)
                            {
                                Shape shape = null;
                                try
                                {
                                    shape = shapes[i];
                                    if (string.Equals(shape.Name, WatermarkShapeName, StringComparison.Ordinal))
                                    {
                                        shape.Delete();
                                        removed++;
                                    }
                                }
                                finally
                                {
                                    TryRelease(shape);
                                }
                            }
                        }
                        catch
                        {
                            // 忽略单个页眉
                        }
                        finally
                        {
                            TryRelease(shapes);
                        }
                    }
                }
                finally
                {
                    TryRelease(section);
                }
            }

            return removed;
        }

        // ---------- 论文增强辅助 ----------

        /// <summary>
        /// 解析题注标签：优先匹配现有标签（兼容中文“图表/表格”与英文“Figure/Table”），
        /// 不存在时创建自定义标签。
        /// </summary>
        private string ResolveCaptionLabel(string label)
        {
            var requested = string.IsNullOrWhiteSpace(label) ? "图" : label.Trim();
            string[] candidates;
            switch (requested.ToLowerInvariant())
            {
                case "图":
                case "图表":
                    candidates = new[] { "图表", "图", "Figure" };
                    break;
                case "figure":
                case "fig":
                    candidates = new[] { "Figure", "图表", "图" };
                    break;
                case "表":
                case "表格":
                    candidates = new[] { "表格", "表", "Table" };
                    break;
                case "table":
                    candidates = new[] { "Table", "表格", "表" };
                    break;
                case "公式":
                    candidates = new[] { "公式", "Equation" };
                    break;
                case "equation":
                case "formula":
                    candidates = new[] { "Equation", "公式" };
                    break;
                default:
                    candidates = new[] { requested };
                    break;
            }

            var existing = new List<string>();
            try
            {
                var labels = _app.CaptionLabels;
                for (var i = 1; i <= labels.Count; i++)
                {
                    CaptionLabel captionLabel = null;
                    try
                    {
                        captionLabel = labels[i];
                        existing.Add(captionLabel.Name);
                    }
                    catch
                    {
                        // 忽略单个标签读取失败
                    }
                    finally
                    {
                        TryRelease(captionLabel);
                    }
                }
            }
            catch
            {
                // 忽略（无文档等场景）
            }

            foreach (var candidate in candidates)
            {
                foreach (var name in existing)
                {
                    if (string.Equals(name, candidate, StringComparison.OrdinalIgnoreCase))
                    {
                        return name;
                    }
                }
            }

            try
            {
                var created = _app.CaptionLabels.Add(requested);
                if (created != null && !string.IsNullOrEmpty(created.Name))
                {
                    return created.Name;
                }
            }
            catch
            {
                // 创建失败时回退到请求值
            }

            return requested;
        }

        private static string NormalizeNoteType(string noteType)
        {
            var type = (noteType ?? "footnote").Trim().ToLowerInvariant();
            switch (type)
            {
                case "footnote":
                case "footnotes":
                case "脚注":
                    return "footnote";
                case "endnote":
                case "endnotes":
                case "尾注":
                    return "endnote";
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的注释类型：{noteType}（支持 footnote / endnote）");
            }
        }

        private static void ValidateNoteIndex(Document doc, string type, int index)
        {
            var count = type == "footnote" ? doc.Footnotes.Count : doc.Endnotes.Count;
            if (index < 1 || index > count)
            {
                throw new McpToolException(McpErrorCodes.NoteNotFound,
                    $"注释索引 {index} 超出范围（当前共 {count} 个{(type == "footnote" ? "脚注" : "尾注")}）");
            }
        }

        private static void ValidateStartingNumber(int startingNumber)
        {
            if (startingNumber < 1)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, "starting_number 必须 >= 1");
            }
        }

        private static NoteInfo DescribeNote(Document doc, string type, int index)
        {
            int referenceStart;
            string noteText;
            int noteCount;
            if (type == "footnote")
            {
                Footnote note = null;
                try
                {
                    note = doc.Footnotes[index];
                    referenceStart = SafeRead(() => note.Reference.Start);
                    noteText = TrimParagraphText(SafeRead(() => note.Range.Text));
                    noteCount = doc.Footnotes.Count;
                }
                finally
                {
                    TryRelease(note);
                }
            }
            else
            {
                Endnote note = null;
                try
                {
                    note = doc.Endnotes[index];
                    referenceStart = SafeRead(() => note.Reference.Start);
                    noteText = TrimParagraphText(SafeRead(() => note.Range.Text));
                    noteCount = doc.Endnotes.Count;
                }
                finally
                {
                    TryRelease(note);
                }
            }

            var referenceParagraph = GetParagraphIndexAt(doc, referenceStart);
            return new NoteInfo
            {
                NoteType = type,
                Index = index,
                Text = noteText,
                ReferenceParagraphIndex = referenceParagraph,
                ReferenceText = GetParagraphPreview(doc, referenceParagraph),
                NoteCount = noteCount,
            };
        }

        private static string GetParagraphPreview(Document doc, int paragraphIndex)
        {
            if (paragraphIndex < 1 || paragraphIndex > SafeRead(() => doc.Paragraphs.Count))
            {
                return null;
            }

            var text = TrimParagraphText(SafeRead(() => doc.Paragraphs[paragraphIndex].Range.Text));
            if (text == null)
            {
                return null;
            }

            return text.Length > 120 ? text.Substring(0, 120) + "…" : text;
        }

        private static NoteOptionsInfo DescribeNoteOptions(Document doc, string type)
        {
            var info = new NoteOptionsInfo
            {
                NoteType = type,
                Location = type == "footnote" ? "bottom" : null,
            };

            if (type == "footnote")
            {
                var notes = doc.Footnotes;
                info.NumberStyle = SafeRead(() => DescribeNoteNumberStyle(notes.NumberStyle));
                info.NumberingRule = SafeRead(() => DescribeNumberingRule(notes.NumberingRule));
                info.StartingNumber = SafeRead(() => notes.StartingNumber);
                info.Location = SafeRead(() => DescribeFootnoteLocation(notes.Location));
                info.NoteCount = SafeRead(() => notes.Count);
            }
            else
            {
                var notes = doc.Endnotes;
                info.NumberStyle = SafeRead(() => DescribeNoteNumberStyle(notes.NumberStyle));
                info.NumberingRule = SafeRead(() => DescribeNumberingRule(notes.NumberingRule));
                info.StartingNumber = SafeRead(() => notes.StartingNumber);
                info.NoteCount = SafeRead(() => notes.Count);
            }

            return info;
        }

        private static WdNoteNumberStyle ParseNoteNumberStyle(string style)
        {
            switch (style.Trim().ToLowerInvariant())
            {
                case "arabic":
                case "1":
                    return WdNoteNumberStyle.wdNoteNumberStyleArabic;
                case "upper-roman":
                case "i":
                    return WdNoteNumberStyle.wdNoteNumberStyleUppercaseRoman;
                case "lower-roman":
                case "ii":
                    return WdNoteNumberStyle.wdNoteNumberStyleLowercaseRoman;
                case "upper-letter":
                case "a":
                    return WdNoteNumberStyle.wdNoteNumberStyleUppercaseLetter;
                case "lower-letter":
                case "b":
                    return WdNoteNumberStyle.wdNoteNumberStyleLowercaseLetter;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的编号样式：{style}（支持 arabic / upper-roman / lower-roman / upper-letter / lower-letter）");
            }
        }

        private static string DescribeNoteNumberStyle(WdNoteNumberStyle style)
        {
            switch (style)
            {
                case WdNoteNumberStyle.wdNoteNumberStyleArabic: return "arabic";
                case WdNoteNumberStyle.wdNoteNumberStyleUppercaseRoman: return "upper-roman";
                case WdNoteNumberStyle.wdNoteNumberStyleLowercaseRoman: return "lower-roman";
                case WdNoteNumberStyle.wdNoteNumberStyleUppercaseLetter: return "upper-letter";
                case WdNoteNumberStyle.wdNoteNumberStyleLowercaseLetter: return "lower-letter";
                default: return style.ToString();
            }
        }

        private static WdNumberingRule ParseNumberingRule(string rule)
        {
            switch (rule.Trim().ToLowerInvariant())
            {
                case "continuous":
                case "restart-continuous":
                    return WdNumberingRule.wdRestartContinuous;
                case "restart-section":
                case "section":
                    return WdNumberingRule.wdRestartSection;
                case "restart-page":
                case "page":
                    return WdNumberingRule.wdRestartPage;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的编号规则：{rule}（支持 continuous / restart-section / restart-page）");
            }
        }

        private static string DescribeNumberingRule(WdNumberingRule rule)
        {
            switch (rule)
            {
                case WdNumberingRule.wdRestartContinuous: return "continuous";
                case WdNumberingRule.wdRestartSection: return "restart-section";
                case WdNumberingRule.wdRestartPage: return "restart-page";
                default: return rule.ToString();
            }
        }

        private static WdFootnoteLocation ParseFootnoteLocation(string location)
        {
            switch (location.Trim().ToLowerInvariant())
            {
                case "bottom":
                case "bottom-of-page":
                    return WdFootnoteLocation.wdBottomOfPage;
                case "beneath-text":
                case "beneath":
                    return WdFootnoteLocation.wdBeneathText;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的脚注位置：{location}（支持 bottom / beneath-text）");
            }
        }

        private static string DescribeFootnoteLocation(WdFootnoteLocation location)
        {
            switch (location)
            {
                case WdFootnoteLocation.wdBottomOfPage: return "bottom";
                case WdFootnoteLocation.wdBeneathText: return "beneath-text";
                default: return location.ToString();
            }
        }

        private static HyperlinkInfo DescribeHyperlink(Document doc, int index)
        {
            Hyperlink link = null;
            try
            {
                link = doc.Hyperlinks[index];
                var start = SafeRead(() => link.Range.Start);
                return new HyperlinkInfo
                {
                    Index = index,
                    Text = SafeRead(() => link.TextToDisplay),
                    Address = SafeRead(() => link.Address),
                    SubAddress = SafeRead(() => link.SubAddress),
                    ScreenTip = SafeRead(() => link.ScreenTip),
                    ParagraphIndex = GetParagraphIndexAt(doc, start),
                    HyperlinkCount = doc.Hyperlinks.Count,
                };
            }
            finally
            {
                TryRelease(link);
            }
        }

        private static int FindHyperlinkIndex(Document doc, Hyperlink link)
        {
            try
            {
                var start = link.Range.Start;
                for (var i = 1; i <= doc.Hyperlinks.Count; i++)
                {
                    Hyperlink candidate = null;
                    try
                    {
                        candidate = doc.Hyperlinks[i];
                        if (candidate.Range.Start == start)
                        {
                            return i;
                        }
                    }
                    finally
                    {
                        TryRelease(candidate);
                    }
                }
            }
            catch
            {
                // 忽略
            }

            return doc.Hyperlinks.Count;
        }

        private static WdReferenceKind ParseReferenceKind(string kind, string referenceType)
        {
            var key = (kind ?? string.Empty).Trim().ToLowerInvariant();
            if (key.Length == 0)
            {
                switch (referenceType)
                {
                    case "footnote": return WdReferenceKind.wdFootnoteNumber;
                    case "endnote": return WdReferenceKind.wdEndnoteNumber;
                    case "caption": return WdReferenceKind.wdOnlyLabelAndNumber;
                    default: return WdReferenceKind.wdContentText;
                }
            }

            switch (key)
            {
                case "text":
                case "content":
                    return referenceType == "caption"
                        ? WdReferenceKind.wdOnlyCaptionText
                        : WdReferenceKind.wdContentText;
                case "number":
                    switch (referenceType)
                    {
                        case "footnote": return WdReferenceKind.wdFootnoteNumber;
                        case "endnote": return WdReferenceKind.wdEndnoteNumber;
                        case "caption": return WdReferenceKind.wdOnlyLabelAndNumber;
                        default: return WdReferenceKind.wdNumberFullContext;
                    }

                case "label-and-number":
                case "label":
                    return WdReferenceKind.wdOnlyLabelAndNumber;
                case "entire-caption":
                case "caption":
                    return WdReferenceKind.wdEntireCaption;
                case "page":
                    return WdReferenceKind.wdPageNumber;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的引用方式：{kind}（支持 text / number / page / label-and-number / entire-caption）");
            }
        }

        private static string DescribeReferenceKind(WdReferenceKind kind)
        {
            switch (kind)
            {
                case WdReferenceKind.wdContentText: return "text";
                case WdReferenceKind.wdNumberFullContext: return "number";
                case WdReferenceKind.wdOnlyLabelAndNumber: return "label-and-number";
                case WdReferenceKind.wdOnlyCaptionText: return "text";
                case WdReferenceKind.wdEntireCaption: return "entire-caption";
                case WdReferenceKind.wdFootnoteNumber: return "number";
                case WdReferenceKind.wdEndnoteNumber: return "number";
                case WdReferenceKind.wdPageNumber: return "page";
                default: return kind.ToString();
            }
        }

        private static int FindFieldIndex(Document doc, Field field)
        {
            try
            {
                var start = field.Result.Start;
                for (var i = 1; i <= doc.Fields.Count; i++)
                {
                    Field candidate = null;
                    try
                    {
                        candidate = doc.Fields[i];
                        if (candidate.Result.Start == start)
                        {
                            return i;
                        }
                    }
                    finally
                    {
                        TryRelease(candidate);
                    }
                }
            }
            catch
            {
                // 忽略
            }

            return doc.Fields.Count;
        }

        /// <summary>
        /// 文本规范化：把真实换行与字面量 \n / \r（不同客户端 JSON 转义不一致）
        /// 统一转为 Word 手动换行符（\v），避免意外拆成多个段落。
        /// </summary>
        private static string NormalizeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            return text.Replace("\r\n", "\v").Replace("\r", "\v").Replace("\n", "\v")
                .Replace("\\n", "\v").Replace("\\r", "\v");
        }

        /// <summary>去掉段落/单元格末尾的段落标记与单元格结束符，并把换行统一为 \n。</summary>
        private static string TrimParagraphText(string text)
        {
            if (text == null)
            {
                return string.Empty;
            }

            return text.TrimEnd('\r', '\a').Replace("\v", "\n").Replace("\r", "\n");
        }

        // ---------- P0 增强辅助 ----------

        private SelectionInfo DescribeSelection(Document doc)
        {
            var selection = _app.Selection;
            var start = SafeRead(() => selection.Start);
            var end = SafeRead(() => selection.End);
            return new SelectionInfo
            {
                SelectionType = SafeRead(() => DescribeSelectionType(selection.Type)),
                Start = start,
                End = end,
                // 光标（空选区）时 Selection.Text 会返回光标后的字符，这里统一置空
                Text = start == end ? string.Empty : SafeRead(() => selection.Text),
                ParagraphIndex = GetParagraphIndexAt(doc, end),
                PageNumber = SafeRead(() => (int)selection.Information[WdInformation.wdActiveEndPageNumber]),
                IsEmpty = start == end,
            };
        }

        private static string DescribeSelectionType(WdSelectionType type)
        {
            switch (type)
            {
                case WdSelectionType.wdNoSelection: return "none";
                case WdSelectionType.wdSelectionIP: return "insertion-point";
                case WdSelectionType.wdSelectionNormal: return "normal";
                case WdSelectionType.wdSelectionRow: return "row";
                case WdSelectionType.wdSelectionColumn: return "column";
                case WdSelectionType.wdSelectionBlock: return "block";
                case WdSelectionType.wdSelectionInlineShape: return "inline-shape";
                case WdSelectionType.wdSelectionShape: return "shape";
                default: return type.ToString();
            }
        }

        private static ImageEntry DescribeInlineImage(Document doc, int index)
        {
            InlineShape shape = null;
            try
            {
                shape = doc.InlineShapes[index];
                var rangeStart = SafeRead(() => shape.Range.Start);
                return new ImageEntry
                {
                    Kind = "inline",
                    Index = index,
                    Width = SafeRead(() => (double)shape.Width),
                    Height = SafeRead(() => (double)shape.Height),
                    AltText = SafeRead(() => shape.AlternativeText),
                    ParagraphIndex = GetParagraphIndexAt(doc, rangeStart),
                };
            }
            finally
            {
                TryRelease(shape);
            }
        }

        private static ImageEntry DescribeFloatingImage(Document doc, Shape shape)
        {
            var anchorStart = SafeRead(() => shape.Anchor.Start);
            return new ImageEntry
            {
                Kind = "floating",
                Index = FindShapeIndex(doc, shape),
                Name = SafeRead(() => shape.Name),
                Width = SafeRead(() => (double)shape.Width),
                Height = SafeRead(() => (double)shape.Height),
                Left = SafeRead(() => (double?)shape.Left),
                Top = SafeRead(() => (double?)shape.Top),
                Wrap = SafeRead(() => DescribeWrapType(shape.WrapFormat.Type)),
                AltText = SafeRead(() => shape.AlternativeText),
                ParagraphIndex = GetParagraphIndexAt(doc, anchorStart),
            };
        }

        private static int FindInlineShapeIndex(Document doc, InlineShape shape)
        {
            try
            {
                var start = shape.Range.Start;
                for (var i = 1; i <= doc.InlineShapes.Count; i++)
                {
                    if (doc.InlineShapes[i].Range.Start == start)
                    {
                        return i;
                    }
                }
            }
            catch
            {
                // 忽略
            }

            return doc.InlineShapes.Count;
        }

        private static int FindShapeIndex(Document doc, Shape shape)
        {
            try
            {
                var name = shape.Name;
                for (var i = 1; i <= doc.Shapes.Count; i++)
                {
                    if (string.Equals(doc.Shapes[i].Name, name, StringComparison.Ordinal))
                    {
                        return i;
                    }
                }
            }
            catch
            {
                // 忽略
            }

            return doc.Shapes.Count;
        }

        private static Shape FindShapeOrNull(Document doc, string name, int? index)
        {
            if (!string.IsNullOrEmpty(name))
            {
                try
                {
                    return doc.Shapes[name];
                }
                catch
                {
                    // 继续按索引查找
                }
            }

            if (index.HasValue)
            {
                if (index.Value < 1 || index.Value > doc.Shapes.Count)
                {
                    throw new McpToolException(McpErrorCodes.ImageNotFound,
                        $"浮动图片索引 {index.Value} 超出范围（当前共 {doc.Shapes.Count} 个）");
                }

                return doc.Shapes[index.Value];
            }

            return null;
        }

        private static string DescribeWrapType(WdWrapType type)
        {
            switch (type)
            {
                case WdWrapType.wdWrapSquare: return "square";
                case WdWrapType.wdWrapTight: return "tight";
                case WdWrapType.wdWrapThrough: return "through";
                case WdWrapType.wdWrapTopBottom: return "top-bottom";
                case WdWrapType.wdWrapBehind: return "behind";
                case WdWrapType.wdWrapNone: return "none";
                default: return type.ToString();
            }
        }

        private static WdWrapType ParseWrapType(string wrap)
        {
            switch (wrap.Trim().ToLowerInvariant())
            {
                case "square": return WdWrapType.wdWrapSquare;
                case "tight": return WdWrapType.wdWrapTight;
                case "through": return WdWrapType.wdWrapThrough;
                case "none": return WdWrapType.wdWrapNone;
                case "top-bottom": return WdWrapType.wdWrapTopBottom;
                case "behind": return WdWrapType.wdWrapBehind;
                case "front": return WdWrapType.wdWrapFront;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的环绕方式：{wrap}（支持 square / tight / through / none / top-bottom / behind / front）");
            }
        }

        private static WdRelativeHorizontalPosition ParseRelativeHorizontal(string relativeTo)
        {
            switch ((relativeTo ?? "page").Trim().ToLowerInvariant())
            {
                case "page": return WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                case "margin": return WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                case "column":
                case "paragraph": return WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的定位基准：{relativeTo}（支持 page / margin / column / paragraph）");
            }
        }

        private static WdRelativeVerticalPosition ParseRelativeVertical(string relativeTo)
        {
            switch ((relativeTo ?? "page").Trim().ToLowerInvariant())
            {
                case "page": return WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                case "margin": return WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
                case "column":
                case "paragraph": return WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的定位基准：{relativeTo}（支持 page / margin / column / paragraph）");
            }
        }

        private static WdHeaderFooterIndex ParseHeaderFooterType(string headerType)
        {
            switch ((headerType ?? "primary").Trim().ToLowerInvariant())
            {
                case "primary":
                case "first": // 兼容旧参数误用
                case "":
                    return WdHeaderFooterIndex.wdHeaderFooterPrimary;
                case "first-page":
                case "firstpage":
                    return WdHeaderFooterIndex.wdHeaderFooterFirstPage;
                case "even-pages":
                case "even":
                    return WdHeaderFooterIndex.wdHeaderFooterEvenPages;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的页眉页脚类型：{headerType}（支持 primary / first-page / even-pages）");
            }
        }

        private static string NormalizeHeaderTypeName(string headerType)
        {
            switch ((headerType ?? "primary").Trim().ToLowerInvariant())
            {
                case "first-page":
                case "firstpage":
                    return "first-page";
                case "even-pages":
                case "even":
                    return "even-pages";
                default:
                    return "primary";
            }
        }

        private static PageNumbersInfo DescribePageNumbers(Document doc, int sectionIndex)
        {
            var numbers = doc.Sections[sectionIndex].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers;
            return new PageNumbersInfo
            {
                SectionIndex = sectionIndex,
                NumberStyle = SafeRead(() => DescribePageNumberStyle(numbers.NumberStyle)),
                StartingNumber = SafeRead(() => numbers.StartingNumber),
                RestartAtSection = SafeRead(() => numbers.RestartNumberingAtSection),
            };
        }

        private static WdPageNumberStyle ParsePageNumberStyle(string style)
        {
            switch (style.Trim().ToLowerInvariant())
            {
                case "arabic": return WdPageNumberStyle.wdPageNumberStyleArabic;
                case "upper-roman": return WdPageNumberStyle.wdPageNumberStyleUppercaseRoman;
                case "lower-roman": return WdPageNumberStyle.wdPageNumberStyleLowercaseRoman;
                case "upper-letter": return WdPageNumberStyle.wdPageNumberStyleUppercaseLetter;
                case "lower-letter": return WdPageNumberStyle.wdPageNumberStyleLowercaseLetter;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的页码样式：{style}（支持 arabic / upper-roman / lower-roman / upper-letter / lower-letter）");
            }
        }

        private static string DescribePageNumberStyle(WdPageNumberStyle style)
        {
            switch (style)
            {
                case WdPageNumberStyle.wdPageNumberStyleArabic: return "arabic";
                case WdPageNumberStyle.wdPageNumberStyleUppercaseRoman: return "upper-roman";
                case WdPageNumberStyle.wdPageNumberStyleLowercaseRoman: return "lower-roman";
                case WdPageNumberStyle.wdPageNumberStyleUppercaseLetter: return "upper-letter";
                case WdPageNumberStyle.wdPageNumberStyleLowercaseLetter: return "lower-letter";
                default: return style.ToString();
            }
        }

        private static void ValidateRevisionIndex(Document doc, int index)
        {
            var count = doc.Revisions.Count;
            if (index < 1 || index > count)
            {
                throw new McpToolException(McpErrorCodes.RevisionNotFound,
                    $"修订索引 {index} 超出范围（当前共 {count} 处修订）");
            }
        }

        private static RevisionInfo DescribeRevision(Document doc, int index)
        {
            Revision revision = null;
            try
            {
                revision = doc.Revisions[index];
                var start = SafeRead(() => revision.Range.Start);
                return new RevisionInfo
                {
                    Index = index,
                    Type = SafeRead(() => DescribeRevisionType(revision.Type)),
                    Author = SafeRead(() => revision.Author),
                    Date = SafeRead(() => revision.Date.ToString("yyyy-MM-dd HH:mm:ss")),
                    Text = TrimParagraphText(SafeRead(() => revision.Range.Text)),
                    ParagraphIndex = GetParagraphIndexAt(doc, start),
                };
            }
            finally
            {
                TryRelease(revision);
            }
        }

        private static string DescribeRevisionType(WdRevisionType type)
        {
            switch (type)
            {
                case WdRevisionType.wdRevisionInsert: return "insert";
                case WdRevisionType.wdRevisionDelete: return "delete";
                case WdRevisionType.wdRevisionProperty: return "property";
                case WdRevisionType.wdRevisionCellInsertion: return "cell-insert";
                case WdRevisionType.wdRevisionCellDeletion: return "cell-delete";
                case WdRevisionType.wdRevisionCellMerge: return "cell-merge";
                case WdRevisionType.wdRevisionCellSplit: return "cell-split";
                case WdRevisionType.wdRevisionParagraphProperty: return "paragraph-format";
                case WdRevisionType.wdRevisionSectionProperty: return "section-format";
                case WdRevisionType.wdRevisionTableProperty: return "table-format";
                case WdRevisionType.wdRevisionStyle: return "style";
                case WdRevisionType.wdRevisionReplace: return "replace";
                default: return type.ToString();
            }
        }

        private static void ValidateCommentIndex(Document doc, int index)
        {
            var count = doc.Comments.Count;
            if (index < 1 || index > count)
            {
                throw new McpToolException(McpErrorCodes.CommentNotFound,
                    $"批注索引 {index} 超出范围（当前共 {count} 条批注）");
            }
        }

        private static WdProtectionType ParseProtectionType(string protectionType)
        {
            switch ((protectionType ?? "read-only").Trim().ToLowerInvariant())
            {
                case "read-only":
                case "readonly":
                case "reading": return WdProtectionType.wdAllowOnlyReading;
                case "comments": return WdProtectionType.wdAllowOnlyComments;
                case "forms": return WdProtectionType.wdAllowOnlyFormFields;
                case "revisions": return WdProtectionType.wdAllowOnlyRevisions;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的保护类型：{protectionType}（支持 read-only / comments / forms / revisions）");
            }
        }

        private static ProtectionInfo DescribeProtection(Document doc)
        {
            return new ProtectionInfo
            {
                ProtectionType = SafeRead(() => DescribeProtectionType(doc.ProtectionType)),
                Protected = SafeRead(() => doc.ProtectionType != WdProtectionType.wdNoProtection),
                IsFinal = SafeRead(() => doc.Final),
            };
        }

        private static string DescribeProtectionType(WdProtectionType type)
        {
            switch (type)
            {
                case WdProtectionType.wdNoProtection: return "none";
                case WdProtectionType.wdAllowOnlyReading: return "read-only";
                case WdProtectionType.wdAllowOnlyComments: return "comments";
                case WdProtectionType.wdAllowOnlyFormFields: return "forms";
                case WdProtectionType.wdAllowOnlyRevisions: return "revisions";
                default: return type.ToString();
            }
        }

        private static WdSortFieldType ParseSortFieldType(string fieldType)
        {
            switch ((fieldType ?? "alphanumeric").Trim().ToLowerInvariant())
            {
                case "numeric":
                case "number": return WdSortFieldType.wdSortFieldNumeric;
                case "date": return WdSortFieldType.wdSortFieldDate;
                default: return WdSortFieldType.wdSortFieldAlphanumeric;
            }
        }

        /// <summary>按范围语义解析查找范围（document / selection / headers-footers）。</summary>
        private List<Range> ResolveFindRanges(Document doc, string scope)
        {
            var ranges = new List<Range>();
            switch ((scope ?? "document").Trim().ToLowerInvariant())
            {
                case "document":
                case "":
                    ranges.Add(doc.Content);
                    break;
                case "selection":
                    ranges.Add(_app.Selection.Range);
                    break;
                case "headers-footers":
                case "headers":
                    for (var s = 1; s <= doc.Sections.Count; s++)
                    {
                        Section section = null;
                        try
                        {
                            section = doc.Sections[s];
                            foreach (var type in new[]
                                     {
                                         WdHeaderFooterIndex.wdHeaderFooterPrimary,
                                         WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                                         WdHeaderFooterIndex.wdHeaderFooterEvenPages,
                                     })
                            {
                                try
                                {
                                    var header = section.Headers[type];
                                    if (header.Exists)
                                    {
                                        ranges.Add(header.Range);
                                    }
                                }
                                catch
                                {
                                    // 忽略
                                }

                                try
                                {
                                    var footer = section.Footers[type];
                                    if (footer.Exists)
                                    {
                                        ranges.Add(footer.Range);
                                    }
                                }
                                catch
                                {
                                    // 忽略
                                }
                            }
                        }
                        finally
                        {
                            TryRelease(section);
                        }
                    }

                    break;
                default:
                    throw new McpToolException(McpErrorCodes.UnsupportedType,
                        $"不支持的查找范围：{scope}（支持 document / selection / headers-footers）");
            }

            return ranges;
        }

        /// <summary>在单个范围内执行替换；occurrence 由调用方跨范围累计。</summary>
        private static int ReplaceInRange(Range range, FindReplaceRequest request, ref int occurrenceCounter)
        {
            var find = range.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Text = request.FindText;
            find.Replacement.Text = NormalizeText(request.ReplaceText) ?? string.Empty;
            find.MatchCase = request.MatchCase;
            find.MatchWholeWord = request.MatchWholeWord;
            find.MatchWildcards = request.MatchWildcards;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;

            object missing = Type.Missing;
            object findTextObj = request.FindText;
            object replaceOne = WdReplace.wdReplaceOne;
            var replaced = 0;

            if (request.Occurrence.HasValue)
            {
                while (occurrenceCounter < request.Occurrence.Value &&
                       find.Execute(ref findTextObj, ref missing, ref missing, ref missing, ref missing,
                           ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                           ref missing, ref missing, ref missing, ref missing))
                {
                    occurrenceCounter++;
                    if (occurrenceCounter == request.Occurrence.Value)
                    {
                        range.Text = NormalizeText(request.ReplaceText) ?? string.Empty;
                        replaced = 1;
                        break;
                    }

                    range.Collapse(WdCollapseDirection.wdCollapseEnd);
                }

                return replaced;
            }

            if (request.ReplaceAll)
            {
                while (replaced < 10000 && find.Execute(ref findTextObj, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceOne,
                    ref missing, ref missing, ref missing, ref missing))
                {
                    replaced++;
                    range.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
            }
            else if (find.Execute(ref findTextObj, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceOne,
                ref missing, ref missing, ref missing, ref missing))
            {
                replaced = 1;
            }

            return replaced;
        }

        /// <summary>找到第 N 处匹配文本的范围（1 基）。</summary>
        private static Range FindNthRange(Document doc, string text, int occurrence)
        {
            if (occurrence < 1)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, "occurrence 必须 >= 1");
            }

            var range = doc.Content;
            var find = range.Find;
            find.ClearFormatting();
            find.Text = text;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;

            object missing = Type.Missing;
            object findTextObj = text;
            var count = 0;
            while (count < 10000 && find.Execute(ref findTextObj, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing))
            {
                count++;
                if (count == occurrence)
                {
                    return range;
                }

                range.Collapse(WdCollapseDirection.wdCollapseEnd);
            }

            return null;
        }

        /// <summary>安全读取 COM 属性：失败返回默认值（避免单个属性异常拖垮整个请求）。</summary>
        private static T SafeRead<T>(Func<T> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return default(T);
            }
        }

        /// <summary>释放新建的 COM 对象（只释放本方法创建的对象，见类注释）。</summary>
        private static void TryRelease(params object[] comObjects)
        {
            foreach (var obj in comObjects)
            {
                if (obj == null)
                {
                    continue;
                }

                try
                {
                    Marshal.FinalReleaseComObject(obj);
                }
                catch
                {
                    // RCW 可能已失效（对象被删除等），忽略
                }
            }
        }
    }
}
