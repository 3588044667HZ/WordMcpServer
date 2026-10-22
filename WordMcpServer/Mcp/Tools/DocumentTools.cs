using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>获取当前活动文档的基本信息。</summary>
    public sealed class GetActiveDocumentInfoTool : ToolBase
    {
        public GetActiveDocumentInfoTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_active_document_info";

        public override string Description => "获取当前活动文档的信息：文件名、完整路径、是否已保存、是否只读、段落数、表格数、节数、页数、字数、字符数。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 get_active_document_info");
            return JObject.FromObject(Service.GetActiveDocumentInfo());
        }
    }

    /// <summary>列出所有已打开的文档。</summary>
    public sealed class ListDocumentsTool : ToolBase
    {
        public ListDocumentsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "list_documents";

        public override string Description => "列出当前 Word 中所有已打开的文档（索引、名称、路径、是否已保存、是否只读、是否活动文档）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 list_documents");
            var documents = Service.ListDocuments();
            return new JObject
            {
                ["count"] = documents.Count,
                ["documents"] = JArray.FromObject(documents),
            };
        }
    }

    /// <summary>新建空白文档。</summary>
    public sealed class NewDocumentTool : ToolBase
    {
        public NewDocumentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "new_document";

        public override string Description => "新建一个空白 Word 文档并返回其信息（该文档成为活动文档）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 new_document");
            return JObject.FromObject(Service.NewDocument());
        }
    }

    /// <summary>打开指定路径的文档。</summary>
    public sealed class OpenDocumentTool : ToolBase
    {
        public OpenDocumentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "open_document";

        public override string Description => "打开指定路径的 Word 文档（该文档成为活动文档）。read_only 为 true 时以只读方式打开。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["path"] = Str("文档完整路径（如 C:\\\\docs\\\\报告.docx）"),
            ["read_only"] = Bool("是否只读打开，默认 false"),
        }, "path");

        public override JObject Execute(JObject arguments)
        {
            var path = ToolArgumentValidator.RequireString(arguments, "path");
            var readOnly = ToolArgumentValidator.OptionalBool(arguments, "read_only") ?? false;
            Logger.Debug($"执行 open_document，path={path}，readOnly={readOnly}");
            return JObject.FromObject(Service.OpenDocument(path, readOnly));
        }
    }

    /// <summary>激活指定文档（切换活动窗口）。</summary>
    public sealed class ActivateDocumentTool : ToolBase
    {
        public ActivateDocumentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "activate_document";

        public override string Description => "激活（切换到）指定文档：按 name（文件名或完整路径）或 index（1 基）指定，两者都不传时返回当前活动文档信息。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["name"] = Str("文档名称或完整路径（可选）"),
            ["index"] = Int("文档索引（1 基，可选）", 1),
        });

        public override JObject Execute(JObject arguments)
        {
            var name = ToolArgumentValidator.OptionalString(arguments, "name");
            var index = ToolArgumentValidator.OptionalInt(arguments, "index", 1);
            Logger.Debug($"执行 activate_document，name={name}，index={index}");
            return JObject.FromObject(Service.ActivateDocument(name, index));
        }
    }

    /// <summary>关闭文档。</summary>
    public sealed class CloseDocumentTool : ToolBase
    {
        public CloseDocumentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "close_document";

        public override string Description => "关闭指定文档（按 name 或 index；都不传时关闭活动文档）。save_changes=true 保存修改，false 丢弃修改。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["name"] = Str("文档名称或完整路径（可选）"),
            ["index"] = Int("文档索引（1 基，可选）", 1),
            ["save_changes"] = Bool("是否保存修改后关闭，默认 false（丢弃未保存的修改）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var name = ToolArgumentValidator.OptionalString(arguments, "name");
            var index = ToolArgumentValidator.OptionalInt(arguments, "index", 1);
            var saveChanges = ToolArgumentValidator.OptionalBool(arguments, "save_changes") ?? false;
            Logger.Debug($"执行 close_document，name={name}，index={index}，saveChanges={saveChanges}");
            return JObject.FromObject(Service.CloseDocument(name, index, saveChanges));
        }
    }

    /// <summary>保存文档（可另存为）。</summary>
    public sealed class SaveDocumentTool : ToolBase
    {
        public SaveDocumentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "save_document";

        public override string Description => "保存当前文档。提供 save_as_path 时另存为新文件（format 可选 docx/doc/rtf/txt/odt，默认按扩展名推断）；不提供时直接保存（未命名文档会失败，请改用 save_as_path）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["save_as_path"] = Str("另存为的完整路径（可选，如 C:\\\\docs\\\\报告.docx）"),
            ["format"] = Str("保存格式（可选）", "docx", "doc", "rtf", "txt", "odt"),
        });

        public override JObject Execute(JObject arguments)
        {
            var saveAsPath = ToolArgumentValidator.OptionalString(arguments, "save_as_path");
            var format = ToolArgumentValidator.OptionalString(arguments, "format");
            Logger.Debug($"执行 save_document，saveAsPath={saveAsPath}，format={format}");
            return JObject.FromObject(Service.SaveDocument(saveAsPath, format));
        }
    }

    /// <summary>备份文档。</summary>
    public sealed class BackupDocumentTool : ToolBase
    {
        public BackupDocumentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "backup_document";

        public override string Description => "备份当前文档：已保存文档直接复制文件到 backup_path；未保存文档另存到 backup_path（会改变文档关联）。backup_path 缺省时在文档目录生成带时间戳的备份。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["backup_path"] = Str("备份文件完整路径（可选）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var backupPath = ToolArgumentValidator.OptionalString(arguments, "backup_path");
            Logger.Debug($"执行 backup_document，backupPath={backupPath}");
            return JObject.FromObject(Service.BackupDocument(backupPath));
        }
    }

    /// <summary>导出 PDF。</summary>
    public sealed class ExportPdfTool : ToolBase
    {
        public ExportPdfTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "export_to_pdf";

        public override string Description => "把当前文档导出为 PDF（ExportAsFixedFormat）。支持整篇/当前页/选区/起止页导出，可启用 PDF/A 与标题书签。output_path 必须以 .pdf 结尾。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["output_path"] = Str("PDF 输出完整路径（必须以 .pdf 结尾）"),
            ["open_after_export"] = Bool("导出后是否用默认阅读器打开，默认 false"),
            ["range"] = Str("导出范围", "all", "selection", "current-page", "from-to"),
            ["from_page"] = Int("起始页（range=from-to 时必填）", 1),
            ["to_page"] = Int("结束页（range=from-to 时必填）", 1),
            ["pdf_a"] = Bool("是否导出 PDF/A 归档格式，默认 false"),
            ["create_bookmarks"] = Bool("是否为标题创建书签，默认 true"),
        }, "output_path");

        public override JObject Execute(JObject arguments)
        {
            var request = new ExportPdfRequest
            {
                OutputPath = ToolArgumentValidator.RequireString(arguments, "output_path"),
                OpenAfterExport = ToolArgumentValidator.OptionalBool(arguments, "open_after_export") ?? false,
                Range = ToolArgumentValidator.OptionalString(arguments, "range") ?? "all",
                FromPage = ToolArgumentValidator.OptionalInt(arguments, "from_page", 1),
                ToPage = ToolArgumentValidator.OptionalInt(arguments, "to_page", 1),
                PdfA = ToolArgumentValidator.OptionalBool(arguments, "pdf_a") ?? false,
                CreateBookmarks = ToolArgumentValidator.OptionalBool(arguments, "create_bookmarks") ?? true,
            };
            Logger.Debug($"执行 export_to_pdf，outputPath={request.OutputPath}，range={request.Range}");
            return JObject.FromObject(Service.ExportToPdf(request));
        }
    }

    /// <summary>读取文档纯文本。</summary>
    public sealed class GetDocumentTextTool : ToolBase
    {
        public GetDocumentTextTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_document_text";

        public override string Description => "读取当前文档的纯文本内容（段落间以换行分隔）。max_chars 控制最大返回字符数（默认 20000），超出时截断并标记 truncated。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["max_chars"] = Int("最大返回字符数（默认 20000）", 100, 500000),
        });

        public override JObject Execute(JObject arguments)
        {
            var maxChars = ToolArgumentValidator.OptionalInt(arguments, "max_chars", 100, 500000) ?? 20000;
            Logger.Debug($"执行 get_document_text，maxChars={maxChars}");
            return JObject.FromObject(Service.GetDocumentText(maxChars));
        }
    }

    /// <summary>读取文档大纲（标题结构）。</summary>
    public sealed class GetDocumentOutlineTool : ToolBase
    {
        public GetDocumentOutlineTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_document_outline";

        public override string Description => "读取当前文档的标题大纲：返回每个标题的段落索引、级别（1-9）、文本与样式。适合快速了解长文档结构。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["max_items"] = Int("最多返回的标题数（默认 200）", 1, 2000),
        });

        public override JObject Execute(JObject arguments)
        {
            var maxItems = ToolArgumentValidator.OptionalInt(arguments, "max_items", 1, 2000) ?? 200;
            Logger.Debug($"执行 get_document_outline，maxItems={maxItems}");
            return JObject.FromObject(Service.GetDocumentOutline(maxItems));
        }
    }

    /// <summary>读取文档统计信息。</summary>
    public sealed class GetDocumentStatisticsTool : ToolBase
    {
        public GetDocumentStatisticsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_document_statistics";

        public override string Description => "读取当前文档的统计信息：页数、字数、字符数、含空格字符数、段落数、行数。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 get_document_statistics");
            return JObject.FromObject(Service.GetDocumentStatistics());
        }
    }

    /// <summary>更新所有域。</summary>
    public sealed class UpdateFieldsTool : ToolBase
    {
        public UpdateFieldsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "update_fields";

        public override string Description => "更新文档中的所有域：正文域、各节页眉页脚中的域（页码等）以及目录（TOC）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 update_fields");
            return JObject.FromObject(Service.UpdateFields());
        }
    }

    /// <summary>读取文档属性。</summary>
    public sealed class GetDocumentPropertiesTool : ToolBase
    {
        public GetDocumentPropertiesTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_document_properties";

        public override string Description => "读取文档属性：标题、作者、主题、关键词、备注、类别、公司、经理、最后修改者、修订号，以及自定义属性。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 get_document_properties");
            return JObject.FromObject(Service.GetDocumentProperties());
        }
    }

    /// <summary>设置文档属性。</summary>
    public sealed class SetDocumentPropertiesTool : ToolBase
    {
        public SetDocumentPropertiesTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_document_properties";

        public override string Description => "设置文档属性（未提供的字段保持不变）；custom 为自定义属性键值对（同名覆盖、不存在则新建）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["title"] = Str("标题"),
            ["author"] = Str("作者"),
            ["subject"] = Str("主题"),
            ["keywords"] = Str("关键词（多个用逗号分隔）"),
            ["comments"] = Str("备注"),
            ["category"] = Str("类别"),
            ["company"] = Str("公司"),
            ["manager"] = Str("经理"),
            ["custom"] = new JObject
            {
                ["type"] = "object",
                ["description"] = "自定义属性键值对（字符串）",
            },
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetDocumentPropertiesRequest
            {
                Title = ToolArgumentValidator.OptionalString(arguments, "title"),
                Author = ToolArgumentValidator.OptionalString(arguments, "author"),
                Subject = ToolArgumentValidator.OptionalString(arguments, "subject"),
                Keywords = ToolArgumentValidator.OptionalString(arguments, "keywords"),
                Comments = ToolArgumentValidator.OptionalString(arguments, "comments"),
                Category = ToolArgumentValidator.OptionalString(arguments, "category"),
                Company = ToolArgumentValidator.OptionalString(arguments, "company"),
                Manager = ToolArgumentValidator.OptionalString(arguments, "manager"),
                Custom = ToolArgumentValidator.OptionalStringDictionary(arguments, "custom"),
            };
            Logger.Debug("执行 set_document_properties");
            return JObject.FromObject(Service.SetDocumentProperties(request));
        }
    }

    /// <summary>设置修订跟踪。</summary>
    public sealed class SetTrackChangesTool : ToolBase
    {
        public SetTrackChangesTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_track_changes";

        public override string Description => "开启/关闭修订跟踪（Track Changes）。开启后所有修改会保留痕迹，供后续接受或拒绝。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["enabled"] = Bool("true 开启，false 关闭"),
        }, "enabled");

        public override JObject Execute(JObject arguments)
        {
            var enabled = ToolArgumentValidator.OptionalBool(arguments, "enabled") ?? true;
            Logger.Debug($"执行 set_track_changes，enabled={enabled}");
            return JObject.FromObject(Service.SetTrackChanges(enabled));
        }
    }

    /// <summary>接受所有修订。</summary>
    public sealed class AcceptRevisionsTool : ToolBase
    {
        public AcceptRevisionsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "accept_revisions";

        public override string Description => "接受文档中的所有修订（保留修改后的内容）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 accept_revisions");
            return JObject.FromObject(Service.AcceptRevisions());
        }
    }

    /// <summary>拒绝所有修订。</summary>
    public sealed class RejectRevisionsTool : ToolBase
    {
        public RejectRevisionsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "reject_revisions";

        public override string Description => "拒绝文档中的所有修订（恢复到修改前的内容）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 reject_revisions");
            return JObject.FromObject(Service.RejectRevisions());
        }
    }

    /// <summary>列出修订。</summary>
    public sealed class ListRevisionsTool : ToolBase
    {
        public ListRevisionsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "list_revisions";

        public override string Description => "列出文档中的修订（索引、类型、作者、时间、文本、所在段落）。";

        public override JObject InputSchema => EmptySchema();

        public override JObject Execute(JObject arguments)
        {
            Logger.Debug("执行 list_revisions");
            var revisions = Service.ListRevisions();
            return new JObject
            {
                ["count"] = revisions.Count,
                ["revisions"] = JArray.FromObject(revisions),
            };
        }
    }

    /// <summary>接受单条修订。</summary>
    public sealed class AcceptRevisionTool : ToolBase
    {
        public AcceptRevisionTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "accept_revision";

        public override string Description => "按索引接受单条修订（索引来自 list_revisions）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["index"] = Int("修订索引（1 基）", 1),
        }, "index");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "index", 1);
            Logger.Debug($"执行 accept_revision，index={index}");
            return JObject.FromObject(Service.AcceptRevision(index));
        }
    }

    /// <summary>拒绝单条修订。</summary>
    public sealed class RejectRevisionTool : ToolBase
    {
        public RejectRevisionTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "reject_revision";

        public override string Description => "按索引拒绝单条修订（索引来自 list_revisions）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["index"] = Int("修订索引（1 基）", 1),
        }, "index");

        public override JObject Execute(JObject arguments)
        {
            var index = ToolArgumentValidator.RequireInt(arguments, "index", 1);
            Logger.Debug($"执行 reject_revision，index={index}");
            return JObject.FromObject(Service.RejectRevision(index));
        }
    }

    /// <summary>保护文档。</summary>
    public sealed class ProtectDocumentTool : ToolBase
    {
        public ProtectDocumentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "protect_document";

        public override string Description => "保护文档：read-only（只读）/ comments（仅批注）/ forms（仅表单）/ revisions（仅修订）。可选密码。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["type"] = Str("保护类型", "read-only", "comments", "forms", "revisions"),
            ["password"] = Str("保护密码（可选）"),
        }, "type");

        public override JObject Execute(JObject arguments)
        {
            var type = ToolArgumentValidator.RequireString(arguments, "type");
            var password = ToolArgumentValidator.OptionalString(arguments, "password");
            Logger.Debug($"执行 protect_document，type={type}");
            return JObject.FromObject(Service.ProtectDocument(type, password));
        }
    }

    /// <summary>取消文档保护。</summary>
    public sealed class UnprotectDocumentTool : ToolBase
    {
        public UnprotectDocumentTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "unprotect_document";

        public override string Description => "取消文档保护（如设置过密码需提供）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["password"] = Str("保护密码（可选）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var password = ToolArgumentValidator.OptionalString(arguments, "password");
            Logger.Debug("执行 unprotect_document");
            return JObject.FromObject(Service.UnprotectDocument(password));
        }
    }

    /// <summary>标记为最终状态。</summary>
    public sealed class MarkAsFinalTool : ToolBase
    {
        public MarkAsFinalTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "mark_as_final";

        public override string Description => "把文档标记为最终状态（或取消标记），提示读者内容已定稿。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["final"] = Bool("true 标记为最终（默认），false 取消标记"),
        });

        public override JObject Execute(JObject arguments)
        {
            var final = ToolArgumentValidator.OptionalBool(arguments, "final") ?? true;
            Logger.Debug($"执行 mark_as_final，final={final}");
            return JObject.FromObject(Service.MarkAsFinal(final));
        }
    }
}
