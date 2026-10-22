using Newtonsoft.Json.Linq;
using WordMcpServer.Logging;
using WordMcpServer.Services;

namespace WordMcpServer.Mcp.Tools
{
    /// <summary>插入表格。</summary>
    public sealed class InsertTableTool : ToolBase
    {
        public InsertTableTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_table";

        public override string Description => "插入表格（默认追加到文档末尾，可用 position 指定在某个段落前插入）。可同时提供 data 二维数组填充单元格（行数/列数以 rows/columns 为准，data 超出部分忽略）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["rows"] = Int("行数（1-200）", 1, 200),
            ["columns"] = Int("列数（1-64）", 1, 64),
            ["data"] = StrMatrix("单元格数据（二维字符串数组，可选）"),
            ["style"] = Str("表格样式名称（可选，可用 list_styles 查看；如网格型）"),
            ["position"] = Int("在该段落之前插入（1 基）；缺省追加到文档末尾", 1),
            ["auto_fit"] = Bool("是否自动适应窗口宽度，默认 true"),
        }, "rows", "columns");

        public override JObject Execute(JObject arguments)
        {
            var request = new InsertTableRequest
            {
                Rows = ToolArgumentValidator.RequireInt(arguments, "rows", 1, 200),
                Columns = ToolArgumentValidator.RequireInt(arguments, "columns", 1, 64),
                Data = ToolArgumentValidator.OptionalStringMatrix(arguments, "data"),
                Style = ToolArgumentValidator.OptionalString(arguments, "style"),
                Position = ToolArgumentValidator.OptionalInt(arguments, "position", 1),
                AutoFit = ToolArgumentValidator.OptionalBool(arguments, "auto_fit") ?? true,
            };
            Logger.Debug($"执行 insert_table，rows={request.Rows}，columns={request.Columns}");
            return JObject.FromObject(Service.InsertTable(request));
        }
    }

    /// <summary>读取表格数据。</summary>
    public sealed class GetTableDataTool : ToolBase
    {
        public GetTableDataTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "get_table_data";

        public override string Description => "读取指定表格的全部数据（二维数组）。table_index 从 1 开始（文档中表格按出现顺序编号）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
        });

        public override JObject Execute(JObject arguments)
        {
            var tableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1;
            Logger.Debug($"执行 get_table_data，tableIndex={tableIndex}");
            return JObject.FromObject(Service.GetTableData(tableIndex));
        }
    }

    /// <summary>更新表格数据（部分更新）。</summary>
    public sealed class UpdateTableDataTool : ToolBase
    {
        public UpdateTableDataTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "update_table_data";

        public override string Description => "写入表格数据（部分更新）：data 为二维数组，从第 1 行第 1 列开始覆盖对应单元格，未提供的单元格保持不变。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["data"] = StrMatrix("要写入的二维字符串数组"),
        }, "data");

        public override JObject Execute(JObject arguments)
        {
            var tableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1;
            var data = ToolArgumentValidator.OptionalStringMatrix(arguments, "data");
            if (data == null)
            {
                throw new McpToolException(McpErrorCodes.InvalidParams, "缺少必填参数：data");
            }

            Logger.Debug($"执行 update_table_data，tableIndex={tableIndex}，rows={data.Count}");
            return JObject.FromObject(Service.UpdateTableData(tableIndex, data));
        }
    }

    /// <summary>插入表格行。</summary>
    public sealed class InsertTableRowsTool : ToolBase
    {
        public InsertTableRowsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_table_rows";

        public override string Description => "在表格中插入若干行（before_row 指定插入到该行之前；缺省追加到表格末尾），可同时填充 data。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["count"] = Int("插入行数（默认 1）", 1, 100),
            ["before_row"] = Int("插入到该行之前（1 基）；缺省追加到末尾", 1),
            ["data"] = StrMatrix("新行的数据（二维数组，可选）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new InsertTableRowsRequest
            {
                TableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1,
                Count = ToolArgumentValidator.OptionalInt(arguments, "count", 1, 100) ?? 1,
                BeforeRow = ToolArgumentValidator.OptionalInt(arguments, "before_row", 1),
                Data = ToolArgumentValidator.OptionalStringMatrix(arguments, "data"),
            };
            Logger.Debug($"执行 insert_table_rows，tableIndex={request.TableIndex}，count={request.Count}");
            return JObject.FromObject(Service.InsertTableRows(request));
        }
    }

    /// <summary>删除表格行。</summary>
    public sealed class DeleteTableRowTool : ToolBase
    {
        public DeleteTableRowTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_table_row";

        public override string Description => "删除表格中的指定行（row_index 从 1 开始）。表格只剩一行时拒绝删除。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["row_index"] = Int("要删除的行（1 基）", 1),
        }, "row_index");

        public override JObject Execute(JObject arguments)
        {
            var tableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1;
            var rowIndex = ToolArgumentValidator.RequireInt(arguments, "row_index", 1);
            Logger.Debug($"执行 delete_table_row，tableIndex={tableIndex}，rowIndex={rowIndex}");
            return JObject.FromObject(Service.DeleteTableRow(tableIndex, rowIndex));
        }
    }

    /// <summary>合并单元格。</summary>
    public sealed class MergeTableCellsTool : ToolBase
    {
        public MergeTableCellsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "merge_table_cells";

        public override string Description => "合并表格中从 (start_row, start_column) 到 (end_row, end_column) 的矩形区域为单个单元格。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["start_row"] = Int("起始行（1 基）", 1),
            ["start_column"] = Int("起始列（1 基）", 1),
            ["end_row"] = Int("结束行（1 基，含）", 1),
            ["end_column"] = Int("结束列（1 基，含）", 1),
        }, "start_row", "start_column", "end_row", "end_column");

        public override JObject Execute(JObject arguments)
        {
            var tableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1;
            var startRow = ToolArgumentValidator.RequireInt(arguments, "start_row", 1);
            var startColumn = ToolArgumentValidator.RequireInt(arguments, "start_column", 1);
            var endRow = ToolArgumentValidator.RequireInt(arguments, "end_row", 1);
            var endColumn = ToolArgumentValidator.RequireInt(arguments, "end_column", 1);
            Logger.Debug($"执行 merge_table_cells，tableIndex={tableIndex}，({startRow},{startColumn})-({endRow},{endColumn})");
            return JObject.FromObject(Service.MergeTableCells(tableIndex, startRow, startColumn, endRow, endColumn));
        }
    }

    /// <summary>设置单元格样式。</summary>
    public sealed class SetTableCellStyleTool : ToolBase
    {
        public SetTableCellStyleTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_table_cell_style";

        public override string Description => "设置表格单元格样式（未提供的参数保持不变）：不传 row/column 时应用到整个表格，传 row 或 column 时应用到对应行/列，两者都传则只作用于单个单元格。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["row"] = Int("目标行（1 基，可选）", 1),
            ["column"] = Int("目标列（1 基，可选）", 1),
            ["bold"] = Bool("加粗"),
            ["italic"] = Bool("斜体"),
            ["font_name"] = Str("字体名称"),
            ["font_size"] = Num("字号（磅）", 1, 400),
            ["color"] = Str("字体颜色：#RRGGBB / #AARRGGBB / 颜色名"),
            ["shading"] = Str("单元格底纹颜色：#RRGGBB / #AARRGGBB / 颜色名"),
            ["vertical_alignment"] = Str("垂直对齐", "top", "center", "bottom"),
            ["alignment"] = Str("水平对齐", "left", "center", "right", "justify"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetTableCellStyleRequest
            {
                TableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1,
                Row = ToolArgumentValidator.OptionalInt(arguments, "row", 1),
                Column = ToolArgumentValidator.OptionalInt(arguments, "column", 1),
                Bold = ToolArgumentValidator.OptionalBool(arguments, "bold"),
                Italic = ToolArgumentValidator.OptionalBool(arguments, "italic"),
                FontName = ToolArgumentValidator.OptionalString(arguments, "font_name"),
                FontSize = ToolArgumentValidator.OptionalDouble(arguments, "font_size", 1, 400),
                ColorArgb = ToolArgumentValidator.OptionalColor(arguments, "color"),
                ShadingArgb = ToolArgumentValidator.OptionalColor(arguments, "shading"),
                VerticalAlignment = ToolArgumentValidator.OptionalString(arguments, "vertical_alignment"),
                Alignment = ToolArgumentValidator.OptionalString(arguments, "alignment"),
            };
            Logger.Debug($"执行 set_table_cell_style，tableIndex={request.TableIndex}，row={request.Row}，column={request.Column}");
            return JObject.FromObject(Service.SetTableCellStyle(request));
        }
    }

    /// <summary>设置表格布局。</summary>
    public sealed class SetTableLayoutTool : ToolBase
    {
        public SetTableLayoutTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_table_layout";

        public override string Description => "设置表格布局（未提供的参数保持不变）：整体对齐、自动调整方式、各列宽度（磅）、首行重复、行内允许跨页断行。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["alignment"] = Str("表格对齐", "left", "center", "right"),
            ["auto_fit"] = Str("自动调整", "fixed", "content", "window"),
            ["column_widths"] = NumArray("各列宽度（磅，按列顺序）"),
            ["repeat_header_row"] = Bool("首行作为标题行跨页重复"),
            ["allow_break_across_pages"] = Bool("允许行跨页断行"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetTableLayoutRequest
            {
                TableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1,
                Alignment = ToolArgumentValidator.OptionalString(arguments, "alignment"),
                AutoFit = ToolArgumentValidator.OptionalString(arguments, "auto_fit"),
                ColumnWidths = ToolArgumentValidator.OptionalDoubleArray(arguments, "column_widths"),
                RepeatHeaderRow = ToolArgumentValidator.OptionalBool(arguments, "repeat_header_row"),
                AllowBreakAcrossPages = ToolArgumentValidator.OptionalBool(arguments, "allow_break_across_pages"),
            };
            Logger.Debug($"执行 set_table_layout，tableIndex={request.TableIndex}");
            return JObject.FromObject(Service.SetTableLayout(request));
        }
    }

    /// <summary>删除整个表格。</summary>
    public sealed class DeleteTableTool : ToolBase
    {
        public DeleteTableTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_table";

        public override string Description => "删除整个表格（连同其中数据）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
        });

        public override JObject Execute(JObject arguments)
        {
            var tableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1;
            Logger.Debug($"执行 delete_table，tableIndex={tableIndex}");
            return JObject.FromObject(Service.DeleteTable(tableIndex));
        }
    }

    /// <summary>插入表格列。</summary>
    public sealed class InsertTableColumnsTool : ToolBase
    {
        public InsertTableColumnsTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "insert_table_columns";

        public override string Description => "在表格中插入若干列（before_column 指定插入到该列之前；缺省追加到表格末尾），可同时填充 data。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["count"] = Int("插入列数（默认 1）", 1, 20),
            ["before_column"] = Int("插入到该列之前（1 基）；缺省追加到末尾", 1),
            ["data"] = StrMatrix("新列的数据（二维数组，每行一个单元格，可选）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new InsertTableColumnsRequest
            {
                TableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1,
                Count = ToolArgumentValidator.OptionalInt(arguments, "count", 1, 20) ?? 1,
                BeforeColumn = ToolArgumentValidator.OptionalInt(arguments, "before_column", 1),
                Data = ToolArgumentValidator.OptionalStringMatrix(arguments, "data"),
            };
            Logger.Debug($"执行 insert_table_columns，tableIndex={request.TableIndex}，count={request.Count}");
            return JObject.FromObject(Service.InsertTableColumns(request));
        }
    }

    /// <summary>删除表格列。</summary>
    public sealed class DeleteTableColumnTool : ToolBase
    {
        public DeleteTableColumnTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "delete_table_column";

        public override string Description => "删除表格中的指定列（column_index 从 1 开始）。表格只剩一列时拒绝删除。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["column_index"] = Int("要删除的列（1 基）", 1),
        }, "column_index");

        public override JObject Execute(JObject arguments)
        {
            var tableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1;
            var columnIndex = ToolArgumentValidator.RequireInt(arguments, "column_index", 1);
            Logger.Debug($"执行 delete_table_column，tableIndex={tableIndex}，columnIndex={columnIndex}");
            return JObject.FromObject(Service.DeleteTableColumn(tableIndex, columnIndex));
        }
    }

    /// <summary>拆分单元格。</summary>
    public sealed class SplitTableCellTool : ToolBase
    {
        public SplitTableCellTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "split_table_cell";

        public override string Description => "把指定单元格拆分为 rows × columns 个单元格（用于取消合并或细分）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["row"] = Int("目标行（1 基）", 1),
            ["column"] = Int("目标列（1 基）", 1),
            ["rows"] = Int("拆分后的行数（默认 1）", 1, 50),
            ["columns"] = Int("拆分后的列数（默认 2）", 1, 50),
        }, "row", "column");

        public override JObject Execute(JObject arguments)
        {
            var tableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1;
            var row = ToolArgumentValidator.RequireInt(arguments, "row", 1);
            var column = ToolArgumentValidator.RequireInt(arguments, "column", 1);
            var rows = ToolArgumentValidator.OptionalInt(arguments, "rows", 1, 50) ?? 1;
            var columns = ToolArgumentValidator.OptionalInt(arguments, "columns", 1, 50) ?? 2;
            Logger.Debug($"执行 split_table_cell，tableIndex={tableIndex}，({row},{column}) → {rows}x{columns}");
            return JObject.FromObject(Service.SplitTableCell(tableIndex, row, column, rows, columns));
        }
    }

    /// <summary>设置行高。</summary>
    public sealed class SetRowHeightTool : ToolBase
    {
        public SetRowHeightTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_row_height";

        public override string Description => "设置表格行高：row 指定单行（缺省所有行），rule 支持 auto / at-least（默认，不小于）/ exactly（固定）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["row"] = Int("目标行（1 基，可选）；缺省所有行", 1),
            ["height"] = Num("行高（磅）；rule=auto 时可省略", 0),
            ["rule"] = Str("行高规则", "auto", "at-least", "exactly"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetRowHeightRequest
            {
                TableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1,
                Row = ToolArgumentValidator.OptionalInt(arguments, "row", 1),
                Height = ToolArgumentValidator.OptionalDouble(arguments, "height", 0),
                Rule = ToolArgumentValidator.OptionalString(arguments, "rule"),
            };
            Logger.Debug($"执行 set_row_height，tableIndex={request.TableIndex}，row={request.Row}");
            return JObject.FromObject(Service.SetRowHeight(request));
        }
    }

    /// <summary>设置表格样式。</summary>
    public sealed class SetTableStyleTool : ToolBase
    {
        public SetTableStyleTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_table_style";

        public override string Description => "为已有表格套用表格样式（可用 list_styles 查看可用样式名，如“网格表 1 浅色 - 着色 1”）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["style"] = Str("表格样式名称"),
        }, "style");

        public override JObject Execute(JObject arguments)
        {
            var tableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1;
            var style = ToolArgumentValidator.RequireString(arguments, "style");
            Logger.Debug($"执行 set_table_style，tableIndex={tableIndex}，style={style}");
            return JObject.FromObject(Service.SetTableStyle(tableIndex, style));
        }
    }

    /// <summary>设置表格边框。</summary>
    public sealed class SetTableBordersTool : ToolBase
    {
        public SetTableBordersTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_table_borders";

        public override string Description => "设置表格边框：样式（single/double/dotted/dashed/none）、粗细、颜色，可选内框线与外框线。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["style"] = Str("边框样式", "single", "double", "dotted", "dashed", "none"),
            ["width"] = Str("粗细：thin / medium / thick 或磅值"),
            ["color"] = Str("边框颜色：#RRGGBB / 颜色名"),
            ["inside"] = Bool("是否设置内框线，默认 true"),
            ["outside"] = Bool("是否设置外框线，默认 true"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SetTableBordersRequest
            {
                TableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1,
                Style = ToolArgumentValidator.OptionalString(arguments, "style") ?? "single",
                Width = ToolArgumentValidator.OptionalString(arguments, "width"),
                ColorArgb = ToolArgumentValidator.OptionalColor(arguments, "color"),
                Inside = ToolArgumentValidator.OptionalBool(arguments, "inside") ?? true,
                Outside = ToolArgumentValidator.OptionalBool(arguments, "outside") ?? true,
            };
            Logger.Debug($"执行 set_table_borders，tableIndex={request.TableIndex}");
            return JObject.FromObject(Service.SetTableBorders(request));
        }
    }

    /// <summary>设置单元格内边距。</summary>
    public sealed class SetCellPaddingTool : ToolBase
    {
        public SetCellPaddingTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "set_cell_padding";

        public override string Description => "设置单元格内边距（磅）：不传 row/column 时应用到整表，可只设 top/bottom/left/right 中的部分。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["row"] = Int("目标行（1 基，可选）", 1),
            ["column"] = Int("目标列（1 基，可选）", 1),
            ["top"] = Num("上内边距（磅）", 0),
            ["bottom"] = Num("下内边距（磅）", 0),
            ["left"] = Num("左内边距（磅）", 0),
            ["right"] = Num("右内边距（磅）", 0),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new CellPaddingRequest
            {
                TableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1,
                Row = ToolArgumentValidator.OptionalInt(arguments, "row", 1),
                Column = ToolArgumentValidator.OptionalInt(arguments, "column", 1),
                Top = ToolArgumentValidator.OptionalDouble(arguments, "top", 0),
                Bottom = ToolArgumentValidator.OptionalDouble(arguments, "bottom", 0),
                Left = ToolArgumentValidator.OptionalDouble(arguments, "left", 0),
                Right = ToolArgumentValidator.OptionalDouble(arguments, "right", 0),
            };
            Logger.Debug($"执行 set_cell_padding，tableIndex={request.TableIndex}");
            return JObject.FromObject(Service.SetCellPadding(request));
        }
    }

    /// <summary>表格排序。</summary>
    public sealed class SortTableTool : ToolBase
    {
        public SortTableTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "sort_table";

        public override string Description => "按指定列对表格排序（可指定升序/降序、是否含表头、字段类型）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["key_column"] = Int("排序列（1 基，默认 1）", 1),
            ["ascending"] = Bool("是否升序，默认 true"),
            ["has_header"] = Bool("首行是否为表头（不参与排序），默认 true"),
            ["field_type"] = Str("字段类型", "alphanumeric", "numeric", "date"),
        });

        public override JObject Execute(JObject arguments)
        {
            var request = new SortTableRequest
            {
                TableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1,
                KeyColumn = ToolArgumentValidator.OptionalInt(arguments, "key_column", 1) ?? 1,
                Ascending = ToolArgumentValidator.OptionalBool(arguments, "ascending") ?? true,
                HasHeader = ToolArgumentValidator.OptionalBool(arguments, "has_header") ?? true,
                FieldType = ToolArgumentValidator.OptionalString(arguments, "field_type") ?? "alphanumeric",
            };
            Logger.Debug($"执行 sort_table，tableIndex={request.TableIndex}，keyColumn={request.KeyColumn}");
            return JObject.FromObject(Service.SortTable(request));
        }
    }

    /// <summary>表格转文本。</summary>
    public sealed class ConvertTableToTextTool : ToolBase
    {
        public ConvertTableToTextTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "convert_table_to_text";

        public override string Description => "把表格转换为文本（单元格之间用指定分隔符分隔，默认制表符）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["table_index"] = Int("表格索引（1 基，默认 1）", 1),
            ["separator"] = Str("单元格分隔符（默认制表符 \\t）"),
        });

        public override JObject Execute(JObject arguments)
        {
            var tableIndex = ToolArgumentValidator.OptionalInt(arguments, "table_index", 1) ?? 1;
            var separator = ToolArgumentValidator.OptionalString(arguments, "separator") ?? "\t";
            Logger.Debug($"执行 convert_table_to_text，tableIndex={tableIndex}");
            return JObject.FromObject(Service.ConvertTableToText(tableIndex, separator));
        }
    }

    /// <summary>文本转表格。</summary>
    public sealed class ConvertTextToTableTool : ToolBase
    {
        public ConvertTextToTableTool(IWordService service, IMcpLogger logger) : base(service, logger) { }

        public override string Name => "convert_text_to_table";

        public override string Description => "把段落区间内的文本转换为表格：指定列数与分隔符（默认制表符）。";

        public override JObject InputSchema => Schema(new JObject
        {
            ["start_paragraph"] = Int("起始段落（1 基）", 1),
            ["end_paragraph"] = Int("结束段落（1 基，含，可选；默认同起始段）", 1),
            ["columns"] = Int("列数（可选，缺省由 Word 推断）", 1),
            ["separator"] = Str("分隔符（默认制表符 \\t）"),
        }, "start_paragraph");

        public override JObject Execute(JObject arguments)
        {
            var start = ToolArgumentValidator.RequireInt(arguments, "start_paragraph", 1);
            var end = ToolArgumentValidator.OptionalInt(arguments, "end_paragraph", 1) ?? start;
            var columns = ToolArgumentValidator.OptionalInt(arguments, "columns", 1) ?? 0;
            var separator = ToolArgumentValidator.OptionalString(arguments, "separator") ?? "\t";
            Logger.Debug($"执行 convert_text_to_table，start={start}，end={end}");
            return JObject.FromObject(Service.ConvertTextToTable(start, end, columns, separator));
        }
    }
}
