# ============================================================================
# WordMcpServer 端到端回归测试
#
# 前置条件：
#   1. Word 已启动且 WordMcpServer 插件已加载（http://127.0.0.1:5101/health 返回 ok）
#   2. 测试会新建/修改文档，请勿在测试期间编辑其他文档
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File tests\run_tests.ps1
#   powershell -ExecutionPolicy Bypass -File tests\run_tests.ps1 -BaseUrl http://127.0.0.1:5101/mcp
# ============================================================================
param(
    [string]$BaseUrl = "http://127.0.0.1:5101/mcp",
    [switch]$KeepDocument
)

$ErrorActionPreference = "Stop"
$script:RequestId = 0
$script:PassCount = 0
$script:FailCount = 0
$script:WorkDir = Join-Path $env:TEMP "word_mcp_e2e"
$script:TestDocPath = Join-Path $script:WorkDir "e2e_test.docx"
$script:TestPdfPath = Join-Path $script:WorkDir "e2e_test.pdf"
$script:TestImagePath = Join-Path $script:WorkDir "e2e_test.png"
$script:BackupPath = Join-Path $script:WorkDir "e2e_backup.docx"

function Mcp-Call([string]$Method, [hashtable]$Params) {
    $script:RequestId++
    $payload = @{ jsonrpc = "2.0"; id = $script:RequestId; method = $Method; params = $Params }
    $body = $payload | ConvertTo-Json -Depth 30 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    $raw = (Invoke-WebRequest -Uri $BaseUrl -Method Post -Body $bytes `
            -ContentType "application/json; charset=utf-8" -TimeoutSec 300 -UseBasicParsing).Content
    return $raw | ConvertFrom-Json
}

function Invoke-Tool([string]$Name, [hashtable]$Arguments) {
    $resp = Mcp-Call "tools/call" @{ name = $Name; arguments = $Arguments }
    $text = $resp.result.content[0].text
    $data = $null
    try { $data = $text | ConvertFrom-Json } catch { }
    return [pscustomobject]@{
        IsError = [bool]$resp.result.isError
        Data    = $data
        Raw     = $text
    }
}

function Assert-True([string]$Label, [bool]$Condition, [string]$Detail = "") {
    if ($Condition) {
        $script:PassCount++
        Write-Host "  [PASS] $Label" -ForegroundColor Green
    }
    else {
        $script:FailCount++
        Write-Host "  [FAIL] $Label  $Detail" -ForegroundColor Red
    }
}

function Assert-ErrorCode([string]$Label, $Result, [string]$ExpectedCode) {
    $code = $null
    if ($Result.IsError -and $Result.Data -and $Result.Data.error) { $code = $Result.Data.error.code }
    Assert-True $Label ($code -eq $ExpectedCode) "expected=$ExpectedCode actual=$code raw=$($Result.Raw)"
}

function Assert-Success([string]$Label, $Result) {
    Assert-True $Label (-not $Result.IsError) "isError=true raw=$($Result.Raw)"
}

function New-TestImage([string]$Path) {
    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object System.Drawing.Bitmap(160, 90)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([System.Drawing.Color]::CornflowerBlue)
    $graphics.Dispose()
    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
}

Write-Host "=== WordMcpServer E2E tests ===" -ForegroundColor Cyan
Write-Host "Endpoint: $BaseUrl"

# ---------------------------------------------------------------------------
Write-Host "`n[1] 协议握手与工具清单" -ForegroundColor Yellow
$init = Mcp-Call "initialize" @{
    protocolVersion = "2025-06-18"
    capabilities    = @{}
    clientInfo      = @{ name = "e2e-test"; version = "1.0" }
}
Assert-True "initialize 返回协议版本" ($init.result.protocolVersion -eq "2025-06-18")
Assert-True "initialize 返回服务器名" ($init.result.serverInfo.name -eq "word-mcp-server")

$tools = Mcp-Call "tools/list" @{}
$toolNames = @($tools.result.tools | ForEach-Object { $_.name })
Assert-True "tools/list 返回 >= 80 个工具（实际 $($toolNames.Count)）" ($toolNames.Count -ge 80)
foreach ($expected in @("get_active_document_info", "add_paragraph", "insert_table", "find_replace_text", "export_to_pdf", "set_page_setup", "get_paragraph", "add_paragraph_runs", "set_list_format", "delete_table", "set_row_height", "update_fields", "create_style", "add_watermark", "insert_footnote", "insert_endnote", "list_notes", "list_hyperlinks", "insert_cross_reference", "insert_caption", "insert_field")) {
    Assert-True "包含工具 $expected" ($toolNames -contains $expected)
}
$readOnly = @($tools.result.tools | Where-Object { $_.annotations.readOnlyHint } | ForEach-Object { $_.name })
Assert-True "只读工具带 readOnlyHint（实际 $($readOnly.Count) 个）" ($readOnly.Count -ge 15)

# ---------------------------------------------------------------------------
Write-Host "`n[2] 文档与内容" -ForegroundColor Yellow
if (-not (Test-Path $script:WorkDir)) { New-Item -ItemType Directory -Path $script:WorkDir -Force | Out-Null }

$newDoc = Invoke-Tool "new_document" @{}
Assert-Success "new_document" $newDoc
Assert-True "新文档段落数 >= 1" ($newDoc.Data.paragraph_count -ge 1)

$info = Invoke-Tool "get_active_document_info" @{}
Assert-Success "get_active_document_info" $info
Assert-True "文档信息包含 section_count" ($info.Data.section_count -ge 1)

$h1 = Invoke-Tool "add_heading" @{ text = "端到端测试文档"; level = 1 }
Assert-Success "add_heading level=1" $h1
Assert-True "标题段落样式已套用（$($h1.Data.style)）" (-not [string]::IsNullOrEmpty($h1.Data.style))

$p1 = Invoke-Tool "add_paragraph" @{ text = "第一段：这是测试内容。" }
Assert-Success "add_paragraph 追加" $p1
$p2 = Invoke-Tool "add_paragraph" @{ text = "第二段：包含关键字 ALPHA。" }
Assert-Success "add_paragraph 第二段" $p2
$p3 = Invoke-Tool "add_paragraph" @{ text = "插入的段落"; position = 2 }
Assert-Success "add_paragraph position=2" $p3
Assert-True "插入后段落索引为 2" ($p3.Data.paragraph_index -eq 2)

$pStyled = Invoke-Tool "add_paragraph" @{ text = "带样式的段落"; style = "quote" }
Assert-Success "add_paragraph style=quote" $pStyled
Assert-True "段落样式已套用" (-not [string]::IsNullOrEmpty($pStyled.Data.style))

$text = Invoke-Tool "get_document_text" @{ max_chars = 5000 }
Assert-Success "get_document_text" $text
Assert-True "文本包含标题" ($text.Data.text -like "*端到端测试文档*")
Assert-True "文本包含插入段落" ($text.Data.text -like "*插入的段落*")

$outline = Invoke-Tool "get_document_outline" @{}
Assert-Success "get_document_outline" $outline
Assert-True "大纲至少 1 个标题" ($outline.Data.heading_count -ge 1)

# ---------------------------------------------------------------------------
Write-Host "`n[3] 段落与文本格式" -ForegroundColor Yellow
$style = Invoke-Tool "set_text_style" @{
    paragraph_index = 2; bold = $true; italic = $true
    font_size = 16; color = "#FF0000"; font_name = "Arial"
}
Assert-Success "set_text_style（段落）" $style
Assert-True "样式作用段落数=1" ($style.Data.styled_count -eq 1)

$fmt = Invoke-Tool "set_paragraph_format" @{
    paragraph_index = 2; alignment = "center"
    space_after = 12; line_spacing = "1.5"; keep_with_next = $true
}
Assert-Success "set_paragraph_format" $fmt
Assert-True "对齐已生效（$($fmt.Data.alignment)）" ($fmt.Data.alignment -eq "center")

$styleByText = Invoke-Tool "set_text_style" @{ find_text = "ALPHA"; italic = $true; color = "blue" }
Assert-Success "set_text_style（find_text）" $styleByText
Assert-True "按文本匹配到 >= 1 处" ($styleByText.Data.styled_count -ge 1)

$applyStyle = Invoke-Tool "apply_style" @{ paragraph_index = 4; style = "quote" }
Assert-Success "apply_style quote" $applyStyle

$styles = Invoke-Tool "list_styles" @{}
Assert-Success "list_styles" $styles
Assert-True "样式列表非空" ($styles.Data.count -gt 0)

# ---------------------------------------------------------------------------
Write-Host "`n[4] 表格" -ForegroundColor Yellow
$table = Invoke-Tool "insert_table" @{
    rows = 3; columns = 3
    data = @(
        @("姓名", "数量", "备注"),
        @("苹果", "10", "红富士"),
        @("香蕉", "20", "进口")
    )
}
Assert-Success "insert_table 3x3" $table
Assert-True "表格索引=1" ($table.Data.table_index -eq 1)
Assert-True "表格行数=3" ($table.Data.rows -eq 3)

$tableData = Invoke-Tool "get_table_data" @{ table_index = 1 }
Assert-Success "get_table_data" $tableData
Assert-True "单元格(1,1)=姓名" ($tableData.Data.data[0][0] -eq "姓名")
Assert-True "单元格(3,2)=20" ($tableData.Data.data[2][1] -eq "20")

$updated = Invoke-Tool "update_table_data" @{ table_index = 1; data = ,@("苹果", "15", "更新后") }
Assert-Success "update_table_data 部分更新" $updated
if ($updated.Data -and $updated.Data.data) {
    Assert-True "更新后(1,2)=15" ($updated.Data.data[0][1] -eq "15")
    Assert-True "未提供行保持不变(2,2)=10" ($updated.Data.data[1][1] -eq "10")
}

$rowsAdded = Invoke-Tool "insert_table_rows" @{ table_index = 1; count = 1; before_row = 2; data = ,@("橙子", "5", "新增行") }
Assert-Success "insert_table_rows before_row=2" $rowsAdded
Assert-True "插入后行数=4" ($rowsAdded.Data.rows -eq 4)

$cellStyle = Invoke-Tool "set_table_cell_style" @{ table_index = 1; row = 1; bold = $true; shading = "#D9E2F3"; alignment = "center" }
Assert-Success "set_table_cell_style" $cellStyle

$layout = Invoke-Tool "set_table_layout" @{ table_index = 1; alignment = "center"; auto_fit = "window"; column_widths = @(120, 80, 160) }
Assert-Success "set_table_layout" $layout

$merged = Invoke-Tool "merge_table_cells" @{ table_index = 1; start_row = 1; start_column = 1; end_row = 1; end_column = 2 }
Assert-Success "merge_table_cells" $merged

$rowDeleted = Invoke-Tool "delete_table_row" @{ table_index = 1; row_index = 4 }
Assert-Success "delete_table_row" $rowDeleted
Assert-True "删除后行数=3" ($rowDeleted.Data.rows -eq 3)

$tableAtPos = Invoke-Tool "insert_table" @{ rows = 2; columns = 2; position = 1 }
Assert-Success "insert_table position=1" $tableAtPos
Assert-True "新表格索引=1（原有表格后移）" ($tableAtPos.Data.table_index -eq 1)
Assert-True "表格总数=2" ($tableAtPos.Data.table_count -eq 2)

# ---------------------------------------------------------------------------
Write-Host "`n[5] 查找替换与书签" -ForegroundColor Yellow
$replace = Invoke-Tool "find_replace_text" @{ find_text = "ALPHA"; replace_text = "BETA"; replace_all = $true }
Assert-Success "find_replace_text" $replace
Assert-True "替换次数 >= 1" ($replace.Data.replacement_count -ge 1)

$found = Invoke-Tool "find_text" @{ text = "BETA"; max_results = 10 }
Assert-Success "find_text" $found
Assert-True "找到 BETA" ($found.Data.match_count -ge 1)

$bookmark = Invoke-Tool "insert_bookmark" @{ name = "e2e_bm"; text = "书签内容" }
Assert-Success "insert_bookmark" $bookmark
$bmRead = Invoke-Tool "get_bookmark_text" @{ name = "e2e_bm" }
Assert-Success "get_bookmark_text" $bmRead
Assert-True "书签文本匹配" ($bmRead.Data.text -eq "书签内容")
$bmDelete = Invoke-Tool "delete_bookmark" @{ name = "e2e_bm" }
Assert-Success "delete_bookmark" $bmDelete
$bmMissing = Invoke-Tool "get_bookmark_text" @{ name = "e2e_bm" }
Assert-ErrorCode "删除后书签不存在" $bmMissing "BOOKMARK_NOT_FOUND"

$link = Invoke-Tool "insert_hyperlink" @{ text = "示例链接"; address = "https://example.com"; screen_tip = "提示" }
Assert-Success "insert_hyperlink" $link
Assert-True "超链接计数 >= 1" ($link.Data.hyperlink_count -ge 1)

# ---------------------------------------------------------------------------
Write-Host "`n[6] 图片 / 目录 / 页眉页脚 / 批注" -ForegroundColor Yellow
New-TestImage $script:TestImagePath
$image = Invoke-Tool "insert_image" @{ path = $script:TestImagePath; width = 160; alignment = "center" }
Assert-Success "insert_image" $image
Assert-True "图片高度等比缩放 > 0" ($image.Data.height -gt 0)

$toc = Invoke-Tool "insert_toc" @{ levels = 3; title = "目录" }
Assert-Success "insert_toc" $toc
Assert-True "目录数量 >= 1" ($toc.Data.toc_count -ge 1)
$tocUpdate = Invoke-Tool "update_toc" @{}
Assert-Success "update_toc" $tocUpdate

$hf = Invoke-Tool "set_header_footer" @{ section_index = 1; header_text = "测试页眉"; footer_text = "测试页脚"; include_page_number = $true }
Assert-Success "set_header_footer" $hf
$hfRead = Invoke-Tool "get_header_footer" @{ section_index = 1 }
Assert-Success "get_header_footer" $hfRead
Assert-True "页眉文本匹配" ($hfRead.Data.header_text -eq "测试页眉")
Assert-True "页脚文本包含" ($hfRead.Data.footer_text -like "*测试页脚*")

$comment = Invoke-Tool "insert_comment" @{ text = "这是一条批注"; anchor_text = "BETA" }
Assert-Success "insert_comment" $comment
$comments = Invoke-Tool "list_comments" @{}
Assert-Success "list_comments" $comments
Assert-True "批注数 >= 1" ($comments.Data.count -ge 1)
$commentDelete = Invoke-Tool "delete_comment" @{ index = 1 }
Assert-Success "delete_comment" $commentDelete

# ---------------------------------------------------------------------------
Write-Host "`n[7] 页面设置 / 分页 / 光标 / 撤销" -ForegroundColor Yellow
$pageSetup = Invoke-Tool "set_page_setup" @{
    section_index = 1; orientation = "landscape"; paper_size = "A4"
    top_margin = 72; bottom_margin = 72; left_margin = 90; right_margin = 90
}
Assert-Success "set_page_setup" $pageSetup
$pageSetupRead = Invoke-Tool "get_page_setup" @{ section_index = 1 }
Assert-Success "get_page_setup" $pageSetupRead
Assert-True "方向=landscape" ($pageSetupRead.Data.orientation -eq "landscape")
Assert-True "纸张=A4" ($pageSetupRead.Data.paper_size -eq "A4")
Assert-True "上边距=72" ([math]::Abs($pageSetupRead.Data.top_margin - 72) -lt 0.5)

$pageBreak = Invoke-Tool "insert_page_break" @{ text = "分页后的内容" }
Assert-Success "insert_page_break" $pageBreak

$sectionBreak = Invoke-Tool "insert_section_break" @{ text = "新节内容" }
Assert-Success "insert_section_break" $sectionBreak
$sectionSetup = Invoke-Tool "get_page_setup" @{ section_index = 2 }
Assert-Success "读取第 2 节页面设置" $sectionSetup
Assert-True "分节后节数 >= 2" ($sectionSetup.Data.section_count -ge 2)

$cursor = Invoke-Tool "insert_text_at_cursor" @{ text = "光标插入内容" }
Assert-Success "insert_text_at_cursor" $cursor

$undo = Invoke-Tool "undo_last" @{ count = 1 }
Assert-Success "undo_last" $undo

$stats = Invoke-Tool "get_document_statistics" @{}
Assert-Success "get_document_statistics" $stats
Assert-True "页数 >= 1" ($stats.Data.pages -ge 1)
Assert-True "字数 >= 1" ($stats.Data.words -ge 1)

# ---------------------------------------------------------------------------
Write-Host "`n[8] P0 新增工具（段落回读/富文本/列表/表格闭环/域/属性/样式/水印/修订/脚注）" -ForegroundColor Yellow

# 段落回读
$readPara = Invoke-Tool "add_paragraph" @{ text = "段落回读测试内容" }
Assert-Success "add_paragraph（回读用）" $readPara
$pd = Invoke-Tool "get_paragraph" @{ paragraph_index = $readPara.Data.paragraph_index }
Assert-Success "get_paragraph" $pd
Assert-True "段落详情含文本" ($pd.Data.text -eq "段落回读测试内容")
Assert-True "段落详情含字号" ($pd.Data.font_size -gt 0)

# 富文本 runs
$runs = Invoke-Tool "add_paragraph_runs" @{ runs = @(
    @{ text = "加粗片段"; bold = $true },
    @{ text = "斜体片段"; italic = $true; color = "#C00000" },
    @{ text = "普通片段" }
) }
Assert-Success "add_paragraph_runs" $runs
$runsIndex = $runs.Data.paragraph_index
$textAfterRuns = Invoke-Tool "get_document_text" @{ max_chars = 50000 }
Assert-True "富文本内容写入" ($textAfterRuns.Data.text -like "*加粗片段斜体片段普通片段*")

# 列表格式
$list = Invoke-Tool "set_list_format" @{ type = "number"; paragraph_index = $runsIndex }
Assert-Success "set_list_format number" $list
$listRead = Invoke-Tool "get_paragraph" @{ paragraph_index = $runsIndex }
Assert-True "列表类型=number" ($listRead.Data.list_type -eq "number")
$listNone = Invoke-Tool "set_list_format" @{ type = "none"; paragraph_index = $runsIndex }
Assert-Success "set_list_format none" $listNone
$listRead2 = Invoke-Tool "get_paragraph" @{ paragraph_index = $runsIndex }
Assert-True "列表已取消" ($listRead2.Data.list_type -eq "none")

# 段落底纹 + 边框
$shade = Invoke-Tool "set_paragraph_shading" @{ paragraph_index = $runsIndex; shading = "#FFF2CC"; border_color = "#BFBFBF"; border_style = "single"; border_width = "thin" }
Assert-Success "set_paragraph_shading" $shade

# 高亮
$highlight = Invoke-Tool "set_text_style" @{ find_text = "BETA"; highlight = "yellow" }
Assert-Success "set_text_style highlight" $highlight
Assert-True "高亮命中 >= 1" ($highlight.Data.styled_count -ge 1)

# 删除文本（指定第 1 处）
$delText = Invoke-Tool "delete_text" @{ text = "普通片段"; occurrence = 1 }
Assert-Success "delete_text" $delText
Assert-True "删除 1 处" ($delText.Data.deleted_count -eq 1)

# 删除段落
$delPara = Invoke-Tool "delete_paragraph" @{ paragraph_index = $runsIndex }
Assert-Success "delete_paragraph" $delPara
Assert-True "删除后段落数减少" ($delPara.Data.paragraph_count -lt $textAfterRuns.Data.paragraph_count)

# 表格闭环
$newTable = Invoke-Tool "insert_table" @{ rows = 2; columns = 2; data = @(@("A", "B"), @("C", "D")) }
Assert-Success "insert_table（闭环测试）" $newTable
$tIndex = $newTable.Data.table_index
$addCol = Invoke-Tool "insert_table_columns" @{ table_index = $tIndex; count = 1; data = @(@("X"), @("Y")) }
Assert-Success "insert_table_columns" $addCol
Assert-True "插入列后 3 列" ($addCol.Data.columns -eq 3)
$delCol = Invoke-Tool "delete_table_column" @{ table_index = $tIndex; column_index = 3 }
Assert-Success "delete_table_column" $delCol
Assert-True "删除列后 2 列" ($delCol.Data.columns -eq 2)
$splitCell = Invoke-Tool "split_table_cell" @{ table_index = $tIndex; row = 1; column = 1; rows = 1; columns = 2 }
Assert-Success "split_table_cell" $splitCell
$rowHeight = Invoke-Tool "set_row_height" @{ table_index = $tIndex; row = 1; height = 40; rule = "exactly" }
Assert-Success "set_row_height" $rowHeight
$tableStyle = Invoke-Tool "set_table_style" @{ table_index = $tIndex; style = "网格表 1 浅色 - 着色 1" }
Assert-Success "set_table_style" $tableStyle
$tablesBefore = (Invoke-Tool "get_active_document_info" @{}).Data.table_count
$delTable = Invoke-Tool "delete_table" @{ table_index = $tIndex }
Assert-Success "delete_table" $delTable
Assert-True "删除后表格数减少" ($delTable.Data.table_count -lt $tablesBefore)

# 域更新
$fields = Invoke-Tool "update_fields" @{}
Assert-Success "update_fields" $fields
Assert-True "fields_updated=true" ($fields.Data.fields_updated -eq $true)

# 文档属性
$setProps = Invoke-Tool "set_document_properties" @{
    title = "E2E 测试文档"; author = "WordMcpServer"; subject = "端到端测试"; keywords = "mcp,test"
    custom = @{ project = "e2e"; stage = "p0" }
}
Assert-Success "set_document_properties" $setProps
Assert-True "标题已写入" ($setProps.Data.title -eq "E2E 测试文档")
$getProps = Invoke-Tool "get_document_properties" @{}
Assert-Success "get_document_properties" $getProps
Assert-True "作者已写入" ($getProps.Data.author -eq "WordMcpServer")
Assert-True "自定义属性已写入" ($getProps.Data.custom.project -eq "e2e")

# 样式管理
$createStyle = Invoke-Tool "create_style" @{ name = "E2E 样式"; type = "paragraph"; font_name = "微软雅黑"; font_size = 14; bold = $true; alignment = "center" }
Assert-Success "create_style" $createStyle
Assert-True "新样式名匹配" ($createStyle.Data.name -eq "E2E 样式")
$modifyStyle = Invoke-Tool "modify_style" @{ name = "E2E 样式"; color = "#C00000"; space_after = 8 }
Assert-Success "modify_style" $modifyStyle
$deleteStyle = Invoke-Tool "delete_style" @{ name = "E2E 样式" }
Assert-Success "delete_style" $deleteStyle

# 水印
$watermark = Invoke-Tool "add_watermark" @{ text = "内部资料"; layout = "diagonal"; transparency = 0.6 }
Assert-Success "add_watermark" $watermark
Assert-True "水印 applied=true" ($watermark.Data.applied -eq $true)
$watermarkRemove = Invoke-Tool "remove_watermark" @{}
Assert-Success "remove_watermark" $watermarkRemove
Assert-True "水印已移除 >= 1" ($watermarkRemove.Data.removed_count -ge 1)

# 修订跟踪
$track = Invoke-Tool "set_track_changes" @{ enabled = $true }
Assert-Success "set_track_changes on" $track
Assert-True "track_changes=true" ($track.Data.track_changes -eq $true)
Invoke-Tool "add_paragraph" @{ text = "修订测试段落" } | Out-Null
$accept = Invoke-Tool "accept_revisions" @{}
Assert-Success "accept_revisions" $accept
Invoke-Tool "set_track_changes" @{ enabled = $false } | Out-Null
$trackOff = Invoke-Tool "set_track_changes" @{ enabled = $false }
Assert-True "track_changes 已关闭" ($trackOff.Data.track_changes -eq $false)

# 脚注
$footnote = Invoke-Tool "insert_footnote" @{ text = "脚注内容测试" }
Assert-Success "insert_footnote" $footnote
Assert-True "脚注数 >= 1" ($footnote.Data.note_count -ge 1)

# ---------------------------------------------------------------------------
Write-Host "`n[9] 论文增强（脚注/尾注管理、内部链接、交叉引用、题注、域）" -ForegroundColor Yellow

# 脚注（锚点文本 + 管理）
$anchorPara = Invoke-Tool "add_paragraph" @{ text = "角标锚点测试文本：数据来源见注。" }
Assert-Success "add_paragraph（锚点）" $anchorPara
$fn2 = Invoke-Tool "insert_footnote" @{ text = "脚注管理测试"; anchor_text = "角标锚点测试文本" }
Assert-Success "insert_footnote（anchor_text）" $fn2
Assert-True "脚注引用段落匹配" ($fn2.Data.reference_paragraph_index -eq $anchorPara.Data.paragraph_index)
$notes = Invoke-Tool "list_notes" @{ note_type = "footnote" }
Assert-Success "list_notes footnote" $notes
Assert-True "脚注数 >= 2" ($notes.Data.count -ge 2)
$note = Invoke-Tool "get_note" @{ note_type = "footnote"; index = $fn2.Data.index }
Assert-Success "get_note" $note
Assert-True "脚注内容匹配" ($note.Data.text -eq "脚注管理测试")
$noteOpts = Invoke-Tool "set_note_options" @{ note_type = "footnote"; number_style = "lower-roman"; numbering_rule = "restart-section"; starting_number = 1 }
Assert-Success "set_note_options" $noteOpts
Assert-True "编号样式=lower-roman" ($noteOpts.Data.number_style -eq "lower-roman")

# 交叉引用（脚注编号）
$xrefNote = Invoke-Tool "insert_cross_reference" @{ reference_type = "footnote"; reference_item = [string]$fn2.Data.index; paragraph_index = $anchorPara.Data.paragraph_index }
Assert-Success "insert_cross_reference footnote" $xrefNote

# 尾注
$en = Invoke-Tool "insert_endnote" @{ text = "尾注管理测试"; paragraph_index = $anchorPara.Data.paragraph_index }
Assert-Success "insert_endnote" $en
$endnotes = Invoke-Tool "list_notes" @{ note_type = "endnote" }
Assert-Success "list_notes endnote" $endnotes
Assert-True "尾注数 >= 1" ($endnotes.Data.count -ge 1)
$delEndnote = Invoke-Tool "delete_note" @{ note_type = "endnote"; index = $en.Data.index }
Assert-Success "delete_note endnote" $delEndnote

# 书签内部链接 + 超链接管理
$bm2 = Invoke-Tool "insert_bookmark" @{ name = "xref_bm"; text = "内部链接目标" }
Assert-Success "insert_bookmark（内部链接）" $bm2
$internalLink = Invoke-Tool "insert_hyperlink" @{ text = "跳转到书签"; bookmark = "xref_bm" }
Assert-Success "insert_hyperlink（bookmark）" $internalLink
$links = Invoke-Tool "list_hyperlinks" @{}
Assert-Success "list_hyperlinks" $links
$internalIndex = $internalLink.Data.index
$internal = @($links.Data.hyperlinks | Where-Object { $_.index -eq $internalIndex })
Assert-True "内部链接已列出" ($internal.Count -eq 1)
if ($internal.Count -eq 1) {
    Assert-True "内部链接 sub_address=xref_bm" ($internal[0].sub_address -eq "xref_bm")
}
$updatedLink = Invoke-Tool "update_hyperlink" @{ index = $internalIndex; text = "跳转到书签（已改）" }
Assert-Success "update_hyperlink" $updatedLink
Assert-True "超链接文本已更新" ($updatedLink.Data.text -eq "跳转到书签（已改）")
$delLink = Invoke-Tool "delete_hyperlinks" @{ index = $internalIndex }
Assert-Success "delete_hyperlinks index" $delLink
Assert-True "删除 1 个超链接" ($delLink.Data.deleted_count -eq 1)

# 交叉引用（书签）
$xrefBm = Invoke-Tool "insert_cross_reference" @{ reference_type = "bookmark"; reference_item = "xref_bm"; paragraph_index = $anchorPara.Data.paragraph_index }
Assert-Success "insert_cross_reference bookmark" $xrefBm

# 富文本上标/下标
$supSub = Invoke-Tool "add_paragraph_runs" @{ runs = @(
    @{ text = "引用标记"; superscript = $true },
    @{ text = "[1]" },
    @{ text = "；化学式 H" },
    @{ text = "2"; subscript = $true },
    @{ text = "O" }
) }
Assert-Success "add_paragraph_runs（上标/下标）" $supSub
$supText = Invoke-Tool "get_document_text" @{ max_chars = 50000 }
Assert-True "上标/下标内容写入" ($supText.Data.text.Contains("引用标记[1]"))

# 题注 + 交叉引用 + 通用域
$caption = Invoke-Tool "insert_caption" @{ label = "图"; title = "测试示意图"; position = "below"; paragraph_index = $supSub.Data.paragraph_index }
Assert-Success "insert_caption" $caption
Assert-True "题注标签已解析" (-not [string]::IsNullOrEmpty($caption.Data.label))
$xrefCap = Invoke-Tool "insert_cross_reference" @{ reference_type = "caption"; caption_label = "图"; reference_item = "1"; reference_kind = "label-and-number"; paragraph_index = $caption.Data.paragraph_index }
Assert-Success "insert_cross_reference caption" $xrefCap
$field = Invoke-Tool "insert_field" @{ field_code = "SEQ 图 \* ARABIC"; paragraph_index = $caption.Data.paragraph_index }
Assert-Success "insert_field SEQ" $field
Assert-True "域已插入" ($field.Data.field_count -ge 1)
$fields2 = Invoke-Tool "update_fields" @{}
Assert-Success "update_fields（论文增强）" $fields2

# 删除脚注
$delNote = Invoke-Tool "delete_note" @{ note_type = "footnote"; index = $fn2.Data.index }
Assert-Success "delete_note footnote" $delNote
$notesAfter = Invoke-Tool "list_notes" @{ note_type = "footnote" }
Assert-True "删除后脚注数减少" ($notesAfter.Data.count -lt $notes.Data.count)

# ---------------------------------------------------------------------------
Write-Host "`n[10] 错误路径" -ForegroundColor Yellow
$noText = Invoke-Tool "add_paragraph" @{}
Assert-ErrorCode "add_paragraph 缺 text" $noText "INVALID_PARAMS"

$badTable = Invoke-Tool "get_table_data" @{ table_index = 99 }
Assert-ErrorCode "get_table_data 越界" $badTable "TABLE_NOT_FOUND"

$badParagraph = Invoke-Tool "set_paragraph_format" @{ paragraph_index = 9999; alignment = "center" }
Assert-ErrorCode "set_paragraph_format 段落越界" $badParagraph "PARAGRAPH_OUT_OF_RANGE"

$badImage = Invoke-Tool "insert_image" @{ path = "C:\not_exists_9f8e7d.png" }
Assert-ErrorCode "insert_image 文件不存在" $badImage "FILE_NOT_FOUND"

$noTarget = Invoke-Tool "set_text_style" @{ bold = $true }
Assert-ErrorCode "set_text_style 缺目标" $noTarget "INVALID_PARAMS"

# ---------------------------------------------------------------------------
Write-Host "`n[11] 保存 / 备份 / 导出 / 打开 / 关闭" -ForegroundColor Yellow
if (Test-Path $script:TestDocPath) { Remove-Item $script:TestDocPath -Force }
if (Test-Path $script:TestPdfPath) { Remove-Item $script:TestPdfPath -Force }
if (Test-Path $script:BackupPath) { Remove-Item $script:BackupPath -Force }

$save = Invoke-Tool "save_document" @{ save_as_path = $script:TestDocPath; format = "docx" }
Assert-Success "save_document 另存为 docx" $save
Assert-True "docx 文件已生成" (Test-Path $script:TestDocPath)
Assert-True "文档 saved=true" ($save.Data.saved -eq $true)

$backup = Invoke-Tool "backup_document" @{ backup_path = $script:BackupPath }
Assert-Success "backup_document" $backup
Assert-True "备份文件已生成" (Test-Path $script:BackupPath)

$pdf = Invoke-Tool "export_to_pdf" @{ output_path = $script:TestPdfPath }
Assert-Success "export_to_pdf" $pdf
Assert-True "PDF 文件已生成" (Test-Path $script:TestPdfPath)
Assert-True "PDF 大小 > 0" ($pdf.Data.size_bytes -gt 0)

$close = Invoke-Tool "close_document" @{ save_changes = $false }
Assert-Success "close_document" $close

$docsAfterClose = Invoke-Tool "list_documents" @{}
if ($docsAfterClose.Data.count -eq 0) {
    $noDoc = Invoke-Tool "get_active_document_info" @{}
    Assert-ErrorCode "关闭后无活动文档" $noDoc "DOC_NOT_OPEN"
}

$open = Invoke-Tool "open_document" @{ path = $script:TestDocPath }
Assert-Success "open_document" $open
Assert-True "打开文档名匹配" ($open.Data.name -like "*e2e_test*")

$docs = Invoke-Tool "list_documents" @{}
Assert-Success "list_documents" $docs
Assert-True "至少 1 个打开文档" ($docs.Data.count -ge 1)

# activate_document 与 apply_style all=true 在独立文档上验证
$newDoc2 = Invoke-Tool "new_document" @{}
Assert-Success "new_document（第二文档）" $newDoc2
Invoke-Tool "add_paragraph" @{ text = "AllOne" } | Out-Null
Invoke-Tool "add_paragraph" @{ text = "AllTwo" } | Out-Null
$allStyle = Invoke-Tool "apply_style" @{ all = $true; style = "normal" }
Assert-Success "apply_style all=true" $allStyle
Assert-True "应用到全部段落（>=2）" ($allStyle.Data.applied_count -ge 2)
$activate = Invoke-Tool "activate_document" @{ index = 1 }
Assert-Success "activate_document index=1" $activate
$docsNow = Invoke-Tool "list_documents" @{}
Assert-True "activate 后第一个文档为活动" ($docsNow.Data.documents[0].active -eq $true)
Invoke-Tool "close_document" @{ name = $newDoc2.Data.name; save_changes = $false } | Out-Null

if (-not $KeepDocument) {
    Invoke-Tool "close_document" @{ save_changes = $false } | Out-Null
    Invoke-Tool "new_document" @{} | Out-Null
}

# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "=== 结果：$script:PassCount 通过，$script:FailCount 失败 ===" -ForegroundColor $(if ($script:FailCount -eq 0) { "Green" } else { "Red" })
if ($script:FailCount -gt 0) { exit 1 }
exit 0
