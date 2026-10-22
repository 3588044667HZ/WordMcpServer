# WordMcpServer 测试提示词集

本文档提供面向 AI 客户端的测试提示词，可直接粘贴给 Claude Desktop / opencode 等
已接入 `WordMcpServer` 的 MCP 客户端使用。每条包含：目的、提示词、预期结果/验证点。

## 测试前置条件

- Word 已启动，插件已加载（日志 `%APPDATA%\WordMcpServer\Logs\wordmcp.log` 出现
  "HTTP 服务已启动"），MCP 服务器地址 `http://127.0.0.1:5101/mcp` 可达。
- 在 Word 中**新建一个专用测试文档**并保持激活（建议先 Ctrl+S 存为 `mcp_test.docx`）。
  以下提示词中的删除/覆盖类操作**只应在该测试文档上执行**。
- 准备一张本地图片（如 `C:\Users\Administrator\Pictures\test.png`）供 insert_image 使用，
  可改用你机器上的任意图片路径。
- 所有操作均可用 `undo_last` 撤销；重要内容可先用 `backup_document` 备份。

---

## 一、冒烟测试（服务与基本读写）

### 1. 服务连通性

**提示词**：

> 请检查 WordMcpServer 是否正常工作：先调用 get_active_document_info。
> 如果返回 DOC_NOT_OPEN，提示我新建文档后重试。成功则汇报：当前文档名称、
> 段落数、页数、字数、是否已保存。

**预期**：isError=false；字段含 name、paragraph_count、page_count、word_count、saved。

### 2. 内容写入与回读一致性

**提示词**：

> 在当前文档依次执行：① add_heading 添加一级标题“测试文档 123”；
> ② add_paragraph 追加“这是第一段测试内容。”；③ add_paragraph 追加“这是第二段。”；
> ④ get_document_text 回读全文；⑤ get_document_outline 回读大纲。
> 汇报每步返回的 paragraph_index，并确认回读文本包含上述内容。

**预期**：三次写调用 isError=false；outline 返回 1 个 level=1 的标题；
get_document_text 的 text 包含三段文字。

### 3. 段落定位与插入

**提示词**：

> 调用 get_document_text 查看当前段落；用 add_paragraph 的 position 参数在第 2 段
> 之前插入“插队段落”；再次 get_document_text 确认“插队段落”位于第 2 段，
> 并汇报 paragraph_count 的变化。

**预期**：返回的 paragraph_index=2；回读文本顺序正确。

---

## 二、格式设置

### 4. 字体与段落格式

**提示词**：

> 对第 2 段执行：① set_text_style 设置字体“微软雅黑”、字号 16、加粗、颜色 #FF0000；
> ② set_paragraph_format 设置居中对齐、段后间距 12 磅、1.5 倍行距、与下段同页。
> 然后 get_document_text 回读确认内容未被破坏。

**预期**：两次调用 isError=false；set_paragraph_format 返回 alignment="center"。

### 5. 按文本匹配样式

**提示词**：

> 先用 add_paragraph 写入“重要提示 IMPORTANT 内容”；
> 再用 set_text_style 的 find_text="IMPORTANT" 对该文本应用斜体 + 蓝色；
> 汇报 styled_count。

**预期**：styled_count≥1；isError=false。

### 6. 样式应用与枚举

**提示词**：

> ① list_styles 列出当前文档样式，找出“引用”或 Quote 样式名称；
> ② apply_style 对第 3 段应用该样式；③ apply_style 用 style="heading2" 把第 4 段
> 变成二级标题；④ get_document_outline 确认新增一个 level=2 标题。

**预期**：list_styles 返回非空；apply_style 成功；outline 的 heading_count 增加。

---

## 三、表格

### 7. 创建与回读

**提示词**：

> 调用 insert_table 创建 3 行 3 列表格并填充数据：
> [[“产品”,“销量”,“增长率”],[“A”,120,“12%”],[“B”,98,“8%”]]，auto_fit=true。
> 然后用 get_table_data 回读，核对单元格值。

**预期**：table_index=1、rows=3、columns=3；回读 data[0][0]="产品"、data[1][1]="120"。

### 8. 部分更新与行操作

**提示词**：

> ① update_table_data 只写第 1 行：[[“产品”,“销量”,“备注”]]；
> ② insert_table_rows 在表格第 2 行前插入一行，数据 [[“C”,75,“-3%”]]；
> ③ get_table_data 确认行数变为 4；
> ④ delete_table_row 删除最后一行；⑤ 再次回读确认行数恢复为 3。

**预期**：每步 isError=false；行数变化符合预期；未更新的单元格保持不变。

### 9. 表格样式与合并

**提示词**：

> ① set_table_cell_style 对第 1 行设置加粗、底纹 #D9E2F3、水平居中、垂直居中；
> ② set_table_layout 设置表格居中、列宽 [120, 80, 160]、首行跨页重复；
> ③ merge_table_cells 合并第 1 行第 1-2 列；④ get_table_data 确认合并后仍可读取。

**预期**：各步 isError=false；合并后不抛异常。

---

## 四、查找替换、书签与链接

### 10. 查找替换闭环

**提示词**：

> ① find_replace_text 把“测试”替换为“验收”（replace_all=true），汇报替换次数；
> ② find_text 查找“验收”，汇报 match_count 与首个匹配的 paragraph_index；
> ③ get_document_text 确认文档中不再包含“测试”。

**预期**：替换次数≥1；match_count≥1；isError 均为 false。

### 11. 书签

**提示词**：

> ① insert_bookmark 名称 "section1"、内容“第一章要点”；
> ② get_bookmark_text 回读书签文本；③ delete_bookmark 删除；
> ④ 再次 get_bookmark_text 应返回 BOOKMARK_NOT_FOUND 错误。

**预期**：步骤 1-3 成功；步骤 4 的 isError=true 且 error.code="BOOKMARK_NOT_FOUND"。

### 12. 超链接

**提示词**：

> insert_hyperlink 添加显示文本“OpenAI”指向 https://openai.com，
> 悬停提示“官网”，插入到文档末尾；汇报 hyperlink_count。

**预期**：isError=false；hyperlink_count≥1。

---

## 五、图片、目录、页眉页脚与批注

### 13. 图片

**提示词**：

> insert_image 插入图片（用 C:\Users\Administrator\Pictures\test.png；
> 文件不存在就告诉我），宽度 200 磅、居中对齐；汇报返回的宽高与 inline_shape_count。

**预期**：isError=false；height>0（等比缩放）；inline_shape_count≥1。

### 14. 目录

**提示词**：

> 确认文档已有至少 2 个 Heading 标题后：① insert_toc 插入 3 级目录，
> 标题文字“目录”；② update_toc 更新目录；③ 汇报 toc_count。

**预期**：toc_count≥1；isError=false。

### 15. 页眉页脚

**提示词**：

> set_header_footer 设置第 1 节页眉“WordMcpServer 测试”、页脚“第 X 页”，
> include_page_number=true；然后 get_header_footer 回读并核对。

**预期**：回读 header_text="WordMcpServer 测试"；footer_text 包含“第 X 页”或页码域。

### 16. 批注

**提示词**：

> ① insert_comment 用 anchor_text="IMPORTANT" 添加批注“请复核此段”；
> ② list_comments 查看批注列表；③ delete_comment 删除第 1 条；④ 再次 list_comments。

**预期**：批注数先增后减；每步 isError=false。

---

## 六、页面设置、分页与撤销

### 17. 页面设置

**提示词**：

> ① set_page_setup 把第 1 节设为横向、A4、四边页边距 72 磅；
> ② get_page_setup 回读，核对 orientation="landscape"、paper_size="A4"、top_margin≈72。

**预期**：回读字段与设置一致（允许 ±0.5 磅误差）。

### 18. 分页、分节与光标输入

**提示词**：

> ① insert_page_break 在文档末尾分页并写入“新的一页”；
> ② insert_section_break 再插入一个分节符；
> ③ get_page_setup 读取第 2 节确认 section_count≥2；
> ④ insert_text_at_cursor 在当前光标处输入“光标文字”。

**预期**：各步 isError=false；section_count 增加。

### 19. 撤销

**提示词**：

> 先用 add_paragraph 追加“将被撤销的段落”，用 get_document_text 确认存在；
> 然后 undo_last count=1；再次 get_document_text 确认该段落消失。

**预期**：撤销后文本不再包含“将被撤销的段落”。

---

## 七、保存、备份与导出

### 20. 另存与备份

**提示词**：

> ① save_document 另存为 C:\Users\Administrator\Documents\mcp_test_saved.docx；
> ② backup_document 备份到 C:\Users\Administrator\Documents\mcp_test_backup.docx；
> ③ export_to_pdf 导出到 C:\Users\Administrator\Documents\mcp_test.pdf。
> 汇报三个文件路径与返回的 size_bytes，并确认 saved=true。

**预期**：三个文件均生成；size_bytes>0；isError=false。

### 21. 多文档切换

**提示词**：

> ① new_document 新建第二个文档；② list_documents 查看两个文档；
> ③ activate_document 按 index 切回第一个；④ 用 get_active_document_info 确认
> 活动文档已切换（核对 name）。

**预期**：list_documents 返回 2 个；activate 后 active=true 的文档索引正确。

---

## 八、错误路径（验证友好错误码）

**提示词**：

> 依次执行以下调用，并逐一汇报 error.code：
> ① add_paragraph 不传 text；② get_table_data 传 table_index=99；
> ③ set_paragraph_format 传 paragraph_index=9999；
> ④ insert_image 传不存在的路径 C:\not_exists.png；
> ⑤ set_text_style 只传 bold 不传目标；⑥ get_bookmark_text 传不存在的书签。

**预期错误码**：① INVALID_PARAMS；② TABLE_NOT_FOUND；③ PARAGRAPH_OUT_OF_RANGE；
④ FILE_NOT_FOUND；⑤ INVALID_PARAMS；⑥ BOOKMARK_NOT_FOUND。
所有错误均以 isError=true 返回，而不是 JSON-RPC error。

---

## 九、P0 新能力（富文本 / 列表 / 表格闭环 / 样式 / 水印 / 修订 / 脚注）

### 23. 富文本与列表

**提示词**：

> ① add_paragraph_runs 写入一个段落，包含三个片段：“重要”加粗红色、“说明”斜体、
> “普通”带黄色高亮；② 用 get_paragraph 回读该段落，核对 bold/italic/highlight 字段；
> ③ add_paragraph 追加三条“清单项 A/B/C”；④ set_list_format 用 start_paragraph/end_paragraph
> 把这三条设为 number 编号；⑤ get_paragraph 回读确认 list_type="number"；
> ⑥ 再 set_list_format type="none" 取消编号。

**预期**：runs 内容完整；list_type 先 number 后 none；每步 isError=false。

### 24. 段落底纹与边框

**提示词**：

> 用 set_paragraph_shading 对第 2 段设置底纹 #FFF2CC、边框颜色 #BF8F00、single 细边框，
> 然后 get_paragraph 回读确认无错误。

**预期**：isError=false；渲染上该段呈现浅黄底纹与边框。

### 25. 表格闭环

**提示词**：

> ① insert_table 建 3×2 表格；② insert_table_columns 在末尾加一列并填数据；
> ③ delete_table_column 删掉第 3 列；④ set_row_height 把表头行固定为 32 磅；
> ⑤ set_table_style 套用“网格表 1 浅色 - 着色 1”；⑥ split_table_cell 拆分一个单元格；
> ⑦ get_table_data 回读核对；⑧ delete_table 删除整个表格。

**预期**：列数 2→3→2；各行操作 isError=false；最终表格数减少。

### 26. 文档属性与域更新

**提示词**：

> ① set_document_properties 设置标题“季度报告”、作者“AI 助手”、关键词“报告,季度”，
> 并添加自定义属性 project="q3"；② get_document_properties 回读核对；
> ③ 先 add_heading 建两个标题并 insert_toc 插入目录，再 update_fields 刷新域。

**预期**：属性回读一致；update_fields 返回 fields_updated=true。

### 27. 样式管理

**提示词**：

> ① create_style 创建段落样式“提示框”（微软雅黑 14 磅、加粗、居中、段后 8 磅）；
> ② apply_style 把某段应用“提示框”；③ modify_style 把该样式颜色改为 #C00000；
> ④ list_styles 确认存在；⑤ delete_style 删除该样式。

**预期**：创建/修改/删除均成功；删除后 list_styles 中不再出现。

### 28. 水印

**提示词**：

> ① add_watermark 添加斜向水印“内部资料”，透明度 0.7，颜色 #C0C0C0；
> ② 导出 PDF 或截图确认每页都有斜向水印；③ remove_watermark 删除并确认 removed_count≥1。

**预期**：水印出现在所有节（链接页眉自动继承）；删除后消失。

### 29. 修订跟踪

**提示词**：

> ① set_track_changes 开启修订；② add_paragraph 追加“修订测试段落”；
> ③ 确认 Word 中该段带有修订标记；④ accept_revisions 接受全部修订；
> ⑤ set_track_changes 关闭；⑥ 再用 add_paragraph 追加内容并 reject_revisions 验证拒绝路径。

**预期**：开启后修改带痕迹；接受/拒绝均成功；最终关闭修订。

### 30. 脚注

**提示词**：

> 用 insert_footnote 在第 2 段末尾插入脚注“数据来源：内部统计”，
> 然后导出 PDF 确认页面底部出现脚注内容。

**预期**：footnote_count≥1；PDF 页面底部显示脚注文本。

---

## 十、论文写作增强（脚注/尾注 / 交叉引用 / 题注 / 内部链接 / 域）

### 31. 脚注与尾注全流程

**提示词**：

> ① 在“自然语言处理”这几个字之后插入脚注“实验数据来源于公开基准测试集”（用
> insert_footnote 的 anchor_text）；② list_notes 查看脚注列表；③ get_note 读第 1 条；
> ④ set_note_options 把脚注编号改为 lower-roman、每节重编；⑤ 在段末插入一条尾注并
> list_notes note_type="endnote"；⑥ delete_note 删除该尾注。

**预期**：正文出现上标角标、页底出现脚注内容；编号样式变为小写罗马数字；尾注可增删。

### 32. 题注与交叉引用

**提示词**：

> ① 在“（此处为架构示意图）”段落后 insert_caption（label="图"，title="模型核心架构示意图"）；
> ② 在上一段末尾 insert_cross_reference 引用该题注（reference_type="caption",
> caption_label="图", reference_item="1", reference_kind="label-and-number"）；
> ③ update_fields 刷新编号；④ 导出 PDF 检查“图 1”编号与引用文本一致。

**预期**：题注自动编号为“图 1 …”，交叉引用显示“图 1”；更新域后编号稳定。

### 33. 内部链接与超链接管理

**提示词**：

> ① insert_bookmark 创建书签 "ref1"；② insert_hyperlink 用 bookmark="ref1" 插入
> “返回参考文献”内部链接；③ list_hyperlinks 确认 sub_address="ref1"；
> ④ update_hyperlink 修改链接显示文本；⑤ delete_hyperlinks 删除该链接。

**预期**：链接可跳转到书签；列表能读到内部锚点；修改/删除均生效。

### 34. 上标引用与通用域

**提示词**：

> ① add_paragraph_runs 写入“准确率提升显著（p<0.05）”+ 上标“[1]”+“，详见附录。”；
> ② 插入一段“生成日期：”，用 insert_field 在该段末尾插入 DATE 域；
> ③ update_fields 后导出 PDF 检查日期已渲染。

**预期**：[1] 显示为上标；DATE 域显示当前日期。

---

## 十一、P0 增强（第三批：选区/图片/页码/修订/批注/表格/保护/封面）

### 35. 选区读取、跳转与替换

**提示词**：

> ① get_selection 查看当前选区（类型、文本、段落、页码）；② goto_page page=1 跳转到第 1 页；
> ③ 选中一段文字后 replace_selection 替换为“已由 AI 改写”；④ get_document_text 确认写入。

**预期**：选区信息含段落与页码；跳转后 page_number=1；替换内容出现在文档中。

### 36. 段落移动与区间复制

**提示词**：

> ① add_paragraph 依次追加 A、B 两段；② move_paragraph 把 A 下移 1 位，
> get_paragraph 确认顺序变为 B、A；③ copy_range 以 A 的新索引复制到文末；
> ④ get_document_text 确认 A 出现两次，且文档段落数只增加 1。

**预期**：移动后顺序正确；复制保留格式；不产生多余空段落。

### 37. 图片管理

**提示词**：

> ① insert_image 插入一张图片（layout="floating"、wrap="square"、alt_text="示意图"）；
> ② list_images 查看内嵌/浮动条目；③ update_image 按 kind/name 调整宽度、写入 alt 与标题；
> ④ 再列一次确认更新；⑤ delete_image 删除该图片。

**预期**：列表包含 kind、index/name、wrap、alt_text；更新与删除均 isError=false。

### 38. 页码、首页页眉与封面

**提示词**：

> ① set_page_numbers section_index=1 number_style="lower-roman" starting_number=5
> restart_at_section=true；② set_header_footer header_type="first-page" 写入“首页页眉”；
> ③ insert_cover_page 插入标题、副标题、作者、单位、日期；
> ④ get_document_text 确认封面标题在文档最开头；⑤ 导出 PDF 检查封面独立成页。

**预期**：页码为小写罗马数字且从 5 开始；首页页眉只出现在首页；
文档以表格开头时也能把封面插到表格之前（内部自动处理分节符）。

### 39. 单条修订处理

**提示词**：

> ① set_track_changes 开启修订；② 修改两处文本；③ list_revisions 读取修订列表；
> ④ accept_revision 接受第 1 条，再 list_revisions 确认条数减少；
> ⑤ reject_revision 拒绝剩余一条；⑥ set_track_changes 关闭。

**预期**：单条接受/拒绝均生效；修订数量递减；最终关闭修订。

### 40. 批注回复与解决

**提示词**：

> ① insert_comment 对“结论”插入批注；② get_comment 读取详情；
> ③ reply_comment 回复“已修改”；④ resolve_comment done=true；
> ⑤ get_comment 确认 done=true 且 reply_count=1；⑥ resolve_comment done=false 重新打开。

**预期**：回复数与解决状态正确；重新打开后 done=false。

### 41. 表格增强闭环

**提示词**：

> ① insert_table 建表并填入 A/10、B/20、C/30；② set_table_borders 设置边框；
> ③ set_cell_padding 设置内边距；④ sort_table 按第 2 列数字升序；
> ⑤ get_table_data 确认顺序；⑥ convert_table_to_text 用 "|" 分隔；
> ⑦ convert_text_to_table 再把该文本转回 2 列表格。

**预期**：排序正确；转换后文本包含 "A|10"；表格数先减后增。

### 42. 文档保护与最终状态

**提示词**：

> ① protect_document type="read-only"；② unprotect_document；
> ③ 先 save_document 保存文档；④ mark_as_final final=true；⑤ mark_as_final final=false。
> 另：对未保存的新文档直接 mark_as_final=true，应返回 OPERATION_FAILED 而不是卡住。

**预期**：保护/解除成功；最终状态可切换；未保存文档得到明确错误（Word 会弹“另存为”）。

### 43. 多格式保存与 PDF 范围导出

**提示词**：

> ① save_document format="html" 另存为 report.html；② 再另存为 docm 与 xml；
> ③ export_to_pdf range="from-to" from_page=1 to_page=1；
> ④ 检查文件均存在；⑤ 用不支持的格式（如 pdf）调用 save_document 应返回 UNSUPPORTED_TYPE。

**预期**：各格式文件生成；范围 PDF 只含指定页；不支持格式返回明确错误码。

---

## 十二、综合任务（端到端）

### 44. 生成一份完整报告

**提示词**：

> 新建文档，生成一份“季度工作总结”报告：
> ① 标题“2026 年第三季度工作总结”（heading1）+ 副标题段；
> ② 三个二级标题：“一、主要成果”“二、存在问题”“三、下一步计划”；
> ③ 每个标题下写 1-2 段正文（其中“主要成果”下写 3 条数据）；
> ④ 在“主要成果”下插入 4 行 3 列表格（指标/目标/完成）；
> ⑤ 给表格首行加底纹并加粗；
> ⑥ 插入 3 级目录（标题“目录”）；
> ⑦ 设置页脚“第 X 页”并加页码；
> ⑧ 导出 PDF 到 C:\Users\Administrator\Documents\quarter_report.pdf；
> ⑨ 最后用 get_document_outline 与 get_document_statistics 汇报结构。

**预期**：所有步骤成功；大纲含 4 个标题；统计返回页数与字数；PDF 文件生成。
