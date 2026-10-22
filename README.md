# WordMcpServer — Word MCP 服务器（VSTO 插件）

一个 Word VSTO 插件。插件随 Word 启动时自动在 `127.0.0.1:5101` 运行本地 HTTP 服务器，
通过 MCP（Model Context Protocol）的 Streamable HTTP 传输暴露 81 个操作 Word 的工具。
AI 客户端（Claude Desktop / opencode 等）可以直接连接，实时操控当前打开的文档，
所有修改立即在 Word 界面中可见。

> 本项目与 `PptMcpServer`（PowerPoint 版，端口 5100）架构一致、端口错开，可同时运行。
> 开发经验见 `DEVELOPMENT_GUIDE.md`。

## 功能一览（81 个工具）

### 文档（18）

| 工具 | 说明 |
| --- | --- |
| `get_active_document_info` | 当前文档信息（文件名、路径、保存状态、段落/表格/节/页/字/字符数） |
| `list_documents` | 列出所有已打开文档（索引、名称、路径、是否活动） |
| `new_document` | 新建空白文档 |
| `open_document` | 按路径打开文档（可只读） |
| `activate_document` | 按名称/索引切换活动文档 |
| `close_document` | 关闭文档（可选保存修改） |
| `save_document` | 保存 / 另存为（docx / doc / rtf / txt / odt） |
| `backup_document` | 备份：已保存文档直接复制文件，未保存文档另存到备份路径 |
| `export_to_pdf` | 导出 PDF（ExportAsFixedFormat） |
| `get_document_text` | 读取全文纯文本（可限制最大字符数） |
| `get_document_outline` | 读取标题大纲（段落索引、级别、文本、样式） |
| `get_document_statistics` | 页数 / 字数 / 字符数 / 段落数 / 行数 |
| `update_fields` | 更新正文域、页眉页脚域与目录 |
| `get_document_properties` / `set_document_properties` | 读取 / 设置文档属性（标题、作者、主题、关键词、备注、自定义属性） |
| `set_track_changes` | 开启 / 关闭修订跟踪 |
| `accept_revisions` / `reject_revisions` | 接受 / 拒绝所有修订 |

### 段落与文本（16）

| 工具 | 说明 |
| --- | --- |
| `add_paragraph` | 追加段落（或插入到指定段落前），可套用样式 |
| `add_heading` | 添加 Heading 1-9 标题（级别与界面语言无关） |
| `add_paragraph_runs` | 富文本段落：同一段落内多段不同格式（字体/颜色/粗斜体/高亮） |
| `set_paragraph_format` | 对齐 / 缩进 / 段前段后 / 行距 / 分页控制 / 大纲级别 |
| `set_text_style` | 字体 / 字号 / 粗斜体 / 下划线 / 颜色 / 突出显示 / 上下标 |
| `get_paragraph` | 读取单段完整信息（含字体、缩进、行距、列表等格式回读） |
| `delete_paragraph` / `delete_text` | 删除段落 / 按文本删除（可指定第 N 处） |
| `set_list_format` | 项目符号 / 编号 / 多级列表（可指定段落区间与级别） |
| `set_paragraph_shading` | 段落底纹与四周边框（提示框效果） |
| `insert_page_break` / `insert_section_break` | 插入分页符 / 分节符 |
| `apply_style` | 对段落 / 匹配文本 / 整篇文档应用样式 |
| `list_styles` | 列出可用样式（内置别名与本地化名称） |
| `insert_text_at_cursor` | 在光标位置插入文本（Selection.TypeText） |
| `undo_last` | 撤销最近 N 次操作（每个 MCP 写操作是一个撤销单元） |

### 表格（14）

| 工具 | 说明 |
| --- | --- |
| `insert_table` | 插入表格（行列、二维数据、样式、位置、自适应） |
| `get_table_data` / `update_table_data` | 读取 / 部分更新表格数据 |
| `insert_table_rows` / `delete_table_row` | 插入行（可在指定行前） / 删除行 |
| `insert_table_columns` / `delete_table_column` | 插入列（可在指定列前） / 删除列 |
| `delete_table` | 删除整个表格 |
| `merge_table_cells` / `split_table_cell` | 合并 / 拆分单元格 |
| `set_table_cell_style` | 单元格字体 / 颜色 / 底纹 / 对齐（整表 / 行 / 列 / 单元格） |
| `set_table_layout` | 表格对齐 / 自动调整 / 列宽 / 标题行重复 / 跨页断行 |
| `set_row_height` | 行高（auto / at-least / exactly） |
| `set_table_style` | 为已有表格套用表格样式 |

### 查找、书签与链接（9）

| 工具 | 说明 |
| --- | --- |
| `find_replace_text` | 全文查找替换（大小写 / 全字匹配，可只替换第一处） |
| `find_text` | 查找所有出现位置（段落索引与字符位置） |
| `insert_bookmark` / `get_bookmark_text` / `delete_bookmark` | 插入 / 读取 / 删除书签 |
| `insert_hyperlink` | 插入超链接（外部 URL 或内部书签锚点） |
| `list_hyperlinks` | 列出全部超链接（地址、内部锚点、所在段落） |
| `update_hyperlink` | 修改链接显示文本 / 地址 / 锚点 / 悬停提示 |
| `delete_hyperlinks` | 按索引删除或删除全部超链接 |

### 插入、批注、水印与页面设置（13）

| 工具 | 说明 |
| --- | --- |
| `insert_image` | 插入内嵌图片（宽高可等比缩放、对齐） |
| `insert_toc` / `update_toc` | 插入 / 更新自动目录 |
| `set_header_footer` / `get_header_footer` | 设置 / 读取页眉页脚（可附加页码域） |
| `insert_comment` / `list_comments` / `delete_comment` | 插入 / 列出 / 删除批注 |
| `insert_footnote` | 插入脚注（真角标；可在指定文本后插入） |
| `add_watermark` / `remove_watermark` | 添加 / 删除文字水印（斜向或水平） |
| `get_page_setup` / `set_page_setup` | 读取 / 设置页面设置（纸张、方向、页边距、分栏） |

### 论文增强（8）

| 工具 | 说明 |
| --- | --- |
| `insert_endnote` | 插入尾注（可在指定文本后插入角标） |
| `list_notes` / `get_note` / `delete_note` | 列出 / 读取 / 删除脚注与尾注 |
| `set_note_options` | 注释编号样式、重编规则、起始编号、脚注位置 |
| `insert_cross_reference` | 交叉引用：标题 / 书签 / 脚注 / 尾注 / 题注（可点击） |
| `insert_caption` | 图/表自动编号题注（SEQ 域，支持中英文标签） |
| `insert_field` | 通用域：PAGEREF / SEQ / NOTEREF / DATE / FILENAME 等 |

### 样式管理（3）

| 工具 | 说明 |
| --- | --- |
| `create_style` | 创建自定义样式（paragraph / character / table / list） |
| `modify_style` | 修改现有样式的字体与段落格式 |
| `delete_style` | 删除自定义样式（内置样式受保护） |

## 目录结构

```
WordMcpServer.sln
build.ps1                        # 一键构建脚本（还原/构建/注册/启动/测试）
WordMcpServer/
├── WordMcpServer.csproj         # VSTO 项目（.NET Framework 4.8，packages.config）
├── packages.config              # NuGet 依赖（ASP.NET Core 2.2 + Newtonsoft.Json）
├── app.config                   # 绑定重定向（System.IO.Pipelines）
├── ThisAddIn.cs                 # 插件入口：启动/关闭 Kestrel、组装依赖
├── ThisAddIn.Designer.cs        # VSTO 生成代码（含 Globals/Factory）
├── ThisAddIn.Designer.xml       # VSTO Blueprint
├── WordMcpServer_TemporaryKey.pfx  # 清单签名测试证书（空密码）
├── Properties/AssemblyInfo.cs
├── Logging/                     # 文件日志（%APPDATA%\WordMcpServer\Logs）
├── Mcp/
│   ├── McpSettings.cs           # 配置（端口等，mcp.config.json）
│   ├── McpErrorCodes.cs         # 业务错误码
│   ├── Threading/               # Word 主线程封送器（核心！见下文）
│   ├── Protocol/                # MCP JSON-RPC 处理（initialize/tools/list/tools/call）
│   ├── Server/                  # Kestrel 自托管 + /mcp 中间件（含 SSE）
│   └── Tools/                   # 81 个工具：Schema 声明 + 参数校验 + 执行
└── Services/                    # Word 对象模型服务层（IWordService / WordService）
tests/
└── run_tests.ps1                # 端到端回归测试（214 项断言）
tools/
└── nuget.exe                    # packages.config 还原工具
```

## 环境要求

- Windows + Word 2013 及以上（推荐 2016 / 2019 / Microsoft 365）
- Visual Studio 2019 / 2022 / 2026（安装 **Office/SharePoint 开发** 工作负载，含 VSTO 工具与 Office PIA）
- .NET Framework 4.8 开发者包（VS 自带）
- MCP 客户端（opencode / Claude Desktop 等）

## 构建与运行

### Visual Studio

1. 用 Visual Studio 打开 `WordMcpServer.sln`（NuGet 包会自动还原）。
2. 选择 **Debug | Any CPU**，按 **F5**。VS 会启动 Word 并自动注册/加载插件，
   插件启动后 Kestrel 监听 `http://127.0.0.1:5101/mcp`。

### 命令行构建（已在本机验证）

```powershell
# 1) 还原 packages.config（msbuild /t:Restore 对 packages.config 无效）
tools\nuget.exe restore WordMcpServer.sln

# 2) 构建（VS 版本自动探测）
& "D:\VisualStudioCommunity\MSBuild\Current\Bin\MSBuild.exe" WordMcpServer.sln /p:Configuration=Debug
```

### 一键构建脚本（build.ps1）

`build.ps1` 封装了完整开发流程：**清理 → 还原 → 构建 → 注册 → 启动 Word → 健康检查 → 可选回归测试**。
脚本自动定位 MSBuild（vswhere + 常见路径），处理 Word 进程占用，注册插件（VSTOInstaller，
失败时回退 `|vstolocal` 手工注册），清除 Word 禁用记录，并从 `mcp.config.json` 读取端口做健康检查。

```powershell
# 只构建（要求 Word 已关闭，否则脚本会提示）
powershell -ExecutionPolicy Bypass -File build.ps1

# 构建 + 注册 + 启动 Word + 等待 MCP 就绪
powershell -ExecutionPolicy Bypass -File build.ps1 -StopWord -Register -StartWord

# 开发闭环：构建 → 注册 → 启动 → 跑完整回归测试（214 项）
powershell -ExecutionPolicy Bypass -File build.ps1 -StopWord -Register -StartWord -Test

# Release 构建（先清理）
powershell -ExecutionPolicy Bypass -File build.ps1 -Configuration Release -Clean -Register
```

| 参数 | 说明 |
| --- | --- |
| `-Configuration Debug\|Release` | 构建配置（默认 Debug） |
| `-Clean` | 构建前清理 `bin/<配置>` 与 `obj/<配置>` |
| `-NoRestore` | 跳过 NuGet 还原（packages 已存在时可用） |
| `-Register` | 注册/启用插件并清除 Word 禁用记录 |
| `-StopWord` | 构建前强制结束 WINWORD.EXE（**未保存文档会丢失**） |
| `-StartWord` | 构建后启动 Word 并等待 `/health` 就绪 |
| `-Test` | 启动后运行 `tests\run_tests.ps1`（214 项断言） |

### 注册插件（命令行构建后）

Debug 配置下 MSBuild 会生成 `WordMcpServer.vsto`，需要注册到 Word 才会加载。

**方式 A：VSTOInstaller（推荐，自动管理 ClickOnce 安装）**

```powershell
& "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" `
  /Install "file:///E:/word-plugin-mcp/WordMcpServer/bin/Debug/WordMcpServer.vsto" /Silent
```

**方式 B：手工注册表 + `|vstolocal`（跳过 ClickOnce，适合频繁重新构建）**

```powershell
$key = "HKCU:\SOFTWARE\Microsoft\Office\Word\Addins\WordMcpServer"
New-Item -Path $key -Force | Out-Null
Set-ItemProperty $key -Name "Description"   -Value "WordMcpServer"
Set-ItemProperty $key -Name "FriendlyName"  -Value "WordMcpServer"
Set-ItemProperty $key -Name "LoadBehavior"  -Value 3 -Type DWord
Set-ItemProperty $key -Name "Manifest" `
  -Value "file:///E:/word-plugin-mcp/WordMcpServer/bin/Debug/WordMcpServer.vsto|vstolocal"
```

> **重新构建后加载失败的排障（实测踩坑）**：
> - 报 `AddInAlreadyInstalledException`（"已安装另一版本的加载项，无法从该位置更新"）：
>   说明 ClickOnce 缓存里已有旧安装（例如先构建过 Release 并注册）。先卸载旧版本再加载：
>   `VSTOInstaller.exe /Uninstall "file:///.../bin/Release/WordMcpServer.vsto"`；
>   日常开发建议用方式 B 的 `|vstolocal` 直接加载本地清单，避免 ClickOnce 版本冲突。
> - 加载失败后 Word 会把插件写入禁用列表（注册表
>   `HKCU\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems`），
>   并可能把 `LoadBehavior` 改为 2；清理该键并把 `LoadBehavior` 改回 3 再重启 Word。
> - 更新插件后必须**重启 Word**（VSTO 无法热加载）；更新前先保存文档。
> - 部署到其他电脑时，VSTO 清单的自签名证书不被信任，需要把 `WordMcpServer_TemporaryKey.pfx`
>   导出的 `.cer` 导入目标机（管理员 PowerShell）：
>   `certutil -addstore -f Root X.cer` + `certutil -addstore -f TrustedPublisher X.cer`，
>   或改用企业/公共 CA 证书重新签名。

## 配置（可选）

插件目录下会自动生成 `mcp.config.json`：

```json
{
  "Host": "127.0.0.1",
  "Port": 5101,
  "LogDirectory": null
}
```

- `Host`：出于安全考虑**只能**是 `127.0.0.1`（代码强制收敛为回环地址，禁止远程连接）。
- `Port`：1-65535（默认 5101，与 PptMcpServer 的 5100 错开），修改后需重启 Word 生效。
- `LogDirectory`：日志目录，默认 `%APPDATA%\WordMcpServer\Logs\wordmcp.log`。

## 客户端配置

### opencode（`opencode.json`）

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "word": {
      "type": "remote",
      "url": "http://127.0.0.1:5101/mcp",
      "enabled": true
    }
  }
}
```

### Claude Desktop（`%APPDATA%\Claude\claude_desktop_config.json`）

```json
{
  "mcpServers": {
    "word": {
      "type": "http",
      "url": "http://127.0.0.1:5101/mcp"
    }
  }
}
```

> 工具数量变化后需重启客户端（客户端会缓存 tools/list）。

## 手动测试

```powershell
# 健康检查
curl http://127.0.0.1:5101/health

# 握手
curl -X POST http://127.0.0.1:5101/mcp -H "Content-Type: application/json" -d "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"initialize\",\"params\":{}}"

# 调用工具：新建文档并添加标题
curl -X POST http://127.0.0.1:5101/mcp -H "Content-Type: application/json" -d "{\"jsonrpc\":\"2.0\",\"id\":2,\"method\":\"tools/call\",\"params\":{\"name\":\"new_document\",\"arguments\":{}}}"
curl -X POST http://127.0.0.1:5101/mcp -H "Content-Type: application/json" -d "{\"jsonrpc\":\"2.0\",\"id\":3,\"method\":\"tools/call\",\"params\":{\"name\":\"add_heading\",\"arguments\":{\"text\":\"Hello\",\"level\":1}}}"
```

也可用官方 [MCP Inspector](https://github.com/modelcontextprotocol/inspector) 连接调试。

## 端到端回归测试

```powershell
# Word 已启动且插件已加载（/health 返回 ok）
powershell -ExecutionPolicy Bypass -File tests\run_tests.ps1
```

覆盖：协议握手、工具清单、文档读写、段落格式、富文本、列表、段落底纹、表格增删改合并、
查找替换、书签、超链接（含内部锚点）、图片、目录、页眉页脚、批注、脚注/尾注、
题注与交叉引用、水印、修订、样式管理、文档属性、域更新、页面设置、分页分节、
光标输入、撤销、保存/备份/导出 PDF/打开/关闭，以及 5 类错误路径（共 214 项断言）。

## 架构与线程模型（必读）

```
Kestrel 线程池线程                      Word 主 UI 线程（STA）
┌──────────────────────┐   Send()     ┌─────────────────────────────┐
│ McpHttpMiddleware    │ ───────────► │ ThisAddIn_Startup 安装的     │
│  → McpProtocolProcessor            │ WindowsFormsSynchronization │
│  → Tool → WordService              │ Context（隐藏控件锚点）      │
│       │ _marshaller.Invoke(...)    │  → Word COM 调用             │
└──────────────────────┘              └─────────────────────────────┘
```

- Word 的对象模型必须在**主 UI 线程**调用，否则会抛 COM 错误（RPC_E_WRONG_THREAD）。
- VSTO 环境下 `SynchronizationContext.Current` 默认为 `null`，插件启动时
  （`OfficeThreadMarshaller.Install`，此时代码正运行在主线程）安装一个
  `WindowsFormsSynchronizationContext` 并创建隐藏控件锚点 —— Word 自己的消息泵
  会处理投递给该控件的消息，从而把 Kestrel 线程池线程的回调"搬"回主线程。
- `OfficeThreadMarshaller.Invoke<T>` 用 `Send` 同步封送（请求/响应语义）；
  若调用方已在主线程则直接执行，避免死锁。异常通过 `ExceptionDispatchInfo`
  原样重抛（保留 COMException 等原始类型与堆栈）。

## 撤销与保存注意事项（必读）

- **每个写操作是一个撤销单元**：服务层用 `Application.UndoRecord.StartCustomRecord`
  把多步 COM 操作合并，`undo_last` 可以整批撤销；在 Word 界面按 Ctrl+Z 同样有效。
- **修改立即写入内存中的文档**：插件不会自动保存。重要操作前建议先调用
  `backup_document` 或手动 Ctrl+S。
- **构建/测试前先保存**：构建测试会重启 Word 进程，未保存的修改会丢失。
- **只操作"当前激活"文档**：多文档场景请先用 `list_documents` / `activate_document`
  确认目标文档，或通过 `name` 参数显式指定。
- **更新/重装插件后需重启 Word**（VSTO 无法热加载）。

## 版本说明：为什么用 ASP.NET Core 2.2 而不是 6.0

VSTO 插件必须运行在 .NET Framework 4.8 进程中（WINWORD.EXE 无法承载 .NET 5+ 运行时），
而 ASP.NET Core 6.0 的程序集依赖 .NET 6 运行时，在 .NET Framework 进程中加载会直接抛
FileLoadException。ASP.NET Core 官方支持 .NET Framework 的最后一个版本是 2.2
（Kestrel 可在 net48 上运行），因此本项目使用 ASP.NET Core 2.2 + Kestrel 自托管。

## 常见问题

- **端口被占用**：修改 `mcp.config.json` 的 `port` 并重启 Word；日志中会有监听失败记录。
- **"当前没有打开的 Word 文档"**：先新建/打开文档再调用工具。
- **工具在会话里看不到 / 工具数不对**：客户端缓存了 tools/list，重启客户端。
- **Word 里看不到新版本插件**：VSTO 未热加载，保存文档后重启 Word。
- **导出 PDF 卡住**：确认默认打印机可用（Word 分页依赖打印驱动）；本机验证
  `export_to_pdf` 正常（< 2 秒）。
- **中文写入显示乱码**：控制台编码问题，实际写入正常；用读工具回读验证。
- **AI 连不上**：确认 Word 已启动且插件加载成功（看日志 `wordmcp.log` 的
  "HTTP 服务已启动"），用 `/health` 排查。

## 安全

- 服务器只监听 `127.0.0.1`，不接受任何远程连接（代码强制）。
- 请求体限制 1MB；参数经过 JSON Schema 声明 + 运行时强校验。
- 所有操作仅作用于当前已打开的文档，不会读取文档以外的文件
  （`open_document` / `insert_image` / `backup_document` / `export_to_pdf`
  仅使用客户端明确给出的路径，且校验文件/目录存在性）。
