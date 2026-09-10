# VSTO + MCP 插件开发经验总结

> 本文档基于 **PptMcpServer**（PowerPoint MCP 插件，53 个工具）的完整开发历程总结，
> 面向"开发一个 Word 自动操作插件（VSTO + MCP 服务器）"的场景。
> 所有 PIA 签名均在 Office 15.0 PIA（本机 Office 16）上反射/实测确认过；
> Word 侧 API 已在本机 `Microsoft.Office.Interop.Word.dll` 上确认。
>
> 阅读建议：第 1-3 章是必读的架构决策；第 4-5 章是工程与线程模板（可直接抄）；
> 第 6 章是 PIA 踩坑清单（能省掉几天调试）；第 12 章是 Word 迁移的具体对照。

---

## 1. 总体架构

```
┌─────────────────────────── POWERPNT.EXE / WINWORD.EXE（.NET Framework 4.8 进程）───────────────────────────┐
│                                                                                                            │
│  VSTO Add-in（ThisAddIn）                                                                                  │
│   ├── Startup: 安装 UI 线程同步上下文 → 组装 DI → 启动 Kestrel（127.0.0.1:5100）                          │
│   └── Shutdown: 优雅停止                                                                                    │
│                                                                                                            │
│  ┌───────────── Kestrel 线程池线程 ─────────────┐        ┌────── Office 主 UI 线程（STA）──────┐           │
│  │ HTTP 中间件（/mcp，JSON-RPC + 可选 SSE）      │        │  WindowsFormsSynchronizationContext │           │
│  │  → 协议处理器（initialize/tools/list/call）   │ Send() │   → PowerPointService / WordService │           │
│  │  → 工具（Schema + 校验）                      │ ─────► │   → COM 对象模型调用                 │           │
│  └──────────────────────────────────────────────┘        └─────────────────────────────────────┘           │
└────────────────────────────────────────────────────────────────────────────────────────────────────────────┘
                                          ▲ HTTP (127.0.0.1)
                              Claude Desktop / opencode / MCP Inspector
```

分层（各层职责单一，便于移植到 Word）：

| 层 | 类 | 职责 |
| --- | --- | --- |
| 入口 | `ThisAddIn` | 生命周期、组装服务、启动/停止 HTTP |
| 线程 | `OfficeThreadMarshaller` | 把任意线程的调用封送到 Office 主线程 |
| HTTP | `McpHttpHost` / `Startup` / `McpHttpMiddleware` | Kestrel 自托管、路由、SSE |
| 协议 | `McpProtocolProcessor` | JSON-RPC、initialize/tools/list/tools/call、错误映射 |
| 工具 | `ITool` 实现 ×N + `ToolRegistry` | 参数校验、JSON Schema、调用服务层 |
| 服务 | `IPowerPointService` / `PowerPointService` | COM 操作、DTO、业务错误码 |
| 基础设施 | `McpSettings` / `FileMcpLogger` / `McpToolException` | 配置、日志、错误 |

---

## 2. 技术选型（以及被否决的方案）

| 决策 | 选择 | 原因 / 被否决方案 |
| --- | --- | --- |
| 插件形态 | **VSTO Add-in（.NET Framework 4.8）** | VSTO 是进程内插件，可直接访问 `Application` 对象且与 UI 实时同步。**否决**独立进程 + COM 自动化：Office 不注册 ROT，无法可靠附加到已运行实例 |
| HTTP 服务器 | **ASP.NET Core 2.2 + Kestrel（自托管）** | .NET Framework 进程**无法加载 .NET 6+ 的 ASP.NET Core 程序集**（FileLoadException）。2.2 是官方支持 netfx 的最后版本。**否决** .NET 6/8（VSTO 无法运行在 .NET Core）；**否决** HttpListener（需自行实现路由/SSE，且不满足规格） |
| JSON | **Newtonsoft.Json 13** | netfx 上最稳；`JObject`/`JArray` 手写协议灵活 |
| MCP 传输 | **Streamable HTTP（POST /mcp）** | 无状态 JSON-RPC 即可满足 Claude Desktop/opencode；可选 SSE 用 `Accept: text/event-stream` 单事件包装 |
| 配置 | 插件目录下 `mcp.config.json` | VSTO 的 app.config 在 ClickOnce 部署下不可靠；自读 JSON 文件最可控 |
| 日志 | 自实现文件日志 `%APPDATA%\<插件名>\Logs` | Office 进程无控制台；日志是排障生命线 |

> **关键提醒**：不要试图在 netfx 项目里引用 net6 的 `Microsoft.AspNetCore.App`。
> 只有 ASP.NET Core 2.x 的 netstandard2.0 程序集能在 netfx 进程中加载。

---

## 3. 工程文件模板与构建

### 3.1 项目文件（.csproj）核心结构

VSTO 项目是**非 SDK 风格**的传统 csproj。与官方模板对齐（本机 VS 2026 模板路径：
`Common7\IDE\ProjectTemplates\CSharp\Office\Addins\2052\VSTOPowerPoint15AddInV4`）：

```xml
<Project ToolsVersion="18.0" DefaultTargets="Build" xmlns="...">
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props" ... />
  <PropertyGroup>
    <!-- VSTO 项目类型：Office 插件 GUID + C# GUID（顺序不能错） -->
    <ProjectTypeGuids>{BAA0C2D2-18E2-41B9-852F-F413020CAA33};{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}</ProjectTypeGuids>
    <OutputType>Library</OutputType>
    <TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
    <LoadBehavior>3</LoadBehavior>                <!-- 3 = 随宿主启动 -->
    <DefineConstants>VSTO40</DefineConstants>
    <OfficeApplication>PowerPoint</OfficeApplication>   <!-- Word 项目改为 Word -->
    <SignManifests>true</SignManifests>
    <ManifestKeyFile>PptMcpServer_TemporaryKey.pfx</ManifestKeyFile>
    <ManifestCertificateThumbprint>...</ManifestCertificateThumbprint>
    <AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>
  </PropertyGroup>

  <!-- VSTO 运行时（GAC 解析，不复制到输出目录） -->
  <Reference Include="Microsoft.Office.Tools.v4.0.Framework, Version=10.0.0.0, ...">
    <Private>False</Private>
  </Reference>
  <Reference Include="Microsoft.VisualStudio.Tools.Applications.Runtime, Version=10.0.0.0, ...">
    <Private>False</Private>
  </Reference>
  <Reference Include="Microsoft.Office.Tools, Version=10.0.0.0, ..."><Private>False</Private></Reference>
  <Reference Include="Microsoft.Office.Tools.Common, Version=10.0.0.0, ..."><Private>False</Private></Reference>
  <Reference Include="Microsoft.Office.Tools.Common.v4.0.Utilities, Version=10.0.0.0, ...">
    <Private>True</Private>
  </Reference>

  <!-- Office PIA：优先 VS 共享 PIA 目录，否则回退 GAC -->
  <Reference Include="office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c">
    <EmbedInteropTypes>True</EmbedInteropTypes>
    <HintPath>$(OfficePiaRoot)\office.dll</HintPath>
  </Reference>
  <!-- Word 项目引用 Microsoft.Office.Interop.Word，PowerPoint 项目引用 .PowerPoint -->
  <Reference Include="Microsoft.Office.Interop.PowerPoint, Version=15.0.0.0, ...">
    <EmbedInteropTypes>True</EmbedInteropTypes>
    <HintPath>$(OfficePiaRoot)\Microsoft.Office.Interop.PowerPoint.dll</HintPath>
  </Reference>
  <Reference Include="stdole, Version=7.0.3300.0, ..."><Private>False</Private></Reference>

  <Import Project="$(MSBuildToolsPath)\Microsoft.CSharp.targets" />
  <!-- VSTO 构建目标：路径是 $(VSToolsPath)\OfficeTools\...（不是 VSTOInstallPath！） -->
  <Import Project="$(VSToolsPath)\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets"
          Condition="'$(VSToolsPath)' != ''" />

  <!-- VS 属性页数据：没有它 VS 可能无法正确识别/调试 VSTO 项目 -->
  <ProjectExtensions>
    <VisualStudio>
      <FlavorProperties GUID="{BAA0C2D2-18E2-41B9-852F-F413020CAA33}">
        <ProjectProperties HostName="PowerPoint" HostPackage="{29A7B9D7-A7F1-4328-8EF0-6B2D1A56B2C1}"
          ApplicationType="PowerPoint" OfficeVersion="15.0" ... />
        <Host Name="PowerPoint" GeneratedCodeNamespace="PptMcpServer">
          <HostItem Name="ThisAddIn" Code="ThisAddIn.cs" CanonicalName="ThisAddIn"
            Blueprint="ThisAddIn.Designer.xml" GeneratedCode="ThisAddIn.Designer.cs" />
        </Host>
      </FlavorProperties>
    </VisualStudio>
  </ProjectExtensions>
</Project>
```

**必须的三个配套文件**（缺一不可）：
- `ThisAddIn.Designer.cs`：**必须用 VS 模板的生成代码**（含 `Globals` 类、`Factory` 构造、
  `ThisAddIn(Factory, IServiceProvider)` 构造签名）。手写旧式无参构造会导致运行时无法加载。
- `ThisAddIn.Designer.xml`：Blueprint（`<hostitem:hostItem ... baseType="Microsoft.Office.Tools.AddInBase" ...>`），
  VS 用它校验/重新生成 Designer.cs。
- `ThisAddIn.cs`：用户代码，`public partial class ThisAddIn` + `InternalStartup()` 绑定 Startup/Shutdown。

### 3.2 命令行构建（无 VS 界面）

```powershell
# 1) 还原 packages.config（msbuild /t:Restore 对 packages.config 无效，必须用 nuget.exe）
tools\nuget.exe restore YourAddin.sln

# 2) 构建（VS 版本自动探测，无需额外参数）
& "D:\VisualStudioCommunity\MSBuild\Current\Bin\MSBuild.exe" YourAddin.sln /p:Configuration=Debug
```

csproj 中的 VS 版本探测（命令行构建时 `VisualStudioVersion` 未注入）：

```xml
<PropertyGroup>
  <VisualStudioVersion Condition="'$(VisualStudioVersion)' == '' And Exists('$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v18.0\OfficeTools')">18.0</VisualStudioVersion>
  <VisualStudioVersion Condition="'$(VisualStudioVersion)' == '' And Exists('$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v17.0\OfficeTools')">17.0</VisualStudioVersion>
  <VisualStudioVersion Condition="'$(VisualStudioVersion)' == ''">10.0</VisualStudioVersion>
  <VSToolsPath Condition="'$(VSToolsPath)' == ''">$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v$(VisualStudioVersion)</VSToolsPath>
</PropertyGroup>
```

### 3.3 解决方案文件（.sln）的坑

- 项目行类型 GUID 用 **C# 的 `{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}`**，
  不是 VSTO 的 `{BAA0C2D2-...}`（后者只出现在 csproj 的 `ProjectTypeGuids` 里）。
  写错会导致 VS 打开提示"不兼容/不支持的项目类型"。

### 3.4 NuGet 依赖与绑定重定向（大坑）

`packages.config` 模式**不会自动还原传递依赖**，必须把全部依赖显式列入：

```xml
<packages>
  <package id="Microsoft.AspNetCore" version="2.2.0" targetFramework="net48" />
  <package id="Microsoft.AspNetCore.Hosting" version="2.2.0" targetFramework="net48" />
  <package id="Microsoft.AspNetCore.Server.Kestrel" version="2.2.0" targetFramework="net48" />
  <!-- ...共 40+ 个包（Kestrel.Core/Http.*/Extensions.*/System.*） -->
  <package id="System.IO.Pipelines" version="4.5.1" targetFramework="net48" />
</packages>
```

csproj 中还要为每个包写 `<Reference>` + `<HintPath>..\packages\...`（见本项目 csproj 的 NuGet 引用块）。

**运行时绑定重定向**（netfx 上 Kestrel 的经典问题）：

```xml
<runtime>
  <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
    <!-- AutoGenerateBindingRedirects 不覆盖 System.IO.Pipelines，需手工添加。
         注意 newVersion 是程序集的实际装配版本（netstandard2.0 版是 4.0.0.0），不是包版本！ -->
    <dependentAssembly>
      <assemblyIdentity name="System.IO.Pipelines" publicKeyToken="cc7b13ffcd2ddd51" culture="neutral" />
      <bindingRedirect oldVersion="0.0.0.0-4.0.1.1" newVersion="4.0.0.0" />
    </dependentAssembly>
  </assemblyBinding>
</runtime>
```

排查方法：异常信息 `FileLoadException: 找到的程序集清单定义与程序集引用不匹配` →
用 `[System.Reflection.AssemblyName]::GetAssemblyName("dll路径")` 看实际装配版本再写重定向。

---

## 4. 线程模型（VSTO 最核心的代码）

Office 对象模型必须在**主 UI 线程（STA）**调用。Kestrel 在线程池线程处理请求，
因此需要把调用"搬"回 UI 线程。VSTO 环境下 `SynchronizationContext.Current` 通常为
`null`，需要手工安装 WinForms 上下文 + 隐藏控件锚点：

```csharp
public sealed class OfficeThreadMarshaller : IDisposable
{
    private readonly SynchronizationContext _uiContext;
    private readonly Control _anchor;   // 保持消息泵常驻

    /// <summary>必须在 ThisAddIn_Startup（主线程）中调用。</summary>
    public static OfficeThreadMarshaller Install(IMcpLogger logger)
    {
        if (SynchronizationContext.Current == null)
        {
            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
        }

        // 隐藏控件锚点：WinForms 上下文通过控件句柄投递回调，
        // Office 自己的消息泵会处理投递过来的消息
        var anchor = new Control { Visible = false };
        anchor.CreateControl();
        return new OfficeThreadMarshaller(SynchronizationContext.Current, anchor, logger);
    }

    public bool IsOnUiThread => SynchronizationContext.Current == _uiContext;

    /// <summary>把委托封送到 Office 主线程同步执行；已在主线程则直接执行（避免死锁）。</summary>
    public TResult Invoke<TResult>(Func<TResult> func)
    {
        if (IsOnUiThread) return func();

        TResult result = default;
        ExceptionDispatchInfo captured = null;
        try
        {
            _uiContext.Send(_ =>
            {
                try { result = func(); }
                catch (Exception ex) { captured = ExceptionDispatchInfo.Capture(ex); }
            }, null);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("无法封送到 Office 主线程：UI 线程可能已退出。", ex);
        }

        captured?.Throw();   // 保留原始异常类型（COMException）与堆栈
        return result;
    }
}
```

服务层模式——**每个公开方法第一步就 Invoke**：

```csharp
public PresentationInfo GetActivePresentationInfo()
{
    return _marshaller.Invoke(() =>
    {
        var pres = GetActivePresentationOrThrow();   // 内部校验 + 友好错误
        return new PresentationInfo { Name = pres.Name, /* ... */ };
    });
}
```

注意事项：
- `Send` 会阻塞 Kestrel 线程直到 UI 完成——请求/响应语义正确，但 UI 卡住时请求会等待；
- `run_slideshow` 这类阻塞 UI 的操作会让后续调用排队（文档要向用户说明）；
- 不要在 UI 线程做长任务（Excel 图表 `ChartData.Activate` 可能耗时数秒）。

---

## 5. Office PIA 开发方法论（省时间的核心经验）

### 5.1 不要凭记忆写 API——用反射确认

PIA 的接口分层（`Presentation` / `_Presentation`）、枚举名、参数类型在不同版本差异很大。
反射确认命令模板（PowerShell + ReflectionOnlyLoad）：

```powershell
$onResolve = { param($s, $e)
  $name = $e.Name.Split(',')[0]
  $path = "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\$name.dll"
  if (!(Test-Path $path)) { $path = "$env:WINDIR\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll" }
  if (Test-Path $path) { return [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($path) }
  return $null
}
[System.AppDomain]::CurrentDomain.add_ReflectionOnlyAssemblyResolve($onResolve)
$pia = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom(
  "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll")
# 查方法签名（注意：公开接口可能没有方法，方法在 _Xxx 接口上）
($pia.GetExportedTypes() | Where-Object { $_.Name -eq "_Document" }).GetMethods() |
  Where-Object { $_.Name -eq "ExportAsFixedFormat" } | ForEach-Object { $_.ToString() }
```

**编译校验技巧**：没有完整 VS 环境时，可用 Roslyn 的 `csc.exe` 直接编译源码做语法/类型检查：
从 NuGet 下载 `Microsoft.Net.Compilers.Toolset`，用
`tasks\net472\csc.exe @response.rsp`（rsp 里列引用与源文件）编译，能提前发现 90% 的 API 错误。

### 5.2 PIA 实测踩坑清单（PowerPoint，Word 同理需要逐个确认）

| 坑 | 表现 | 正确做法 |
| --- | --- | --- |
| 强类型可选参数 | `ExportAsFixedFormat` 的可选参数是强类型枚举（不是 object），传 `Type.Missing` 编译失败 | **省略**可选参数：`pres.ExportAsFixedFormat(path, type)` |
| `AddEffect` 的 Index 参数 | 传 0/-1 报"不在 1 到 N 范围"；传合法值也不生效（实测） | 只传必填参数；需要排序时"快照→清空→重建" |
| 没有 `AddExitEffect` | 方法不存在 | `AddEffect` 后设置 `effect.Exit = msoTrue` |
| 没有 `Sequence.Remove` | 方法不存在 | 用 `Effect.Delete()` |
| `Chart.HasTitle` | 是 `bool` 不是 MsoTriState | `chart.HasTitle = true` |
| `Table.FirstRow` | 是 `bool` | `table.FirstRow = true` |
| `Slide.Background` | 返回 **ShapeRange**（不是 SlideBackground 对象，该类型不存在） | 通过 `slide.Background.Fill` 操作；`FollowMasterBackground` 在 **Slide** 上 |
| 翻转属性不可写 | 没有 `FlipH/FlipV` 属性 | 只有 `Shape.Flip(MsoFlipCmd.msoFlipHorizontal)`（翻转一次）+ 只读的 `HorizontalFlip` |
| 枚举成员名 | `msoShapeIsoscelesTriangle`（不是 Isoceles）、`PpPlaceholderType`（不是 PP）、`ppAlignLeft`（不是 ppAlignmentLeft） | 反射 `[System.Enum]::GetNames()` 确认 |
| 跨 PIA 同名类型 | `ColorFormat`/`XlLegendPosition`/`XlAxisType` 在 office.dll 与 PPT PIA 都有 | 用 `using` 别名（`using PpXlAxisType = Microsoft.Office.Interop.PowerPoint.XlAxisType;`）或显式全名 |
| 闭合形状设置箭头 | 抛"指定的值超出范围"（且不是 COMException） | try/catch 转 INVALID_PARAMS 友好提示 |
| 圆环图设置轴标题 | 抛"该对象无效"（环形图无坐标轴） | try/catch 转友好错误 |

### 5.3 异常处理模式

```csharp
catch (McpToolException) { throw; }                       // 业务错误原样抛出
catch (Exception ex) when (!(ex is McpToolException))     // 其他全部转为可读错误
{
    throw new McpToolException(McpErrorCodes.OperationFailed, "操作失败：" + ex.Message);
}
```

**不要只 catch `COMException`**：PIA 经常抛出 `ArgumentException`、`InvalidCastException`、
甚至被包装的 `TargetInvocationException`（dynamic 调用）。放宽到 `catch (Exception)` 后再
映射为业务错误，避免客户端收到无意义的 INTERNAL_ERROR。

### 5.4 COM 对象释放策略

- **只释放本方法内新建的对象**（`Shapes.AddXxx` 返回的 Shape、`Slides[i]` 等），
  `try/finally` + `Marshal.FinalReleaseComObject`，失败忽略；
- **不要释放 `ActiveDocument`/`ActivePresentation` 这类宿主持有的对象**——释放 RCW 会破坏后续访问；
- 进程内插件（VSTO）泄漏影响小于独立自动化进程，GC 兜底可接受。

---

## 6. MCP 协议层实现要点

### 6.1 必须实现的方法

| 方法 | 返回 |
| --- | --- |
| `initialize` | `protocolVersion`（协商：客户端版本在已知集合内则沿用，否则返回最新）+ `capabilities.tools` + `serverInfo` |
| `notifications/initialized` 及其他 `notifications/*` | **无响应**（HTTP 202 空体） |
| `ping` | `{}` |
| `tools/list` | 每个工具 `{name, description, inputSchema}`；只读工具加 `annotations: {readOnlyHint: true}` |
| `tools/call` | 成功 `{content:[{type:"text", text:...}], isError:false}`；**业务失败用 `isError:true` + 错误码**，而不是 JSON-RPC error |

### 6.2 关键细节

- JSON-RPC 错误码：`-32700` 解析失败、`-32600` 请求非法、`-32601` 未知方法、`-32602` 参数非法；
- `tools/call` 的业务错误格式：`content[0].text = {"error":{"code":"...","message":"..."}}` + `isError:true`；
- **图片内容**：`content` 可包含 `{type:"image", data:"<base64>", mimeType:"image/png"}` ——
  多模态模型可直接"看到"排版（截图自检）。实现方式：工具返回特殊键 `$content`（JArray），协议层原样透传；
- SSE：客户端 `Accept: text/event-stream` 时用单事件包装（`event: message\ndata: {...}\n\n`）；
- 请求日志：method + id + 耗时，是排障的第一手资料。

---

## 7. 工具设计规范（53 个工具沉淀出的约定）

1. **命名**：工具名与参数一律 `snake_case`（`slide_index`、`font_size`）；
   返回 JSON 用 `camelCase`（DTO 加 `[JsonProperty("slide_index")]`）；
2. **定位参数**：`xxx_name` 或 `xxx_id` 二选一（至少一个），找不到时返回明确的 `*_NOT_FOUND`；
3. **默认值**：位置/尺寸等给合理默认（如 left=72、top=150、width=200），单位统一用**磅**；
4. **颜色**：统一 `#RRGGBB` / 颜色名 / `theme:accent1` 主题色三种形式，服务层一个 `ApplyColor` 收敛；
5. **文本换行**：统一规范化字面量 `\n` → 真实换行（不同客户端 JSON 转义不一致，服务层兜底）；
6. **范围校验**：索引 1 基、范围检查先于 COM 调用（`SLIDE_OUT_OF_RANGE` 等友好错误）；
7. **错误码清单**：`PPT_NOT_OPEN`、`SLIDE_OUT_OF_RANGE`、`SHAPE_NOT_FOUND`、`FILE_NOT_FOUND`、
   `INVALID_PARAMS`、`UNSUPPORTED_TYPE`、`OPERATION_FAILED`、`ANIMATION_NOT_FOUND`、`INTERNAL_ERROR`；
8. **工具粒度**：一个工具做一件事但支持"部分更新"（未传的参数保持不变），减少调用次数；
9. **边界防御**：行列上限（表格 75×25）、字号 1-400、透明度 0-100、文件路径先校验存在与扩展名。

---

## 8. 部署与证书（其他电脑"证书无效"的根因）

- VSTO 清单必须签名。开发用自签名证书在**本机**可用，**其他电脑**报"证书无效/不受信任"
  （根证书不在信任存储）——与 Debug/Release 无关（两个配置用同一张证书）。
- 分发方案：
  - 测试：把 `.cer` 拷到目标机导入（管理员 PowerShell）：
    `certutil -addstore -f Root X.cer` + `certutil -addstore -f TrustedPublisher X.cer`；
  - 正式：用企业 CA / 公共 CA 代码签名证书，替换 csproj 的
    `ManifestKeyFile` + `ManifestCertificateThumbprint` 后重新构建。
- Debug 配置 VS 会自动注册（`RegisterForVSTO`）；Release 不注册，需要双击 `.vsto` 或 ClickOnce 发布；
- **更新插件后必须重启 Office**（VSTO 无法热加载）；**更新前先保存文档**（重启会丢未保存内容）。

---

## 9. 测试方法（无 GUI 自动化）

**端到端测试模式**（本项目验证 53 个工具的方式）：

1. 测试程序（或 PowerShell）通过 COM 启动独立 Office 实例 → 新建文档；
2. 通过 HTTP POST `/mcp` 调用工具，核对返回 JSON（`isError`、字段值、回读一致性）；
3. 覆盖三类用例：正常路径、回读验证（写入后用读工具核对）、错误路径（各类错误码）。

PowerShell 调用模板：

```powershell
function Mcp($id, $name, $argsJson) {
  $body = '{"jsonrpc":"2.0","id":' + $id + ',"method":"tools/call","params":{"name":"' + $name + '","arguments":' + $argsJson + '}}'
  $raw = (Invoke-WebRequest -Uri "http://127.0.0.1:5100/mcp" -Method Post -Body $body `
           -ContentType "application/json" -TimeoutSec 180 -UseBasicParsing).Content
  ($raw | ConvertFrom-Json).result.content[0].text
}
```

注意事项：
- **测试会修改当前激活文档**：先 `new_presentation` 或确认测试文档，避免动用户正在编辑的文件；
- 测试结束不要 `Quit()` 掉用户实例；独占测试时先结束 `WINWORD.EXE` 再启动干净实例；
- 断言注意响应里 text 字段是**转义的 JSON 字符串**（`\"key\"`），带引号的 `Contains` 会失败——
  用 `ConvertFrom-Json` 解析后再断言。

---

## 10. 常见问题排查表

| 现象 | 原因 | 解决 |
| --- | --- | --- |
| VS 打开提示"不兼容" | sln 项目类型 GUID 写成了 VSTO GUID | 改为 `{FAE04EC0-...}`（见 3.3） |
| 构建报找不到 VSTO targets | 导入路径错 / VS 版本未探测 | `$(VSToolsPath)\OfficeTools\...` + 版本探测（见 3.2） |
| 启动报 `FileLoadException`（某个 System.* 或 Pipelines） | 缺依赖或绑定重定向 | 补包 + 手工 bindingRedirect（见 3.4） |
| 启动报缺 `Microsoft.Extensions.FileProviders.Abstractions` 等 | packages.config 传递依赖未列全 | 从 nuspec 提取全部依赖补齐 |
| 运行时 `RPC_E_WRONG_THREAD` / COM 调用失败 | 未封送 UI 线程 | 检查所有服务方法都走 `Invoke`（见第 4 章） |
| 首次调用偶发 `PPT_NOT_OPEN` | 宿主启动竞态 | 服务层"数量为 0 时 Sleep 250ms 重试一次" |
| 中文写入变成乱码/问号 | 控制台显示问题（实际写入正常） | 用读工具回读验证，不要看控制台 |
| 工具在会话里看不到（工具数不对） | 客户端缓存了 tools/list | 重启客户端（Claude Desktop / opencode） |
| Office 里看不到新版本插件 | VSTO 未热加载 | 保存文档后重启 Office |

---

## 11. Word 插件迁移对照（本机已确认 Word 15.0 PIA）

### 11.1 工程改动清单

| 项 | PowerPoint 项目 | Word 项目 |
| --- | --- | --- |
| ProjectTypeGuids / sln | 相同（VSTO 项目类型 GUID 不变） | 相同 |
| `OfficeApplication` | `PowerPoint` | `Word` |
| PIA 引用 | `Microsoft.Office.Interop.PowerPoint` | `Microsoft.Office.Interop.Word` |
| ProjectExtensions | HostName/HostPackage 换 Word 模板值 | 参考本机模板 `VSTOWord15AddInV4\WordAddIn.csproj` |
| ThisAddIn.Designer.xml | `type="Microsoft.Office.Interop.PowerPoint.Application"` | `type="Microsoft.Office.Interop.Word.Application"` |
| ThisAddIn.cs | `this.Application` 类型为 PPT Application | 类型为 Word Application |

### 11.2 Word 关键 API（已反射确认，可直接用）

```csharp
// Application（Word）
_Application.ActiveDocument : Document      // 无文档时抛异常/返回 null，需先判断 Documents.Count
_Application.Documents : Documents
_Application.UndoRecord : UndoRecord        // ★ Word 有撤销记录对象
_Application.ScreenUpdating : Boolean       // 可写：批量操作前 false，finally 恢复 true
_Application.Selection : Selection

// Document
_Document.Save()
_Document.SaveAs2(ref object FileName, ref object FileFormat, ... 17 个 ByRef object 参数)  // 全 ByRef！
_Document.ExportAsFixedFormat(
    string OutputFileName, WdExportFormat ExportFormat, bool OpenAfterExport,
    WdExportOptimizeFor OptimizeFor, WdExportRange Range, int From, int To,
    WdExportItem Item, bool IncludeDocProps, bool KeepIRM,
    WdExportCreateBookmarks CreateBookmarks, bool DocStructureTags,
    bool BitmapMissingFonts, bool UseISO19005_1, ref object FixedFormatExtClassPtr)
_Document.Saved : Boolean (可写)           // 可手动标记保存状态
_Document.Content : Range                  // 全文范围
_Document.Tables / Paragraphs / Shapes / InlineShapes / Bookmarks

// Find（Range.Find 或 Content.Find）
Find.Execute(ref object FindText, ref object MatchCase, ref object MatchWholeWord, ... 15 个 ByRef 参数)
// 常用：find.Text = "a"; find.Replacement.Text = "b";
//       find.Execute(ref missing, ref missing, ..., Replace: ref wdReplaceAll ...)
```

### 11.3 Word 特有的实现建议

1. **撤销支持（Word 比 PPT 强）**：用 `UndoRecord` 把一批操作合并成一个撤销单元：

   ```csharp
   var undo = _app.UndoRecord;
   undo.StartCustomRecord("AI 修改");
   try { /* 一批 COM 操作 */ }
   finally { undo.EndCustomRecord(); }
   ```

2. **性能**：批量操作前 `_app.ScreenUpdating = false;`，`finally` 恢复 `true`；
   大文档遍历用 `Range` 而不是 `Selection`（Selection 会抢焦点/滚动）。

3. **表格单元格文本**：`Cell.Range.Text` 末尾带单元格结束符（`\r\a`），
   写入时用 `cell.Range.Text = "值"`，读取时需 `TrimEnd('\r','\a')`。

4. **备份方案**：Word **没有** `SaveCopyAs`。已保存文档用 `File.Copy(FullName, 备份路径)`；
   未保存文档用 `SaveAs2` 到备份路径（注意这会改变文档关联，或备份后 `SaveAs2` 回原路径）。

5. **导出 PDF**：`ExportAsFixedFormat` 参数是**强类型**（不像 PPT 的 object 可选参数），
   但必须传满 15 个参数——用 `WdExportFormat.wdExportFormatPDF` +
   `WdExportRange.wdExportAllDocument` + `WdExportItem.wdExportDocumentContent` 等枚举填充。

6. **查找替换**：`Find.Execute` 的 15 个参数全部 ByRef，C# 调用需要
   `object missing = Type.Missing; find.Execute(ref findText, ref missing, ...)`；
   或更简单：`document.Content.Find` 设置 `Text`/`Replacement.Text` 后
   `Execute(Replace: WdReplace.wdReplaceAll)`（C# 不支持命名参数调用 COM，需按位置传）。

7. **多文档**：`Application.Documents` 遍历；激活用 `document.Activate()`（Word 的 Document 有 Activate 方法）。

8. **线程**：与 PPT 完全一致（STA 主线程 + SynchronizationContext 封送），
   `OfficeThreadMarshaller` 可原样复用（把类型名换成 Word）。

### 11.4 建议的 Word 工具清单（按本项目经验排序）

| 阶段 | 工具 |
| --- | --- |
| MVP（10 个） | `get_active_document_info`、`get_document_content`（或大纲）、`add_paragraph`、`add_heading`、`set_paragraph_format`、`insert_table`、`get_table_data`、`update_table_data`、`find_replace_text`、`save_document` |
| 第二批 | `insert_image`、`set_text_frame`（字体/颜色/加粗）、`set_page_setup`、`export_to_pdf`、`backup_document`、`insert_page_break`、`set_document_style`（套用样式） |
| 第三批 | `insert_toc`（目录）、`insert_bookmark`、页眉页脚、`set_table_cell_style`、`merge_table_cells`、`set_table_layout`、`insert_hyperlink`、批注（Comments）、`list_documents`/`activate_document`/`new_document` |
| 增强 | 样式管理（Styles）、分节、水印、信封/邮件合并、`export_section_image`（Word 无单页导出，需 PDF 转图） |

---

## 12. 开发流程建议（按本项目实际节奏）

1. **先跑通最小闭环**：ThisAddIn 启动 Kestrel → `initialize` → `tools/list` → 一个 `get_*` 工具；
   验证线程封送、日志、健康检查都工作，再铺工具；
2. **反射先行**：每个新工具动手前先反射确认 PIA 签名与枚举名（第 5.1 节），
   比"写完编译报错再查"快得多；
3. **每加 3-5 个工具构建一次**，随时保持可运行；
4. **端到端回归**：维护一份 PowerShell/HTTP 测试脚本（覆盖正常+错误路径），
   每次改动后全量跑一遍（本项目 53 工具约 2 分钟）；
5. **日志是生命线**：每个工具调用记录 method/参数摘要/耗时/错误码；
6. **文档同步**：工具数量、部署步骤、保存注意事项写进 README，避免"文档漂移"。

---

## 附：本项目可直接复用的文件清单

| 文件 | 用途 | Word 复用方式 |
| --- | --- | --- |
| `Mcp/Threading/PowerPointThreadMarshaller.cs` | UI 线程封送 | 改名 `OfficeThreadMarshaller`，逻辑不变 |
| `Mcp/Server/*`（Host/Startup/Middleware） | Kestrel + HTTP + SSE | 原样复用 |
| `Mcp/Protocol/McpProtocolProcessor.cs` | JSON-RPC 协议 | 原样复用（工具注册表注入） |
| `Mcp/Tools/ITool.cs`、`ToolRegistry.cs`、`ToolArgumentValidator.cs`、`McpToolException.cs` | 工具框架 | 原样复用 |
| `Logging/*`、`Mcp/McpSettings.cs` | 日志与配置 | 改目录名即可 |
| `Services/IPowerPointService.cs` 的 DTO/错误码模式 | 服务契约 | 替换为 Word 语义的 DTO |
| `PptMcpServer.csproj` | VSTO 工程模板 | 换 `OfficeApplication`、PIA 引用、ProjectExtensions |
| `TEST_PROMPTS.md`、`README.md` | 测试提示词与文档 | 改工具名即可 |
