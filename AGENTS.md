# AGENTS.md — 项目须知（每次会话必读）

> 本文件是这个项目的**长期记忆**。上下文丢失、换新窗口、换模型之后，
> 先读本文件 + `docs/` 目录，即可恢复到完整状态。不要凭猜测行事。

## 项目是什么

个人自用工具：把**固定结构**的报表数据（CSV/Excel），在本地自动生成为**固定模板**的可视化报表。
替代"每次把 Excel 发给 AI、重新解释需求、等半天、烧 token"的做法。

目标平台：Windows 桌面优先，兼顾安卓 / 苹果手机（详见 `docs/01-技术决策.md`）。

## 铁律（不要违反）

1. **主流程不使用大模型。** 结构固定 + 数据变化 = 确定性转换。0 token、毫秒级、结果可复现。
2. **报表模板用配置文件描述**，不硬编码进代码。加一列、换图表类型只改配置。
3. **引擎与界面严格分层。** 解析 / 计算 / 渲染不依赖 UI 框架，换壳不重写引擎。
4. **校验宁可报错，不可静默算错。** 列名、类型、编码不符就明确失败。
5. **不提交真实业务数据、密钥、个人信息。** 用脱敏样例。
6. 不擅自引入新框架或依赖；需要时先说明理由再动。

## 仓库与 Git

- 仓库：`ChinaMobileApp_Xilin_1`
- 远程：`ssh://git@ssh.github.com:443/Xilin-Tan/ChinaMobileApp_Xilin_1.git`
  - ⚠️ **必须保持这个形式。** `https://github.com/...` 在当前网络会被重置（`ERR_CONNECTION_RESET`）。
  - 历史：曾用 `xilin_chinamobile_1` 仓库和 `Andy-Tan-8964` 账号，均已弃用。
- 账号：`Xilin-Tan` / `1366641201@qq.com`
- 分支：`main`。单人项目，直接提交 main，不需要 PR 流程。
- 账号已配置好 SSH 密钥（ed25519），推送无需额外操作。

## 环境事实

- Windows；Git 2.56.1 位于 `G:\Git\Git\`，已在系统 PATH。
- 技术栈已定：**.NET Framework 4.8 + C# + WinForms**，Windows 自带，
  编译器 `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`（C# 5 语法）。
  **不需要装 JDK / Flutter / Node**，见 `docs/02-环境与工具.md`。
- 编译：`powershell -ExecutionPolicy Bypass -File build.ps1` → `dist/ReportTool.exe`
- 沙箱限制：本会话的 shell 用户与仓库所有者不同，**无法写 `.git/` 和 `~/.gitconfig`**。
  - 只读查询用 `git -c safe.directory=* ...`
  - 提交 / 推送由用户在自己的 PowerShell 执行
- 用户偏好：**给一条明确的路，不要给多个备选**；命令要能整段复制粘贴。

## 文档索引

| 文件 | 内容 |
|---|---|
| `docs/00-需求.md` | 需求、范围、明确不做的部分 |
| `docs/01-技术决策.md` | 已定的方案与理由 |
| `docs/02-环境与工具.md` | 环境现状、踩过的坑与解决办法 |
| `docs/03-待确认.md` | 尚未决定的问题 |

## 代码结构（已实现）

| 路径 | 作用 |
|---|---|
| `core/` | 引擎。读 xlsx → 重算全部公式 → 写回保留格式。不依赖任何 UI |
| `core/Model.cs` | 数据结构（Workbook / Sheet / CellValue） |
| `core/XlsxReader.cs` | 零依赖 xlsx 读取（解 zip + OOXML） |
| `core/Lexer.cs` `core/Evaluator.cs` | 公式词法与求值（SUMIFS/RANK/ROUND/…） |
| `core/Calculator.cs` | 不动点迭代重算 |
| `core/XlsxWriter.cs` | 写回 xlsx，原格式一行不动 |
| `app/` | WinForms 桌面界面 + 批处理流程 |
| `app/Pipeline.cs` | 批量入口：识别报表类型 → 重算 → 输出 → 抽摘要 |
| `app/Summary.cs` | 从结果里抽取要显示在界面上的表 |
| `app/Inbox.cs` | 扫描企业微信接收目录，按内容里的通报时间挑最新，供一键取数 |
| `app/MainForm.cs` | 主窗口（拖拽区 + 日志 + 结果表） |
| `tools/` | 诊断 / 冒烟测试，不参与发布 |
| `build.ps1` | 一键编译成 `dist/ReportTool.exe`（单文件，约 50 KB） |

编译与测试：

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

## 已知的坑（实现相关）

- 写回 xlsx 时**绝不能 `XmlNode.RemoveAll()`** —— 会连 `r`/`s` 属性一起删掉，格式全丢。
  只删 `f`/`v`/`is` 子节点。
- 单元格文本里常有换行（如 `宽带（…）\n-不含未知网格`），按表头找行要先截到第一行。
- 区县表第二块（宽带）标题下方**隔一行**才是数据，偏移是 `标题行+2`，不是 `+1`。
- 重点业务表的「指标大类/小类」是合并单元格，只在首行有值，要向下继承。
- 共享公式（`<f t="shared">`）不支持，`XlsxReader` 会直接抛异常而不是算错。
- 已去掉输出文件的 `fullCalcOnLoad`，否则 Excel 打开会自己重算一遍（慢）。
- **企业微信用文件名序号 / 修改时间判断新旧都会错。** 同一天的多个 xlsx 会被批量重写成同一个
  `LastWriteTime`，序号 `(11)` `(12)` 和先后无关。必须读内容里的通报时间
  （`auto` 表「通报时间」列第 2 行，列位置会变，顺着「通报」表 A1 公式找），见 `app/Inbox.cs`。
  实测：`(12)` 是 11-19 13:45，`(11)` 才是 11-19 17:15。
- 别把整个企业微信目录直接丢给 `XlsxReader` 做全量读：里面混着 14 MB 的无关大表，读它会卡几分钟。

## 会话结束时的动作

本次若有**新决定**或**新踩的坑**，追加到对应 docs 文件，然后提醒用户提交：

```powershell
git add -A
git commit -m "docs: 记录XX"
git push
```

## 新窗口开场白（用户可直接复制）

> 先读 AGENTS.md 和 docs/ 下所有文件，然后我们继续。
