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
| `app/Summary.cs` | 从结果里抽取界面要显示的表（PreviewTable）和图表数据（ChartData） |
| `app/Inbox.cs` | 扫描企业微信接收目录，按内容里的通报时间挑最新，供一键取数 |
| `app/Chart.cs` | 零依赖画图（System.Drawing）：横向条形 / 分组柱状，存 PNG，多张竖向拼「一页纸」 |
| `app/MainForm.cs` | 主窗口（拖拽区 + 日志 + 图表页 + 结果表） |
| `tools/` | 诊断 / 冒烟测试，不参与发布 |
| `build.ps1` | 一键编译成 `dist/ReportTool.exe`（单文件，约 74 KB） |

编译与测试：

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

## 图表模块（app/Chart.cs）

不引第三方画图库，全部用 `System.Drawing` 手画，保证单文件、无需安装。

- `ChartData`：标题 / 副标题 / 类别 / 系列列表 / 是否百分比（`Percent`）/ 参考线
  （`Reference` + `ShowReference`）/ `Horizontal`（横向条形）/ `MaxBars` / `TailBars`。
  - `MaxBars > 0` 时只画前 N 条；`TailBars = true` 表示取**末尾** N 条（即最落后的，
    「一页纸」里的网格图就是这么截的），否则取最前面 N 条。
- `ChartSeries`：系列名 + 颜色 + `double[]`，值直接来自 `CellValue.Num`，
  **绝不从界面格式化的字符串里反解数字**。
- `ChartRenderer.Render(d, w, h)` → `Bitmap`；`Save(d, path)`（固定宽 900，高度按条目数算）；
  `SaveStack(list, path, 900)` 把多张竖向拼成一页。
- 配色规则 `RateColor`：≥100% 绿、≥80% 黄、<80% 红。
- `AxisCap(d)`：百分比图的坐标轴上限定为 `Reference*2`，超出的用箭头表示，
  免得个别 400% 的极值把整张图压扁。

出图在 `Pipeline.Run` 末尾进行，**出图失败只记日志，不影响已经算好的 Excel**。
输出目录里是 `图1_xxx.png` … 加一张 `日报_一页纸.png`。

界面侧：图表页在 `MainForm` 构造时就建好并常驻，底部一排按钮在「一页纸 / 各单图」之间切换，
默认显示第一张单图，双击图片用系统看图打开原尺寸。

## 已知的坑（实现相关）

- 写回 xlsx 时**绝不能 `XmlNode.RemoveAll()`** —— 会连 `r`/`s` 属性一起删掉，格式全丢。
  只删 `f`/`v`/`is` 子节点。
- 单元格文本里常有换行（如 `宽带（…）\n-不含未知网格`），按表头找行要先截到第一行。
- 区县表第二块（宽带）标题下方**隔一行**才是数据，偏移是 `标题行+2`，不是 `+1`。
- 重点业务表的「指标大类/小类」是合并单元格，只在首行有值，要向下继承。
- 共享公式（`<f t="shared">`）不支持，`XlsxReader` 会直接抛异常而不是算错。
- 已去掉输出文件的 `fullCalcOnLoad`，否则 Excel 打开会自己重算一遍（慢）。
- 网格数不是固定的：2026-11-19 的通报里是 **33 个**（`通报` 表第 5 行起，到空行为止），
  别照抄某个历史数字去校验，按行扫到空就停。
- **`TabControl.TabPages.Insert` 在控件句柄还没创建时会被静默丢弃**：页面会留在 `Controls` 里，
  却不进 `TabPages`，表现就是"标签页凭空消失"（`bounds` 停在默认的 200×100）。
  只用 `TabPages.Add`。因此「图表」页在构造函数里就 Add 进去，清空时用 `ClearTables()`
  只删数据表页，绝不调 `TabPages.Clear()`。
- `Image.FromStream` 会一直占用那个流，直接读文件会锁住 PNG。必须先 `ReadAllBytes` 再
  `new Bitmap(img)` 复制一份（`MainForm.ShowChart` 已这么做）。
- **企业微信用文件名序号 / 修改时间判断新旧都会错。** 同一天的多个 xlsx 会被批量重写成同一个
  `LastWriteTime`，序号 `(11)` `(12)` 和先后无关。必须读内容里的通报时间
  （`auto` 表「通报时间」列第 2 行，列位置会变，顺着「通报」表 A1 公式找），见 `app/Inbox.cs`。
  实测：`(12)` 是 11-19 13:45，`(11)` 才是 11-19 17:15。
- 别把整个企业微信目录直接丢给 `XlsxReader` 做全量读：里面混着 14 MB 的无关大表，读它会卡几分钟
  （那份大表还用了共享公式，会抛 `NotSupportedException`）。
- **读整表等于卡死。** 判断"这是哪一份报表"只用 `Inbox.QuickSheetNames`（只读 `workbook.xml`），
  实测 22 份 28 ms；`Pipeline.Peek` 已改成走它。拖入文件夹时先按前缀筛、再按通报时间取每类最新，
  整个目录 111 个文件只选出 3 份，耗时 225 ms。

## 会话结束时的动作

本次若有**新决定**或**新踩的坑**，追加到对应 docs 文件，然后提醒用户提交：

```powershell
git add -A
git commit -m "docs: 记录XX"
git push
```

## 新窗口开场白（用户可直接复制）

> 先读 AGENTS.md 和 docs/ 下所有文件，然后我们继续。
