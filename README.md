# ChinaMobileApp_Xilin_1

个人自用的报表生成工具：把**固定结构**的 CSV/Excel 数据，
在本地自动生成为**固定模板**的可视化报表。

> 不使用大模型，纯确定性转换。毫秒级出结果，0 token。

## 先读这些

| 文件 | 内容 |
|---|---|
| `AGENTS.md` | 项目须知（AI 每次会话必读） |
| `docs/00-需求.md` | 要做什么、不做什么 |
| `docs/01-技术决策.md` | 已定的方案与理由 |
| `docs/02-环境与工具.md` | 环境现状、踩过的坑 |
| `docs/03-待确认.md` | 还没定的问题 |

## 目录

```
docs/      项目文档（进仓库）
core/      报表引擎：读 xlsx、重算公式、写回（不依赖 UI）
app/       Windows 桌面界面 + 批处理流程
tools/     诊断与冒烟测试（不参与发布）
dist/      编译产物 ReportTool.exe（不进仓库）
work/      中间产物（不进仓库）
outputs/   交付物（不进仓库）
```

## 怎么用

1. 双击 `dist\ReportTool.exe`
2. 把当天的 xlsx 拖进上方虚线框（可一次拖三个，也可以拖整个文件夹）
3. 等进度跑完，结果表出现在右边，文件在 `dist\outputs\日期\`

输出文件保留原模板的全部格式（合并单元格、配色、条件格式、列宽），
只是把公式格换成了算好的数值。

## 怎么编译

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

只需要 Windows 自带的 .NET Framework 4.8，不用装任何东西。

## 日常存档

```powershell
git add -A
git commit -m "feat: 说明改了什么"
git push
```

远程：`ssh://git@ssh.github.com:443/Xilin-Tan/ChinaMobileApp_Xilin_1.git`
