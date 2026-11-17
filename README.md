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
core/      报表引擎：解析、计算、渲染（不依赖 UI）
shell/     桌面 / 网页 / 手机外壳
assets/    图片、字体
work/      中间产物（不进仓库）
outputs/   交付物（不进仓库）
```

## 日常存档

```powershell
git add -A
git commit -m "feat: 说明改了什么"
git push
```

远程：`ssh://git@ssh.github.com:443/Xilin-Tan/ChinaMobileApp_Xilin_1.git`