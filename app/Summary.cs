using System;
using System.Collections.Generic;
using System.Globalization;
using ReportTool.Core;

namespace ReportTool.App
{
    /// <summary>界面上直接显示的预览表。</summary>
    public class PreviewTable
    {
        public string Title;
        public string[] Columns;
        public List<string[]> Rows = new List<string[]>();
    }

    /// <summary>
    /// 从重算后的工作簿里抽取「人最想先看的那几行」。
    /// 全部按表头文字定位，不写死坐标，模板微调也不会错位。
    /// </summary>
    public static class Summary
    {
        public static void ExtractCounty(Workbook wb, List<PreviewTable> into)
        {
            Sheet sh;
            if (!wb.Sheets.TryGetValue("通报", out sh)) return;

            var t = new PreviewTable
            {
                Title = "区县正向销售完成情况（" + TitleOf(sh) + "）",
                Columns = new[]
                {
                    "区县", "收入目标", "收入完成值", "收入完成率", "收入排名",
                    "新增目标", "新增完成值", "新增完成率",
                    "宽带目标", "宽带完成值", "宽带完成率"
                }
            };

            // 第一块：第 5~13 行；第二块（宽带）标题下隔一行再起，两块区县顺序一致。
            int label = FindRow(sh, "宽带（按100元/户初装费计算）");
            int offset = label > 0 ? (label + 2) - 5 : 0;

            for (int r = 5; r <= 13; r++)
            {
                string name = Clean(Text(sh, r, 1));
                if (name.Length == 0) continue;
                int r2 = offset > 0 ? r + offset : r;
                t.Rows.Add(new[]
                {
                    name,
                    Num(sh, r, 2), Num(sh, r, 3), Pct(sh, r, 4), Num(sh, r, 5),
                    Num(sh, r, 6), Num(sh, r, 7), Pct(sh, r, 8),
                    Num(sh, r2, 2), Num(sh, r2, 3), Pct(sh, r2, 4)
                });
            }
            into.Add(t);
        }

        public static void ExtractGrid(Workbook wb, List<PreviewTable> into)
        {
            Sheet sh;
            if (!wb.Sheets.TryGetValue("通报", out sh)) return;

            var t = new PreviewTable
            {
                Title = "网格正向销售完成情况（" + TitleOf(sh) + "）",
                Columns = new[]
                {
                    "区县", "网格", "收入目标", "收入完成值", "收入完成率", "收入排名",
                    "新增完成值", "新增完成率", "宽带完成值", "爆款收入"
                }
            };

            int last = sh.MaxRow;
            var rows = new List<string[]>();
            for (int r = 5; r <= last; r++)
            {
                string name = Text(sh, r, 2);
                if (name.Length == 0 || name.StartsWith("备注")) break;
                rows.Add(new[]
                {
                    Clean(Text(sh, r, 1)), name,
                    Num(sh, r, 3), Num(sh, r, 4), Pct(sh, r, 5), Num(sh, r, 6),
                    Num(sh, r, 7), Pct(sh, r, 10), Num(sh, r, 13), Num(sh, r, 36)
                });
            }
            t.Rows = rows;
            into.Add(t);
        }

        public static void ExtractKeyBusiness(Workbook wb, List<PreviewTable> into)
        {
            Sheet sh;
            if (!wb.Sheets.TryGetValue("通报", out sh)) return;

            // 表头：E..M 是 8 个区县 + 全州
            var cols = new List<string> { "指标" };
            for (int c = 5; c <= 13; c++)
            {
                string h = Text(sh, 2, c);
                if (h.Length == 0) break;
                cols.Add(h);
            }

            var t = new PreviewTable
            {
                Title = "重点指标调度通报（" + TitleOf(sh) + "）",
                Columns = cols.ToArray()
            };

            // 大类/小类是合并单元格，只在首行出现，需要向下继承。
            string carryB = "", carryC = "";
            for (int r = 3; r <= sh.MaxRow; r++)
            {
                string item = Text(sh, r, 4);
                if (item.Length == 0) continue;
                if (item.StartsWith("注")) break;
                if (item == "指标项") continue;

                string cellB = Text(sh, r, 2), cellC = Text(sh, r, 3);
                if (cellB.Length > 0) { carryB = cellB; carryC = cellC; }
                else if (cellC.Length > 0) carryC = cellC;

                string label = Join(carryB, carryC, item);

                var row = new List<string> { label };
                bool percent = item.Contains("完成率");
                for (int c = 5; c < 5 + cols.Count - 1; c++)
                {
                    if (c > sh.MaxCol) { row.Add(""); continue; }
                    row.Add(percent ? Pct(sh, r, c) : Num(sh, r, c));
                }
                t.Rows.Add(row.ToArray());
            }
            into.Add(t);
        }

        // ---------- 读值小工具 ----------

        static int FindRow(Sheet sh, string text)
        {
            for (int r = 1; r <= sh.MaxRow; r++)
                for (int c = 1; c <= Math.Min(sh.MaxCol, 8); c++)
                {
                    string v = Text(sh, r, c);
                    if (v.Length == 0) continue;
                    // 单元格里常带换行（如 "宽带（…）\n-不含未知网格"），按首行比。
                    int nl = v.IndexOfAny(new[] { '\r', '\n' });
                    if (nl > 0) v = v.Substring(0, nl).Trim();
                    if (v == text || v.StartsWith(text)) return r;
                }
            return -1;
        }

        static string TitleOf(Sheet sh)
        {
            string v = Text(sh, 1, 1);
            int sp = v.LastIndexOf(' ');
            return sp > 0 ? v.Substring(sp + 1) : v;
        }

        static string Join(params string[] parts)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var p in parts)
            {
                if (string.IsNullOrEmpty(p)) continue;
                if (sb.Length > 0) sb.Append(' ');
                sb.Append(p.Replace("\n", "").Replace("\r", "").Trim());
            }
            return sb.ToString();
        }

        static string Text(Sheet sh, int row, int col)
        {
            var v = sh.Get(row, col);
            if (v.Kind == ValueKind.Text) return (v.Text ?? "").Trim();
            return "";
        }

        // 表里的区县名带 "A_" 这类排序前缀，显示时去掉。
        static string Clean(string s)
        {
            if (s.Length > 2 && s[1] == '_') return s.Substring(2);
            return s;
        }

        static string Num(Sheet sh, int row, int col)
        {
            var v = sh.Get(row, col);
            switch (v.Kind)
            {
                case ValueKind.Number:
                    if (Math.Abs(v.Num) < 1e-9) return "0";
                    if (Math.Abs(v.Num - Math.Round(v.Num)) < 1e-9 && Math.Abs(v.Num) < 1e12)
                        return ((long)Math.Round(v.Num)).ToString(CultureInfo.InvariantCulture);
                    return v.Num.ToString("#,0.##", CultureInfo.InvariantCulture);
                case ValueKind.Text: return v.Text ?? "";
                default: return "";
            }
        }

        static string Pct(Sheet sh, int row, int col)
        {
            var v = sh.Get(row, col);
            if (v.Kind == ValueKind.Number)
                return (v.Num * 100).ToString("0.0", CultureInfo.InvariantCulture) + "%";
            if (v.Kind == ValueKind.Text) return v.Text ?? "";
            return "";
        }
    }
}
