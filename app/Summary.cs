using System;
using System.Collections.Generic;
using System.Drawing;
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

        // ---------- 图表数据（数值一律取自 CellValue.Num，不从界面字符串反解） ----------

        /// <summary>区县：收入完成率条形图 + 新增/宽带完成率对比柱状图。</summary>
        public static void BuildCountyCharts(Workbook wb, List<ChartData> into)
        {
            Sheet sh;
            if (!wb.Sheets.TryGetValue("通报", out sh)) return;

            int label = FindRow(sh, "宽带（按100元/户初装费计算）");
            int offset = label > 0 ? (label + 2) - 5 : 0;

            var names = new List<string>();
            var income = new List<double>();
            var addRate = new List<double>();
            var bbRate = new List<double>();

            for (int r = 5; r <= 13; r++)
            {
                string name = Clean(Text(sh, r, 1));
                if (name.Length == 0) continue;
                int r2 = offset > 0 ? r + offset : r;
                names.Add(name);
                income.Add(Value(sh, r, 4));
                addRate.Add(Value(sh, r, 8));
                bbRate.Add(Value(sh, r2, 4));
            }
            if (names.Count == 0) return;

            // 只画 8 个区县，末尾的"全州"合计另看
            int n = names.Count;
            if (n > 8 && names[n - 1] == "全州") n = 8;

            var c1 = new ChartData
            {
                Title = "区县收入完成率",
                Subtitle = "按 100% 达标线排序 · " + TitleOf(sh),
                Categories = names.GetRange(0, n).ToArray(),
                Horizontal = true,
                Percent = true,
                Reference = 1.0,
                ShowReference = true
            };
            var s1 = new ChartSeries { Name = "收入完成率", Values = new double[n] };
            for (int i = 0; i < n; i++) s1.Values[i] = income[i];
            SortDescending(c1.Categories, s1.Values);
            c1.Series.Add(s1);
            into.Add(c1);

            var c2 = new ChartData
            {
                Title = "区县新增 / 宽带完成率",
                Subtitle = "100% = 达标 · " + TitleOf(sh),
                Categories = names.GetRange(0, n).ToArray(),
                Percent = true,
                Reference = 1.0,
                ShowReference = true
            };
            var a = new ChartSeries { Name = "新增完成率", Color = Color.FromArgb(52, 116, 214), Values = addRate.GetRange(0, n).ToArray() };
            var b = new ChartSeries { Name = "宽带完成率", Color = Color.FromArgb(240, 154, 60), Values = bbRate.GetRange(0, n).ToArray() };
            c2.Series.Add(a);
            c2.Series.Add(b);
            into.Add(c2);
        }

        /// <summary>网格：收入完成率条形图（全部网格，从高到低）。</summary>
        public static void BuildGridCharts(Workbook wb, List<ChartData> into)
        {
            Sheet sh;
            if (!wb.Sheets.TryGetValue("通报", out sh)) return;

            var names = new List<string>();
            var rates = new List<double>();
            var amount = new List<double>();

            for (int r = 5; r <= sh.MaxRow; r++)
            {
                string name = Text(sh, r, 2);
                if (name.Length == 0 || name.StartsWith("备注")) break;
                names.Add(Clean(name));
                rates.Add(Value(sh, r, 5));
                amount.Add(Value(sh, r, 4));
            }
            if (names.Count == 0) return;

            var c = new ChartData
            {
                Title = "网格收入完成率",
                Subtitle = "全部 " + names.Count + " 个网格 · 按 100% 达标线排序 · " + TitleOf(sh),
                Categories = names.ToArray(),
                Horizontal = true,
                Percent = true,
                Reference = 1.0,
                ShowReference = true
            };
            var s = new ChartSeries { Name = "收入完成率", Values = rates.ToArray() };
            SortDescending(c.Categories, s.Values);
            c.Series.Add(s);
            into.Add(c);
        }

        /// <summary>重点业务：几个关键指标的完成率，按区县分组对比。</summary>
        public static void BuildKeyBusinessCharts(Workbook wb, List<ChartData> into)
        {
            Sheet sh;
            if (!wb.Sheets.TryGetValue("通报", out sh)) return;

            var areas = new List<string>();
            for (int col = 5; col <= 13; col++)
            {
                string h = Text(sh, 2, col);
                if (h.Length == 0) break;
                areas.Add(h);
            }
            if (areas.Count == 0) return;

            var colors = new[]
            {
                Color.FromArgb(52, 116, 214), Color.FromArgb(240, 154, 60),
                Color.FromArgb(46, 155, 92), Color.FromArgb(158, 108, 200)
            };

            var series = new List<ChartSeries>();
            string carryB = "", carryC = "";

            for (int r = 3; r <= sh.MaxRow && series.Count < 4; r++)
            {
                string item = Text(sh, r, 4);
                if (item.Length == 0) continue;
                if (item.StartsWith("注")) break;

                string cellB = Text(sh, r, 2), cellC = Text(sh, r, 3);
                if (cellB.Length > 0) { carryB = cellB; carryC = cellC; }
                else if (cellC.Length > 0) carryC = cellC;

                if (!item.Contains("完成率")) continue;

                string label = Join(carryB, carryC, item);
                // 图例要短：去掉"完成率"和"（40分）"这类评分后缀
                label = label.Replace("完成率", "");
                int br = label.LastIndexOf('（');
                if (br > 0 && label.Length - br <= 10) label = label.Substring(0, br).Trim();
                if (label.Length == 0) continue;

                var vals = new double[areas.Count];
                for (int i = 0; i < areas.Count; i++)
                    vals[i] = Value(sh, r, 5 + i);

                series.Add(new ChartSeries
                {
                    Name = label,
                    Color = colors[series.Count % colors.Length],
                    Values = vals
                });
            }
            if (series.Count == 0) return;

            var c = new ChartData
            {
                Title = "重点指标完成率（分区县）",
                Subtitle = "100% = 达标 · " + TitleOf(sh),
                Categories = areas.ToArray(),
                Percent = true,
                Reference = 1.0,
                ShowReference = true
            };
            c.Series.AddRange(series);
            into.Add(c);
        }

        static void SortDescending(string[] names, double[] values)
        {
            int n = names.Length;
            var idx = new int[n];
            for (int i = 0; i < n; i++) idx[i] = i;
            Array.Sort(idx, delegate(int x, int y)
            {
                int c = values[y].CompareTo(values[x]);
                return c != 0 ? c : string.CompareOrdinal(names[x], names[y]);
            });
            var nn = new string[n];
            var vv = new double[n];
            for (int i = 0; i < n; i++) { nn[i] = names[idx[i]]; vv[i] = values[idx[i]]; }
            Array.Copy(nn, names, n);
            Array.Copy(vv, values, n);
        }

        /// <summary>取数值单元格；空或文本一律当 0，图表不该因为空格子崩掉。</summary>
        static double Value(Sheet sh, int row, int col)
        {
            var v = sh.Get(row, col);
            if (v.Kind == ValueKind.Number) return v.Num;
            if (v.Kind == ValueKind.Bool) return v.Bool ? 1 : 0;
            return 0;
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

        /// <summary>取标题里的时间部分，如 "2026年11月18日 16:45"。</summary>
        static string TitleOf(Sheet sh)
        {
            string v = Text(sh, 1, 1);
            int y = v.IndexOf("20");
            if (y > 0 && v.IndexOf('年') > y) return v.Substring(y).Trim();
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
