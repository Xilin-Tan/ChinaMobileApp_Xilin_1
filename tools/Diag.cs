using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using ReportTool.Core;

namespace ReportTool.Tools
{
    /// <summary>
    /// 诊断工具：读取 xlsx、重算所有公式、输出结果到文本，便于与原版对照。
    /// 用法: diag.exe &lt;xlsx路径&gt; [输出txt]
    /// </summary>
    public static class Diag
    {
        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            if (args.Length < 1)
            {
                Console.WriteLine("用法: diag.exe <xlsx> [out.txt]");
                return 2;
            }

            string path = args[0];
            var sb = new StringBuilder();
            var wb = XlsxReader.Read(path);
            sb.AppendLine("FILE: " + Path.GetFileName(path));
            sb.AppendLine("SHEETS: " + string.Join(", ", new List<string>(wb.Sheets.Keys).ToArray()));

            foreach (var kv in wb.Sheets)
            {
                var sh = kv.Value;
                sb.AppendLine();
                sb.AppendLine("== SHEET " + sh.Name + "  vals=" + sh.Values.Count + " formulas=" + sh.Formulas.Count + " max=" + sh.MaxRow + "x" + sh.MaxCol);
            }

            var calc = new Calculator();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var report = calc.Recalculate(wb);
            sw.Stop();

            sb.AppendLine();
            sb.AppendLine("=== 重算耗时 " + sw.ElapsedMilliseconds + " ms");
            sb.AppendLine("公式格数: " + report.Computed + " | 迭代轮次: " + report.Passes + " | 错误: " + report.Errors);
            foreach (var e in report.ErrorList) sb.AppendLine("  ERR " + e.Sheet + "!" + e.Cell + " = " + e.Message);

            // dump 通报表数值区
            foreach (var name in new[] { "通报" })
            {
                Sheet sh;
                if (!wb.Sheets.TryGetValue(name, out sh)) continue;
                sb.AppendLine();
                sb.AppendLine("### " + name + " 数值");
                for (int r = 1; r <= sh.MaxRow; r++)
                {
                    var parts = new List<string>();
                    for (int c = 1; c <= sh.MaxCol; c++)
                    {
                        var v = sh.Get(r, c);
                        if (v.Kind == ValueKind.Blank) continue;
                        parts.Add(ColName(c) + r + "=" + Fmt(v));
                    }
                    if (parts.Count > 0) sb.AppendLine("R" + r + " | " + string.Join(" | ", parts.ToArray()));
                }
            }

            string text = sb.ToString();
            if (args.Length > 1) { File.WriteAllText(args[1], text, new UTF8Encoding(false)); Console.WriteLine("written: " + args[1]); }
            else Console.Write(text);
            return 0;
        }

        static string ColName(int col)
        {
            var sb = new StringBuilder();
            while (col > 0)
            {
                int m = (col - 1) % 26;
                sb.Insert(0, (char)('A' + m));
                col = (col - 1) / 26;
            }
            return sb.ToString();
        }

        static string Fmt(CellValue v)
        {
            switch (v.Kind)
            {
                case ValueKind.Number: return Math.Abs(v.Num) < 1e-9 ? "0" : v.Num.ToString("0.####", CultureInfo.InvariantCulture);
                case ValueKind.Text: return v.Text;
                case ValueKind.Bool: return v.Bool ? "TRUE" : "FALSE";
                case ValueKind.Error: return "[" + v.Text + "]";
                default: return "";
            }
        }
    }
}