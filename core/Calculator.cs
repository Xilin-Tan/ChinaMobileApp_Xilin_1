using System;
using System.Collections.Generic;

namespace ReportTool.Core
{
    public class CalcReport
    {
        public int Computed;
        public int Passes;
        public int Errors;
        public List<CellError> ErrorList = new List<CellError>();
    }

    public class FormulaCell
    {
        public Sheet Sheet;
        public int Row;
        public int Col;
        public string Formula;
    }

    internal class CompiledCell
    {
        public Sheet Sheet;
        public int Row;
        public int Col;
        public Evaluator.Compiled Formula;
    }

    public class CellError
    {
        public string Sheet;
        public string Cell;
        public string Message;
    }

    /// <summary>
    /// 不动点迭代重算：反复计算全部公式，直到所有单元格数值不再变化。
    /// 这样不依赖公式的书写顺序，能正确处理「前面引用后面」的情况
    /// （原模板 C5 引用 X17/AD17 就是这种）。
    /// </summary>
    public class Calculator
    {
        readonly Evaluator _eval = new Evaluator();

        public CalcReport Recalculate(Workbook wb)
        {
            return Recalculate(wb, null);
        }

        /// <summary>progress: (当前轮次, 总轮次上限) —— 用于界面显示进度。</summary>
        public CalcReport Recalculate(Workbook wb, Action<int, int> progress)
        {
            var report = new CalcReport();

            // 收集所有公式
            var all = new List<FormulaCell>();
            foreach (var kv in wb.Sheets)
            {
                var sheet = kv.Value;
                foreach (var f in sheet.Formulas)
                {
                    long key = f.Key;
                    all.Add(new FormulaCell
                    {
                        Sheet = sheet,
                        Row = (int)(key >> 20),
                        Col = (int)(key & 0xFFFFF),
                        Formula = f.Value
                    });
                }
            }

            // 公式只解析一次
            var compiled = new CompiledCell[all.Count];
            for (int i = 0; i < all.Count; i++)
            {
                var fc = all[i];
                compiled[i] = new CompiledCell
                {
                    Sheet = fc.Sheet,
                    Row = fc.Row,
                    Col = fc.Col,
                    Formula = _eval.Compile(fc.Formula)
                };
            }

            const int MaxPasses = 40;
            for (int pass = 0; pass < MaxPasses; pass++)
            {
                if (progress != null) progress(pass + 1, MaxPasses);
                int changed = 0;
                foreach (var fc in compiled)
                {
                    var ctx = new EvalContext
                    {
                        Workbook = wb,
                        Sheet = fc.Sheet,
                        Row = fc.Row,
                        Col = fc.Col,
                        Depth = 0
                    };

                    CellValue result;
                    try
                    {
                        result = _eval.Evaluate(fc.Formula, ctx);
                    }
                    catch (Exception ex)
                    {
                        result = CellValue.Error("#EX:" + ex.GetType().Name);
                    }

                    if (result.Kind == ValueKind.Error && (result.Text == "#CIRC!" || result.Text.StartsWith("#FUNC") || result.Text.StartsWith("#EX")))
                        continue;

                    var prev = fc.Sheet.Get(fc.Row, fc.Col);
                    if (SameValue(prev, result)) continue;

                    fc.Sheet.Set(fc.Row, fc.Col, result);
                    changed++;
                }

                report.Passes = pass + 1;
                if (changed == 0) break;
            }

            // 统计
            report.Computed = compiled.Length;
            foreach (var fc in compiled)
            {
                var v = fc.Sheet.Get(fc.Row, fc.Col);
                if (v.Kind == ValueKind.Error)
                {
                    report.Errors++;
                    report.ErrorList.Add(new CellError
                    {
                        Sheet = fc.Sheet.Name,
                        Cell = ColName(fc.Col) + fc.Row,
                        Message = v.Text
                    });
                }
            }

            return report;
        }

        static bool SameValue(CellValue a, CellValue b)
        {
            if (a.Kind != b.Kind) return false;
            switch (a.Kind)
            {
                case ValueKind.Number:
                    if (a.Num == b.Num) return true;
                    double scale = Math.Max(Math.Abs(a.Num), Math.Abs(b.Num));
                    return Math.Abs(a.Num - b.Num) <= 1e-12 * Math.Max(1.0, scale);
                case ValueKind.Text: return string.Equals(a.Text, b.Text, StringComparison.Ordinal);
                case ValueKind.Bool: return a.Bool == b.Bool;
                case ValueKind.Error: return string.Equals(a.Text, b.Text, StringComparison.Ordinal);
                default: return true;
            }
        }

        public static string ColName(int col)
        {
            var sb = new System.Text.StringBuilder();
            while (col > 0)
            {
                int m = (col - 1) % 26;
                sb.Insert(0, (char)('A' + m));
                col = (col - 1) / 26;
            }
            return sb.ToString();
        }
    }
}
