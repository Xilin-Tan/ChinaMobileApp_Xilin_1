using System;
using System.Collections.Generic;

namespace ReportTool.Core
{
    public enum ValueKind { Blank, Number, Text, Bool, Error }

    public struct CellValue
    {
        public ValueKind Kind;
        public double Num;
        public string Text;
        public bool Bool;

        public static readonly CellValue Blank = new CellValue { Kind = ValueKind.Blank };
        public static CellValue FromNumber(double d) { return new CellValue { Kind = ValueKind.Number, Num = d }; }
        public static CellValue FromText(string s) { return new CellValue { Kind = ValueKind.Text, Text = s ?? "" }; }
        public static CellValue FromBool(bool b) { return new CellValue { Kind = ValueKind.Bool, Bool = b }; }
        public static CellValue Error(string code) { return new CellValue { Kind = ValueKind.Error, Text = code }; }

        public bool IsNumeric { get { return Kind == ValueKind.Number || Kind == ValueKind.Bool; } }

        public double ToNumber()
        {
            switch (Kind)
            {
                case ValueKind.Number: return Num;
                case ValueKind.Bool: return Bool ? 1 : 0;
                case ValueKind.Blank: return 0;
                case ValueKind.Text:
                    double d;
                    if (double.TryParse(Text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out d)) return d;
                    return 0;
                default: return 0;
            }
        }

        // 用于 SUMIFS 条件比较：统一成可比字符串
        public string ToCompareText()
        {
            switch (Kind)
            {
                case ValueKind.Number: return FormatNumber(Num);
                case ValueKind.Bool: return Bool ? "TRUE" : "FALSE";
                case ValueKind.Text: return Text ?? "";
                default: return "";
            }
        }

        public static string FormatNumber(double d)
        {
            if (d == Math.Floor(d) && Math.Abs(d) < 1e15) return ((long)d).ToString(System.Globalization.CultureInfo.InvariantCulture);
            return d.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        }

        public override string ToString() { return ToCompareText(); }
    }

    public class Sheet
    {
        public string Name;
        public Dictionary<long, CellValue> Values = new Dictionary<long, CellValue>();
        public Dictionary<long, string> Formulas = new Dictionary<long, string>();
        public int MaxRow = 1;
        public int MaxCol = 1;

        // Column cache. Only enabled for sheets without formulas (the raw detail sheets):
        // such sheets never change during recalculation, so each column is built once.
        public Dictionary<int, CellValue[]> FastCols;

        public static long Key(int row, int col) { return ((long)row << 20) | (uint)col; }

        public CellValue Get(int row, int col)
        {
            CellValue v;
            if (Values.TryGetValue(Key(row, col), out v)) return v;
            return CellValue.Blank;
        }

        public void Set(int row, int col, CellValue v)
        {
            Values[Key(row, col)] = v;
            if (row > MaxRow) MaxRow = row;
            if (col > MaxCol) MaxCol = col;
            if (FastCols != null) FastCols.Clear();
        }

        // Returns cached column values 1..MaxRow, or null when caching is off.
        public CellValue[] FastColumn(int col)
        {
            if (FastCols == null) return null;
            CellValue[] arr;
            if (FastCols.TryGetValue(col, out arr)) return arr;
            arr = new CellValue[MaxRow + 1];
            for (int r = 1; r <= MaxRow; r++)
            {
                CellValue v;
                if (Values.TryGetValue(Key(r, col), out v)) arr[r] = v;
            }
            FastCols[col] = arr;
            return arr;
        }
    }

    public class Workbook
    {
        public string Path;
        public Dictionary<string, Sheet> Sheets = new Dictionary<string, Sheet>(StringComparer.OrdinalIgnoreCase);
        public string FirstSheetName;

        // 表名 -> zip 内的 xml 路径（输出时按此替换）
        public Dictionary<string, string> SheetEntryPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public Sheet GetSheet(string name)
        {
            Sheet s;
            if (name != null && Sheets.TryGetValue(name, out s)) return s;
            if (Sheets.Count > 0)
            {
                foreach (var kv in Sheets) return kv.Value;
            }
            return null;
        }

        public IEnumerable<Sheet> AllSheets { get { return Sheets.Values; } }
    }
}
