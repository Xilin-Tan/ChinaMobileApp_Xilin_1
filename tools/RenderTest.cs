using System;
using System.IO;
using ReportTool.Core;

namespace ReportTool.Tools
{
    /// <summary>
    /// 端到端测试：读 xlsx -> 重算 -> 写回（保留格式）-> 再读一次校验数值。
    /// 用法: rendertest.exe <xlsx> <out.xlsx>
    /// </summary>
    public static class RenderTest
    {
        public static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            if (args.Length < 2) { Console.WriteLine("用法: rendertest.exe <in.xlsx> <out.xlsx>"); return 2; }

            string inp = args[0], outp = args[1];
            var wb = XlsxReader.Read(inp);
            var report = new Calculator().Recalculate(wb);
            Console.WriteLine("重算: 公式=" + report.Computed + " 轮次=" + report.Passes + " 错误=" + report.Errors);

            XlsxWriter.Write(wb, outp);
            Console.WriteLine("已写出: " + outp + " (" + new FileInfo(outp).Length + " 字节)");

            // 复核：重新读回输出文件，确认公式格已有值
            var wb2 = XlsxReader.Read(outp);
            int checkedCells = 0, mismatched = 0;
            foreach (var kv in wb.Sheets)
            {
                Sheet outSheet;
                if (!wb2.Sheets.TryGetValue(kv.Key, out outSheet)) { Console.WriteLine("缺表: " + kv.Key); continue; }
                var sh = kv.Value;
                foreach (var f in sh.Formulas)
                {
                    long key = f.Key;
                    int r = (int)(key >> 20), c = (int)(key & 0xFFFFF);
                    var want = sh.Get(r, c);
                    var got = outSheet.Get(r, c);
                    checkedCells++;
                    bool same;
                    if (want.Kind == ValueKind.Number && got.Kind == ValueKind.Number)
                        same = Math.Abs(want.Num - got.Num) <= 1e-9 * Math.Max(1.0, Math.Abs(want.Num));
                    else if (want.Kind == ValueKind.Text && got.Kind == ValueKind.Text)
                        same = string.Equals(want.Text, got.Text, StringComparison.Ordinal);
                    else same = want.Kind == got.Kind && want.Kind != ValueKind.Blank;

                    if (!same)
                    {
                        mismatched++;
                        if (mismatched <= 10)
                            Console.WriteLine("  不符 " + kv.Key + "!" + Calculator.ColName(c) + r +
                                              " 期望=" + want + " 实际=" + got);
                    }
                }
            }
            Console.WriteLine("复核: 检查 " + checkedCells + " 个公式格，不符 " + mismatched);
            return mismatched == 0 ? 0 : 1;
        }
    }
}
