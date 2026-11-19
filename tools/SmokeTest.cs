using System;
using System.Collections.Generic;
using System.IO;
using ReportTool.App;

namespace ReportTool.Tools
{
    /// <summary>不开界面，直接跑一遍完整流程并打印结果，用于回归测试。</summary>
    public static class SmokeTest
    {
        public static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            if (args.Length < 1) { Console.WriteLine("用法: smoke.exe <xlsx|xlsx...>"); return 2; }

            var inputs = new List<string>(args);
            var r = Pipeline.Run(inputs, Path.Combine(Path.GetTempPath(), "reporttool_smoke"),
                                 delegate(string s) { Console.WriteLine(s); });

            Console.WriteLine();
            int fail = 0;
            foreach (var f in r.Files)
            {
                Console.WriteLine((f.Error == null ? "[OK]   " : "[FAIL] ") + Path.GetFileName(f.InputPath)
                                  + "  -> " + (f.OutputPath ?? f.Error));
                if (f.Error != null) fail++;
            }
            foreach (var t in r.Tables)
            {
                Console.WriteLine();
                Console.WriteLine("### " + t.Title);
                Console.WriteLine(string.Join(" | ", t.Columns));
                foreach (var row in t.Rows) Console.WriteLine(string.Join(" | ", row));
            }
            return fail == 0 ? 0 : 1;
        }
    }
}
