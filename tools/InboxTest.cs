using System;
using System.Collections.Generic;
using System.IO;
using ReportTool.App;

namespace ReportTool.Tools
{
    /// <summary>
    /// 验证企业微信目录扫描：找目录 → 列出报表 → 读每份的通报时间 → 取每类最新 → 可选实跑。
    /// 用法：inboxtest.exe [--run]
    /// 重点看：每类选出来的那份，通报时间是不是真的最新（文件名序号和修改时间都不作数）。
    /// </summary>
    public static class InboxTest
    {
        public static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            bool run = args.Length > 0 && args[0] == "--run";

            var folders = Inbox.FindWeChatFolders();
            if (folders.Count == 0) { Console.WriteLine("企业微信目录: <未找到>"); return 1; }
            foreach (var fd in folders) Console.WriteLine("企业微信目录: " + fd);

            var all = Inbox.ListAll(folders);
            Console.WriteLine("识别出的报表总数: " + all.Count);
            if (all.Count == 0) return 1;

            Console.WriteLine();
            Console.WriteLine("读每份的通报时间：");
            Inbox.ProbeTimes(all, delegate(string s) { Console.WriteLine(s); });

            int noTime = 0;
            foreach (var f in all) if (!f.ReportTime.HasValue) noTime++;
            Console.WriteLine("读不到时间的: " + noTime + " 份" +
                              (noTime > 0 ? "（会退回文件修改时间）" : ""));

            Console.WriteLine();
            var byKind = new Dictionary<ReportKind, int>();
            foreach (var f in all)
            {
                int n;
                byKind.TryGetValue(f.Kind, out n);
                byKind[f.Kind] = n + 1;
            }
            foreach (var kv in byKind)
                Console.WriteLine("  " + KindNames.Of(kv.Key) + ": " + kv.Value + " 份");

            Console.WriteLine();
            Console.WriteLine("每类最新一份:");
            var latest = Inbox.LatestPerKind(all);
            foreach (var f in latest)
            {
                bool ready = Inbox.WaitUntilReady(f.Path, 2000);
                string when = f.ReportTime.HasValue
                    ? f.ReportTime.Value.ToString("yyyy-MM-dd HH:mm")
                    : "<读不到>";
                Console.WriteLine("  " + KindNames.Of(f.Kind).PadRight(12) + " 通报 " + when +
                                  " | " + (f.Size / 1024) + " KB | 可读=" + ready);
                Console.WriteLine("      " + Path.GetFileName(f.Path));
            }

            // 反例自检：必须比"按文件名序号"选得准
            Console.WriteLine();
            Console.WriteLine("命名识别自检:");
            foreach (var n in new[]
            {
                "区县正向销售完成情况日预测通报(2).xlsx",
                "网格正向销售完成情况日预测通报(12).xlsx",
                "重点业务日调度通报(6).xlsx",
                "全州入驻抖音本地生活渠道情况通报(2026.11.18).xlsx",
                "守机宝 截止（1118).xlsx"
            })
                Console.WriteLine("  " + KindNames.Of(Inbox.KindOfName(n)).PadRight(12) + " <- " + n);

            // 时间解析自检
            Console.WriteLine();
            Console.WriteLine("时间解析自检:");
            foreach (var s in new[]
            {
                "2026年11月19日 17:15",
                "2026年11月05日 12:15",
                "2026-11-19 17:15",
                "正向销售收入完成情况日预测通报 2026年11月18日 16:45",
                "978"
            })
            {
                var t = Inbox.ParseReportTime(s);
                Console.WriteLine("  " + s.PadRight(46) + " -> " +
                                  (t.HasValue ? t.Value.ToString("yyyy-MM-dd HH:mm") : "<认不出>"));
            }

            if (!run || latest.Count == 0) return 0;

            Console.WriteLine();
            Console.WriteLine("=== 用扫描结果实际跑一遍 ===");
            var inputs = new List<string>();
            foreach (var f in latest) inputs.Add(f.Path);
            var result = Pipeline.Run(inputs, Path.Combine(Path.GetTempPath(), "reporttool_inbox"),
                                      delegate(string s) { Console.WriteLine(s); });

            int fail = 0;
            foreach (var f in result.Files) if (f.Error != null) fail++;
            Console.WriteLine();
            Console.WriteLine(fail == 0 ? "全部成功" : ("失败 " + fail + " 份"));
            return fail == 0 ? 0 : 1;
        }
    }
}