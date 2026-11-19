using System;
using System.Collections.Generic;
using System.IO;
using ReportTool.Core;

namespace ReportTool.App
{
    public enum ReportKind { Unknown, County, Grid, KeyBusiness }

    public static class KindNames
    {
        public static string Of(ReportKind k)
        {
            switch (k)
            {
                case ReportKind.County: return "区县正向销售";
                case ReportKind.Grid: return "网格正向销售";
                case ReportKind.KeyBusiness: return "重点业务调度";
                default: return "未识别";
            }
        }

        public static string ShortFilePrefix(ReportKind k)
        {
            switch (k)
            {
                case ReportKind.County: return "区县";
                case ReportKind.Grid: return "网格";
                case ReportKind.KeyBusiness: return "重点业务";
                default: return "未识别";
            }
        }
    }

    public class FileResult
    {
        public string InputPath;
        public ReportKind Kind;
        public string OutputPath;
        public int Formulas;
        public int Passes;
        public int Errors;
        public long ElapsedMs;
        public string Error;
    }

    public class BatchResult
    {
        public string OutputDir;
        public List<FileResult> Files = new List<FileResult>();
        // 关键指标表，供界面展示
        public List<PreviewTable> Tables = new List<PreviewTable>();
        // 图表数据，界面按它画预览
        public List<ChartData> Charts = new List<ChartData>();
        // 生成的图片路径（每张图一个文件，外加一张拼好的"一页纸"）
        public List<string> ChartFiles = new List<string>();
        public string CombinedChartPath;

        public bool Ok
        {
            get
            {
                foreach (var f in Files) if (f.Error != null) return false;
                return Files.Count > 0;
            }
        }
    }

    public static class Pipeline
    {
        /// <summary>按工作表结构判断这是哪一份报表。</summary>
        public static ReportKind Detect(Workbook wb)
        {
            var names = new List<string>();
            foreach (var s in wb.AllSheets) names.Add(s.Name);
            return Detect(names);
        }

        /// <summary>按工作表名判断报表类型。与 Detect(Workbook) 同一套规则。</summary>
        public static ReportKind Detect(IList<string> sheetNames)
        {
            bool hasGrid = false, hasIncome = false, hasAuto = false, hasReport = false;
            foreach (var n in sheetNames)
            {
                if (string.Equals(n, "网格层级重点业务通报", StringComparison.OrdinalIgnoreCase)) hasGrid = true;
                else if (string.Equals(n, "收入", StringComparison.OrdinalIgnoreCase)) hasIncome = true;
                else if (string.Equals(n, "auto", StringComparison.OrdinalIgnoreCase)) hasAuto = true;
                else if (string.Equals(n, "通报", StringComparison.OrdinalIgnoreCase)) hasReport = true;
            }
            if (!hasAuto) return ReportKind.Unknown;
            if (hasGrid) return ReportKind.Grid;
            if (hasIncome) return ReportKind.KeyBusiness;
            if (hasReport) return ReportKind.County;
            return ReportKind.Unknown;
        }

        /// <summary>
        /// 快速识别：只读工作表名，不解析整表。
        /// 拖入文件夹时可能有上百个文件、还混着几十 MB 的大表，
        /// 整表读会把界面卡死，这里必须保持毫秒级。
        /// </summary>
        public static ReportKind Peek(string path)
        {
            return Detect(Inbox.QuickSheetNames(path));
        }

        /// <summary>
        /// 执行整批：重算每一份、写回 xlsx（保留原格式）、提取关键指标。
        /// outputRoot 为空时写到「程序所在目录\outputs\日期」。
        /// </summary>
        public static BatchResult Run(IList<string> inputs, string outputRoot, Action<string> log)
        {
            var batch = new BatchResult();
            DateTime now = DateTime.Now;
            string dir = outputRoot;
            if (string.IsNullOrEmpty(dir))
                dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "outputs");
            dir = Path.Combine(dir, now.ToString("yyyy-MM-dd"));
            Directory.CreateDirectory(dir);
            batch.OutputDir = dir;

            var seen = new Dictionary<ReportKind, string>();

            foreach (var path in inputs)
            {
                var fr = new FileResult { InputPath = path };
                batch.Files.Add(fr);
                var sw = System.Diagnostics.Stopwatch.StartNew();

                try
                {
                    if (!File.Exists(path)) throw new FileNotFoundException("文件不存在");

                    var wb = XlsxReader.Read(path);
                    fr.Kind = Detect(wb);
                    if (fr.Kind == ReportKind.Unknown)
                        throw new InvalidDataException("认不出这是哪一份报表（缺少 auto / 通报 等工作表）");
                    if (seen.ContainsKey(fr.Kind))
                        throw new InvalidDataException("这份报表已经拖进来过了：" + Path.GetFileName(seen[fr.Kind]));
                    seen[fr.Kind] = path;

                    if (log != null) log("开始重算：" + Path.GetFileName(path) + "（" + KindNames.Of(fr.Kind) + "）");

                    var report = new Calculator().Recalculate(wb);
                    fr.Formulas = report.Computed;
                    fr.Passes = report.Passes;
                    fr.Errors = report.Errors;

                    if (report.Errors > 0)
                    {
                        // 不中断，但把出错的单元格报出来
                        var msg = new System.Text.StringBuilder();
                        msg.Append("有 " + report.Errors + " 个公式算不出来：");
                        for (int i = 0; i < report.ErrorList.Count && i < 5; i++)
                        {
                            var e = report.ErrorList[i];
                            msg.Append(" " + e.Sheet + "!" + e.Cell + "(" + e.Message + ")");
                        }
                        if (log != null) log("  警告：" + msg);
                    }

                    string outName = KindNames.ShortFilePrefix(fr.Kind) + "_" + now.ToString("HHmm") + ".xlsx";
                    string outPath = Path.Combine(dir, outName);
                    int n = 2;
                    while (File.Exists(outPath))
                        outPath = Path.Combine(dir, KindNames.ShortFilePrefix(fr.Kind) + "_" + now.ToString("HHmm") + "_" + (n++) + ".xlsx");

                    XlsxWriter.Write(wb, outPath);
                    fr.OutputPath = outPath;
                    fr.ElapsedMs = sw.ElapsedMilliseconds;

                    if (log != null)
                        log("  完成：" + Path.GetFileName(outPath) + "（公式 " + fr.Formulas +
                            "，迭代 " + fr.Passes + " 轮，" + fr.ElapsedMs + " 毫秒）");

                    if (fr.Kind == ReportKind.County)
                    {
                        Summary.ExtractCounty(wb, batch.Tables);
                        Summary.BuildCountyCharts(wb, batch.Charts);
                    }
                    else if (fr.Kind == ReportKind.Grid)
                    {
                        Summary.ExtractGrid(wb, batch.Tables);
                        Summary.BuildGridCharts(wb, batch.Charts);
                    }
                    else if (fr.Kind == ReportKind.KeyBusiness)
                    {
                        Summary.ExtractKeyBusiness(wb, batch.Tables);
                        Summary.BuildKeyBusinessCharts(wb, batch.Charts);
                    }
                }
                catch (Exception ex)
                {
                    fr.Error = ex.Message;
                    fr.ElapsedMs = sw.ElapsedMilliseconds;
                    if (log != null) log("  失败：" + ex.Message);
                }
            }

            // 出图：每张单独一个 PNG，"一页纸"再拼一张
            if (batch.Charts.Count > 0)
            {
                try
                {
                    int i = 1;
                    foreach (var c in batch.Charts)
                    {
                        string name = "图" + i + "_" + SafeName(c.Title) + ".png";
                        string path = Path.Combine(dir, name);
                        ChartRenderer.Save(c, path);
                        batch.ChartFiles.Add(path);
                        if (log != null) log("  出图：" + name);
                        i++;
                    }
                    if (batch.Charts.Count > 1)
                    {
                        string all = Path.Combine(dir, "日报_一页纸.png");
                        ChartRenderer.SaveStack(batch.Charts, all, 900);
                        batch.CombinedChartPath = all;
                        if (log != null) log("  出图：日报_一页纸.png（" + batch.Charts.Count + " 张拼在一起）");
                    }
                }
                catch (Exception ex)
                {
                    // 出图失败不该影响已经算好的 Excel
                    if (log != null) log("  出图失败：" + ex.Message);
                }
            }

            // 缺哪一份就提醒
            foreach (ReportKind k in new[] { ReportKind.County, ReportKind.Grid, ReportKind.KeyBusiness })
            {
                if (!seen.ContainsKey(k) && log != null)
                    log("提示：没有拖入「" + KindNames.Of(k) + "」，只处理了已有的 " + seen.Count + " 份。");
            }

            return batch;
        }

        /// <summary>文件名里不能出现的字符换成下划线。</summary>
        static string SafeName(string s)
        {
            if (string.IsNullOrEmpty(s)) return "未命名";
            var bad = new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };
            var sb = new System.Text.StringBuilder();
            foreach (char ch in s)
            {
                bool ok = true;
                foreach (char b in bad) if (ch == b) { ok = false; break; }
                sb.Append(ok ? ch : '_');
            }
            string r = sb.ToString().Trim();
            return r.Length > 30 ? r.Substring(0, 30) : r;
        }
    }
}
