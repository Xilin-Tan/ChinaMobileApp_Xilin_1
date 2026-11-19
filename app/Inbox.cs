using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace ReportTool.App
{
    public class FoundFile
    {
        public string Path;
        public ReportKind Kind;
        public DateTime Modified;
        public long Size;
        /// <summary>文件内容里的「通报时间」。读不到为 null。</summary>
        public DateTime? ReportTime;
    }

    /// <summary>
    /// 从企业微信的接收目录里自动找出最新的三份报表。
    /// 企业微信把收到的文件存在：
    ///   %APPDATA%\Tencent\xwechat\xwechat_files\&lt;wxid&gt;\msg\file\&lt;yyyy-MM&gt;\
    ///
    /// 两个坑：
    ///   1. 文件名里的 " (1)" " (2)" 是重名计数，和先后顺序无关；
    ///   2. 企业微信会把同一天的多个文件批量重写一遍，所有文件的修改时间会变成同一个。
    /// 所以不能按文件名、也不能按修改时间判断新旧，只能读文件内容里的「通报时间」。
    /// </summary>
    public static class Inbox
    {
        public const string CountyPrefix = "区县正向销售完成情况日预测通报";
        public const string GridPrefix = "网格正向销售完成情况日预测通报";
        public const string KeyPrefix = "重点业务日调度通报";

        /// <summary>按文件名前缀判断这是哪一份报表；认不出返回 Unknown。</summary>
        public static ReportKind KindOfName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return ReportKind.Unknown;
            if (fileName.StartsWith(CountyPrefix)) return ReportKind.County;
            if (fileName.StartsWith(GridPrefix)) return ReportKind.Grid;
            if (fileName.StartsWith(KeyPrefix)) return ReportKind.KeyBusiness;
            return ReportKind.Unknown;
        }

        /// <summary>
        /// 找企业微信的接收目录：当月 + 上月。
        /// 上月是给跨月那几小时用的 —— 月初刚打开软件、当月目录里还没收到文件时，
        /// 最新那份其实还在上个月的目录里。全都返回，调用方两个都扫。
        /// </summary>
        public static List<string> FindWeChatFolders()
        {
            var candidates = new List<string>();
            try
            {
                string root = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    @"Tencent\xwechat\xwechat_files");
                if (!Directory.Exists(root)) return candidates;

                foreach (var acct in Directory.GetDirectories(root))
                {
                    string fileRoot = Path.Combine(acct, "msg", "file");
                    if (!Directory.Exists(fileRoot)) continue;

                    DateTime now = DateTime.Now;
                    foreach (var stamp in new[] { now, now.AddMonths(-1) })
                    {
                        string month = Path.Combine(fileRoot, stamp.ToString("yyyy-MM"));
                        if (Directory.Exists(month) && !candidates.Contains(month)) candidates.Add(month);
                    }
                    if (candidates.Count > 0) break;
                }
            }
            catch (Exception)
            {
                return candidates;
            }
            return candidates;
        }

        /// <summary>只要一个目录时用这个（当月优先）。找不到返回 null。</summary>
        public static string FindWeChatFolder()
        {
            var all = FindWeChatFolders();
            return all.Count > 0 ? all[0] : null;
        }

        /// <summary>列出目录下所有能认出来的报表（不过滤新旧）。</summary>
        public static List<FoundFile> ListAll(IEnumerable<string> folders)
        {
            var list = new List<FoundFile>();
            foreach (var folder in folders)
            {
                if (!Directory.Exists(folder)) continue;
                string[] files;
                try { files = Directory.GetFiles(folder, "*.xlsx"); }
                catch (Exception) { continue; }

                foreach (var path in files)
                {
                    string name = Path.GetFileName(path);
                    if (name.StartsWith("~$")) continue;
                    var kind = KindOfName(name);
                    if (kind == ReportKind.Unknown) continue;

                    FileInfo fi;
                    try { fi = new FileInfo(path); } catch (Exception) { continue; }
                    list.Add(new FoundFile
                    {
                        Path = path,
                        Kind = kind,
                        Modified = fi.LastWriteTime,
                        Size = fi.Length
                    });
                }
            }
            return list;
        }

        /// <summary>给每份文件补上通报时间，边读边报告进度。读不到留 null。</summary>
        public static void ProbeTimes(List<FoundFile> files, Action<string> log)
        {
            foreach (var f in files)
            {
                f.ReportTime = ProbeReportTime(f.Path);
                if (log != null)
                {
                    string shown = f.ReportTime.HasValue
                        ? f.ReportTime.Value.ToString("MM-dd HH:mm")
                        : "时间读不到";
                    log("  " + Path.GetFileName(f.Path) + "  →  " + shown);
                }
            }
        }

        /// <summary>
        /// 每类只留最新的一份。
        /// 优先比内容里的通报时间；读不到才退回文件修改时间。
        /// </summary>
        public static List<FoundFile> LatestPerKind(List<FoundFile> all)
        {
            var best = new Dictionary<ReportKind, FoundFile>();
            foreach (var f in all)
            {
                FoundFile cur;
                if (!best.TryGetValue(f.Kind, out cur) || Newer(f, cur))
                    best[f.Kind] = f;
            }
            var result = new List<FoundFile>();
            foreach (ReportKind k in new[] { ReportKind.County, ReportKind.Grid, ReportKind.KeyBusiness })
            {
                FoundFile f;
                if (best.TryGetValue(k, out f)) result.Add(f);
            }
            return result;
        }

        static bool Newer(FoundFile a, FoundFile b)
        {
            if (a.ReportTime.HasValue && b.ReportTime.HasValue)
            {
                if (a.ReportTime.Value != b.ReportTime.Value)
                    return a.ReportTime.Value > b.ReportTime.Value;
                return a.Modified > b.Modified;   // 通报时间相同，取后落盘的那份
            }
            if (a.ReportTime.HasValue) return true;
            if (b.ReportTime.HasValue) return false;
            return a.Modified > b.Modified;
        }

        /// <summary>
        /// 只读工作表名（毫秒级），用来判断这是哪一份报表。
        /// 不做整表解析 —— 目录里混着 14 MB 的大表，全量读会卡很久。
        /// 读不出来返回空列表。
        /// </summary>
        public static List<string> QuickSheetNames(string path)
        {
            var names = new List<string>();
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    string wbXml = ReadEntryText(zip, "xl/workbook.xml");
                    if (wbXml == null) return names;
                    var doc = new System.Xml.XmlDocument();
                    doc.XmlResolver = null;
                    doc.LoadXml(wbXml);
                    foreach (System.Xml.XmlNode n in doc.GetElementsByTagName("sheet"))
                    {
                        var a = n.Attributes["name"];
                        if (a != null) names.Add(a.Value);
                    }
                }
            }
            catch (Exception)
            {
                names.Clear();
            }
            return names;
        }
        // ================= 读「通报时间」用的轻量探针 =================
        // 只解开 zip 取两张工作表 xml、用正则定位，不建整表模型、不算公式。
        // 单份约 20~60 毫秒，扫一遍目录的十几份文件不到 1 秒。

        /// <summary>
        /// 读文件里的「通报时间」，如 2026年11月19日 17:15。读不到返回 null。
        ///
        /// 依据：「通报」表 A1 是公式  "标题"&amp;" "&amp;auto!XX4 ，
        ///      它指向 auto 表里的时间格。列位置会随模板变（见过 AP / AR / AY / BB），
        ///      所以顺着公式找，不能写死坐标。
        /// </summary>
        public static DateTime? ProbeReportTime(string path)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    var map = SheetEntryMap(zip);
                    string titleEntry, autoEntry;
                    if (!map.TryGetValue("通报", out titleEntry)) return null;

                    string titleXml = ReadEntryText(zip, titleEntry);
                    if (titleXml == null) return null;

                    string cellRef = FindTimeCellRef(titleXml);
                    if (cellRef == null) return null;

                    string autoXml = map.TryGetValue("auto", out autoEntry)
                        ? ReadEntryText(zip, autoEntry)
                        : null;
                    if (autoXml == null) return null;

                    string raw = CellText(autoXml, cellRef, zip);
                    return ParseReportTime(raw);
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        static Dictionary<string, string> SheetEntryMap(ZipArchive zip)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string wbXml = ReadEntryText(zip, "xl/workbook.xml");
            string relXml = ReadEntryText(zip, "xl/_rels/workbook.xml.rels");
            if (wbXml == null || relXml == null) return map;

            var rels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var relDoc = new System.Xml.XmlDocument();
            relDoc.XmlResolver = null;
            relDoc.LoadXml(relXml);
            foreach (System.Xml.XmlNode n in relDoc.GetElementsByTagName("Relationship"))
            {
                var idAttr = n.Attributes["Id"];
                var tgtAttr = n.Attributes["Target"];
                if (idAttr != null && tgtAttr != null) rels[idAttr.Value] = tgtAttr.Value;
            }

            var wbDoc = new System.Xml.XmlDocument();
            wbDoc.XmlResolver = null;
            wbDoc.LoadXml(wbXml);
            foreach (System.Xml.XmlNode n in wbDoc.GetElementsByTagName("sheet"))
            {
                var nameAttr = n.Attributes["name"];
                var ridAttr = n.Attributes["r:id"];
                if (nameAttr == null || ridAttr == null) continue;
                string target;
                if (!rels.TryGetValue(ridAttr.Value, out target)) continue;
                map[nameAttr.Value] = NormalizeEntry(target);
            }
            return map;
        }

        static string NormalizeEntry(string target)
        {
            target = target.Replace('\\', '/');
            if (target.StartsWith("/")) return target.Substring(1);
            if (target.StartsWith("xl/")) return target;
            return "xl/" + target;
        }

        static string ReadEntryText(ZipArchive zip, string entryName)
        {
            var e = zip.GetEntry(entryName);
            if (e == null)
            {
                foreach (var cand in zip.Entries)
                    if (string.Equals(cand.FullName, entryName, StringComparison.OrdinalIgnoreCase)) { e = cand; break; }
            }
            if (e == null) return null;
            using (var s = e.Open())
            using (var r = new StreamReader(s, Encoding.UTF8, true))
                return r.ReadToEnd();
        }

        /// <summary>从「通报」表 A1 的公式里取出它引用 auto 表的哪个单元格，如 AP5。</summary>
        static string FindTimeCellRef(string titleXml)
        {
            var cell = Regex.Match(titleXml,
                "<c(?=[^>]*\\br=\"A1\")[^>]*>(.*?)</c>", RegexOptions.Singleline);
            if (!cell.Success) return null;

            var f = Regex.Match(cell.Groups[1].Value, "<f[^>]*>(.*?)</f>", RegexOptions.Singleline);
            if (!f.Success) return null;

            var refs = Regex.Matches(f.Groups[1].Value, "!\\s*\\$?([A-Z]{1,3})\\$?(\\d+)");
            if (refs.Count == 0) return null;
            var last = refs[refs.Count - 1];
            return last.Groups[1].Value + last.Groups[2].Value;
        }

        /// <summary>取一个单元格的文本值（inlineStr / 共享字符串 / 数字都认）。</summary>
        static string CellText(string sheetXml, string cellRef, ZipArchive zip)
        {
            var m = Regex.Match(sheetXml,
                "<c(?=[^>]*\\br=\"" + cellRef + "\")([^>]*)(?:/>|>(.*?)</c>)",
                RegexOptions.Singleline);
            if (!m.Success) return null;

            string attrs = m.Groups[1].Value;
            string body = m.Groups[2].Value;
            if (string.IsNullOrEmpty(body)) return null;

            var t = Regex.Match(body, "<t[^>]*>(.*?)</t>", RegexOptions.Singleline);
            if (t.Success) return Unescape(t.Groups[1].Value);

            var v = Regex.Match(body, "<v[^>]*>(.*?)</v>", RegexOptions.Singleline);
            if (!v.Success) return null;

            string raw = Unescape(v.Groups[1].Value);
            if (raw == null || raw.Length == 0) return null;

            if (Regex.IsMatch(attrs, "\\bt=\"s\""))
            {
                int idx;
                if (!int.TryParse(raw.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out idx)) return null;
                var shared = ReadSharedStrings(zip);
                if (idx < 0 || idx >= shared.Length) return null;
                return shared[idx];
            }
            return raw;
        }

        static string[] ReadSharedStrings(ZipArchive zip)
        {
            string xml = ReadEntryText(zip, "xl/sharedStrings.xml");
            if (string.IsNullOrEmpty(xml)) return new string[0];

            var list = new List<string>();
            foreach (Match si in Regex.Matches(xml, "<si[^>]*>(.*?)</si>", RegexOptions.Singleline))
            {
                var sb = new StringBuilder();
                foreach (Match t in Regex.Matches(si.Groups[1].Value, "<t[^>]*>(.*?)</t>", RegexOptions.Singleline))
                    sb.Append(Unescape(t.Groups[1].Value));
                list.Add(sb.ToString());
            }
            return list.ToArray();
        }

        static string Unescape(string s)
        {
            if (s == null) return null;
            if (s.IndexOf('&') < 0) return s;
            return s.Replace("&lt;", "<").Replace("&gt;", ">")
                    .Replace("&quot;", "\"").Replace("&apos;", "'")
                    .Replace("&amp;", "&");
        }

        static readonly Regex TimeRx = new Regex(
            "(\\d{4})\\s*[年\\-/\\.]\\s*(\\d{1,2})\\s*[月\\-/\\.]\\s*(\\d{1,2})\\s*日?\\s*(\\d{1,2})\\s*[:：]\\s*(\\d{1,2})",
            RegexOptions.Compiled);

        /// <summary>把「2026年11月19日 17:15」这类文本转成时间。认不出返回 null。</summary>
        public static DateTime? ParseReportTime(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;

            var m = TimeRx.Match(text);
            if (m.Success)
            {
                try
                {
                    return new DateTime(
                        int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture),
                        int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture),
                        int.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture),
                        int.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture),
                        int.Parse(m.Groups[5].Value, CultureInfo.InvariantCulture), 0);
                }
                catch (ArgumentOutOfRangeException) { return null; }
            }

            // 少数文件这里是 Excel 日期序列号
            double serial;
            if (double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out serial)
                && serial > 20000 && serial < 80000)
                return DateTime.FromOADate(serial);

            return null;
        }

        /// <summary>
        /// 文件可能正在下载。等它大小稳定、并且能读出 zip 头，才算可用。
        /// 最多等 timeoutMs，超时返回 false。
        /// </summary>
        public static bool WaitUntilReady(string path, int timeoutMs)
        {
            long lastSize = -1;
            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    var fi = new FileInfo(path);
                    if (!fi.Exists) return false;
                    if (fi.Length == lastSize && fi.Length > 0)
                    {
                        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            var buf = new byte[4];
                            fs.Read(buf, 0, 4);
                            // xlsx 是 zip，头 4 字节固定是 50 4B 03 04
                            if (buf[0] == 0x50 && buf[1] == 0x4B) return true;
                        }
                    }
                    lastSize = fi.Length;
                }
                catch (IOException)
                {
                    // 还被占用，继续等
                }
                catch (Exception)
                {
                    return false;
                }
                System.Threading.Thread.Sleep(250);
            }
            return false;
        }
    }
}