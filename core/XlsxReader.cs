using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ReportTool.Core
{
    /// <summary>
    /// 零依赖 xlsx 读取器：直接解 zip + 解析 OOXML，不依赖 Excel 或任何第三方库。
    /// 只读取值（含共享字符串、内联字符串、数字、布尔），公式单独存放。
    /// </summary>
    public static class XlsxReader
    {
        public static Workbook Read(string path)
        {
            var wb = new Workbook { Path = path };
            using (var fs = File.OpenRead(path))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var workbookXml = ReadEntry(zip, "xl/workbook.xml");
                if (workbookXml == null) throw new InvalidDataException("不是有效的 xlsx 文件：" + path);

                var relsXml = ReadEntry(zip, "xl/_rels/workbook.xml.rels") ?? "";
                var relMap = ParseRelationships(relsXml);
                var sharedStrings = ReadSharedStrings(zip);

                var doc = LoadXml(workbookXml);
                var ns = Ns(doc);
                var sheetsNode = doc.SelectNodes("//m:sheets/m:sheet", ns);
                var order = new List<string>();

                if (sheetsNode != null)
                {
                    foreach (XmlNode n in sheetsNode)
                    {
                        string name = Attr(n, "name");
                        string rid = Attr(n, "r:id");
                        if (name == null) continue;
                        order.Add(name);

                        string target = null;
                        if (rid != null) relMap.TryGetValue(rid, out target);
                        if (target == null) continue;

                        string entry = NormalizeSheetPath(target);
                        var sheet = new Sheet { Name = name };
                        var sheetXml = ReadEntry(zip, entry);
                        if (sheetXml != null) ParseSheet(sheet, sheetXml, sharedStrings);
                        wb.Sheets[name] = sheet;
                        wb.SheetEntryPaths[name] = entry;
                        if (wb.FirstSheetName == null) wb.FirstSheetName = name;
                    }
                }
            }

            // 没有公式的表（明细表）在重算期间不会变化，开启整列缓存加速 SUMIFS。
            foreach (var kv in wb.Sheets)
            {
                var sh = kv.Value;
                if (sh.Formulas.Count == 0) sh.FastCols = new Dictionary<int, CellValue[]>();
            }
            return wb;
        }

        static string NormalizeSheetPath(string target)
        {
            target = target.Replace('\\', '/');
            if (target.StartsWith("/")) return target.Substring(1);
            if (target.StartsWith("xl/")) return target;
            return "xl/" + target;
        }

        static Dictionary<string, string> ParseRelationships(string xml)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(xml)) return map;
            var doc = LoadXml(xml);
            foreach (XmlNode n in doc.GetElementsByTagName("Relationship"))
            {
                string id = Attr(n, "Id");
                string tgt = Attr(n, "Target");
                if (id != null && tgt != null) map[id] = tgt;
            }
            return map;
        }

        static string[] ReadSharedStrings(ZipArchive zip)
        {
            string xml = ReadEntry(zip, "xl/sharedStrings.xml");
            if (string.IsNullOrEmpty(xml)) return new string[0];
            var doc = LoadXml(xml);
            var ns = Ns(doc);
            var list = new List<string>();
            var nodes = doc.SelectNodes("//m:si", ns);
            if (nodes != null)
            {
                foreach (XmlNode si in nodes)
                {
                    var sb = new StringBuilder();
                    var ts = si.SelectNodes(".//m:t", ns);
                    if (ts != null) foreach (XmlNode t in ts) sb.Append(t.InnerText);
                    list.Add(sb.ToString());
                }
            }
            return list.ToArray();
        }

        static void ParseSheet(Sheet sheet, string xml, string[] sharedStrings)
        {
            var doc = LoadXml(xml);
            var ns = Ns(doc);
            var rows = doc.SelectNodes("//m:sheetData/m:row", ns);
            if (rows == null) return;

            foreach (XmlNode rowNode in rows)
            {
                int rowIndex = ParseInt(Attr(rowNode, "r"), 0);
                if (rowIndex <= 0) continue;
                foreach (XmlNode c in rowNode.ChildNodes)
                {
                    if (c.Name != "c" && !c.Name.EndsWith(":c")) continue;
                    string cellRef = Attr(c, "r");
                    if (cellRef == null) continue;
                    int col = ColIndex(cellRef);
                    if (col <= 0) continue;

                    string type = Attr(c, "t") ?? "n";
                    string formula = null;
                    string rawValue = null;
                    string inlineText = null;
                    bool hasChildIs = false;

                    foreach (XmlNode ch in c.ChildNodes)
                    {
                        string local = LocalName(ch.Name);
                        if (local == "f")
                        {
                            // 共享公式（t="shared"）只存主体、其余靠位移推导，本引擎不支持。
                            // 与其算错，不如明确报出来。
                            if (Attr(ch, "t") == "shared")
                                throw new NotSupportedException(
                                    "工作表「" + sheet.Name + "」的 " + cellRef +
                                    " 使用了共享公式，当前引擎不支持。请先在 Excel 里另存一份再试。");
                            formula = ch.InnerText;
                        }
                        else if (local == "v") rawValue = ch.InnerText;
                        else if (local == "is")
                        {
                            hasChildIs = true;
                            var sb = new StringBuilder();
                            var ts = ch.SelectNodes(".//m:t", ns);
                            if (ts != null) foreach (XmlNode t in ts) sb.Append(t.InnerText);
                            inlineText = sb.ToString();
                        }
                    }

                    CellValue value = CellValue.Blank;
                    if (type == "s")
                    {
                        int idx;
                        if (rawValue != null && int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx)
                            && idx >= 0 && idx < sharedStrings.Length)
                            value = CellValue.FromText(sharedStrings[idx]);
                    }
                    else if (type == "inlineStr" || hasChildIs)
                    {
                        value = CellValue.FromText(inlineText ?? "");
                    }
                    else if (type == "str")
                    {
                        value = CellValue.FromText(rawValue ?? "");
                    }
                    else if (type == "b")
                    {
                        value = CellValue.FromBool(rawValue == "1");
                    }
                    else if (type == "e")
                    {
                        value = CellValue.Error(rawValue ?? "#ERR");
                    }
                    else
                    {
                        double d;
                        if (rawValue != null && !rawValue.Trim().Equals("", StringComparison.Ordinal)
                            && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                            value = CellValue.FromNumber(d);
                    }

                    if (formula != null) sheet.Formulas[Sheet.Key(rowIndex, col)] = formula;
                    if (value.Kind != ValueKind.Blank || formula == null) sheet.Set(rowIndex, col, value);
                    else
                    {
                        // 公式但无缓存值：登记行列范围，值保持空白
                        if (rowIndex > sheet.MaxRow) sheet.MaxRow = rowIndex;
                        if (col > sheet.MaxCol) sheet.MaxCol = col;
                    }
                }
            }
        }

        static XmlDocument LoadXml(string xml)
        {
            var doc = new XmlDocument();
            doc.XmlResolver = null;
            doc.LoadXml(xml);
            return doc;
        }

        static XmlNamespaceManager Ns(XmlDocument doc)
        {
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            return ns;
        }

        static string LocalName(string name)
        {
            int i = name.IndexOf(':');
            return i < 0 ? name : name.Substring(i + 1);
        }

        static string Attr(XmlNode n, string name)
        {
            if (n.Attributes == null) return null;
            var a = n.Attributes[name];
            if (a != null) return a.Value;
            foreach (XmlAttribute at in n.Attributes)
            {
                if (LocalName(at.Name) == LocalName(name)) return at.Value;
            }
            return null;
        }

        static int ParseInt(string s, int fallback)
        {
            int v;
            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out v) ? v : fallback;
        }

        public static int ColIndex(string cellRef)
        {
            int col = 0, i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i]))
            {
                col = col * 26 + (char.ToUpperInvariant(cellRef[i]) - 'A' + 1);
                i++;
            }
            return col;
        }

        public static int RowIndex(string cellRef)
        {
            int i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
            int r = 0;
            while (i < cellRef.Length && char.IsDigit(cellRef[i]))
            {
                r = r * 10 + (cellRef[i] - '0');
                i++;
            }
            return r;
        }

        static string ReadEntry(ZipArchive zip, string entryName)
        {
            var e = zip.GetEntry(entryName);
            if (e == null)
            {
                foreach (var cand in zip.Entries)
                {
                    if (string.Equals(cand.FullName, entryName, StringComparison.OrdinalIgnoreCase))
                    { e = cand; break; }
                }
            }
            if (e == null) return null;
            using (var s = e.Open())
            using (var r = new StreamReader(s, Encoding.UTF8, true))
                return r.ReadToEnd();
        }
    }
}
