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
    /// 把重算后的数值写回 xlsx：原文件的所有格式（合并单元格、列宽、配色、
    /// 条件格式、行高）原样保留，只把「公式格」换成计算结果。
    /// 做法是复制整个 zip，仅替换工作表的 xml 内容，其余条目逐字节照搬。
    /// </summary>
    public static class XlsxWriter
    {
        const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        public static void Write(Workbook wb, string outputPath)
        {
            if (wb == null) throw new ArgumentNullException("wb");

            // 每个表生成新的 xml（只有含公式的表需要改）
            var replaced = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in wb.Sheets)
            {
                string entry;
                if (!wb.SheetEntryPaths.TryGetValue(kv.Key, out entry)) continue;
                var sheet = kv.Value;
                if (sheet.Formulas.Count == 0) continue;
                replaced[entry] = BuildSheetXml(wb, sheet, entry);
            }

            // 去掉 fullCalcOnLoad：否则 Excel 打开时会把全部公式再算一遍（很慢）。
            // 我们已经把结果写成缓存值，直接显示即可。
            foreach (var name in new[] { "xl/workbook.xml" })
            {
                string text = ReadEntryText(wb.Path, name);
                string patched = text.Replace(" fullCalcOnLoad=\"1\"", "")
                                     .Replace("fullCalcOnLoad=\"1\"", "");
                if (patched != text) replaced[name] = new UTF8Encoding(false).GetBytes(patched);
            }

            string dir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

            string tmp = outputPath + ".tmp";
            if (File.Exists(tmp)) File.Delete(tmp);

            using (var src = ZipFile.OpenRead(wb.Path))
            using (var dstStream = File.Create(tmp))
            using (var dst = new ZipArchive(dstStream, ZipArchiveMode.Create))
            {
                foreach (var entry in src.Entries)
                {
                    byte[] blob;
                    if (!replaced.TryGetValue(entry.FullName, out blob))
                        blob = ReadAll(entry);

                    var ne = dst.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                    using (var s = ne.Open())
                        s.Write(blob, 0, blob.Length);
                }
            }

            if (File.Exists(outputPath)) File.Delete(outputPath);
            File.Move(tmp, outputPath);
        }

        static byte[] ReadAll(ZipArchiveEntry entry)
        {
            using (var s = entry.Open())
            using (var ms = new MemoryStream())
            {
                s.CopyTo(ms);
                return ms.ToArray();
            }
        }

        static byte[] BuildSheetXml(Workbook wb, Sheet sheet, string entryName)
        {
            string xml = ReadEntryText(wb.Path, entryName);
            var doc = new XmlDocument();
            doc.XmlResolver = null;
            doc.PreserveWhitespace = true;
            doc.LoadXml(xml);

            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("m", NsMain);

            var rows = doc.SelectNodes("//m:sheetData/m:row", ns);
            if (rows != null)
            {
                foreach (XmlNode row in rows)
                {
                    int rowIndex;
                    if (!int.TryParse(Attr(row, "r"), out rowIndex)) continue;
                    foreach (XmlNode c in row.ChildNodes)
                    {
                        if (Local(c.Name) != "c") continue;
                        string cellRef = Attr(c, "r");
                        if (cellRef == null) continue;
                        if (!sheet.Formulas.ContainsKey(Sheet.Key(rowIndex, XlsxReader.ColIndex(cellRef)))) continue;
                        ApplyValue(c, sheet.Get(rowIndex, XlsxReader.ColIndex(cellRef)));
                    }
                }
            }

            var sw = new UTF8Encoding(false);
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings
                {
                    Encoding = sw,
                    OmitXmlDeclaration = true,
                    Indent = false
                };
                using (var xw = XmlWriter.Create(ms, settings))
                    doc.Save(xw);
                return ms.ToArray();
            }
        }

        /// <summary>把单元格里的公式换成数值，保持样式属性 s 不变。</summary>
        static void ApplyValue(XmlNode cell, CellValue value)
        {
            // 只清子节点，绝不能 RemoveAll() —— 那会连 r / s 属性一起删掉，格式全丢。
            var doomed = new List<XmlNode>();
            foreach (XmlNode ch in cell.ChildNodes)
            {
                string n = Local(ch.Name);
                if (n == "f" || n == "v" || n == "is") doomed.Add(ch);
            }
            foreach (var ch in doomed) cell.RemoveChild(ch);

            switch (value.Kind)
            {
                case ValueKind.Number:
                    SetAttr(cell, "t", null);
                    AppendText(cell, "v", value.Num.ToString("R", CultureInfo.InvariantCulture));
                    break;
                case ValueKind.Bool:
                    SetAttr(cell, "t", "b");
                    AppendText(cell, "v", value.Bool ? "1" : "0");
                    break;
                case ValueKind.Text:
                    if (string.IsNullOrEmpty(value.Text)) { SetAttr(cell, "t", null); break; }
                    SetAttr(cell, "t", "str");
                    AppendText(cell, "v", value.Text);
                    break;
                case ValueKind.Error:
                    SetAttr(cell, "t", "e");
                    AppendText(cell, "v", value.Text ?? "#VALUE!");
                    break;
                default:
                    SetAttr(cell, "t", null);
                    break;
            }
        }

        static void AppendText(XmlNode parent, string localName, string text)
        {
            var el = parent.OwnerDocument.CreateElement(localName, NsMain);
            el.InnerText = text;
            parent.AppendChild(el);
        }

        static string Local(string name)
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
                if (Local(at.Name) == name) return at.Value;
            return null;
        }

        static void SetAttr(XmlNode n, string name, string value)
        {
            var attrs = n.Attributes;
            if (value == null)
            {
                if (attrs != null && attrs[name] != null) attrs.RemoveNamedItem(name);
                return;
            }
            var doc = n.OwnerDocument;
            var a = attrs == null ? null : attrs[name];
            if (a == null)
            {
                a = doc.CreateAttribute(name);
                n.Attributes.Append(a);
            }
            a.Value = value;
        }

        static string ReadEntryText(string path, string entryName)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                var e = zip.GetEntry(entryName);
                if (e == null)
                    throw new InvalidDataException("找不到工作表：" + entryName);
                using (var s = e.Open())
                using (var r = new StreamReader(s, Encoding.UTF8, true))
                    return r.ReadToEnd();
            }
        }
    }
}
