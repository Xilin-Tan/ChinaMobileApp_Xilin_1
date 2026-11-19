using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace ReportTool.App
{
    /// <summary>一条数据系列。值都是原始数值（完成率是 0~1 的小数，不是字符串）。</summary>
    public class ChartSeries
    {
        public string Name;
        public Color Color;
        public double[] Values;
    }

    /// <summary>
    /// 一张图要画什么。由 Summary 从重算后的工作簿里抽出来，
    /// 数值一律取 CellValue.Num，不从界面表格的字符串反解。
    /// </summary>
    public class ChartData
    {
        public string Title;
        /// <summary>副标题，一般放通报时间。</summary>
        public string Subtitle;
        public string[] Categories;
        public List<ChartSeries> Series = new List<ChartSeries>();

        /// <summary>true = 横向条形（类别多、名字长时用）；false = 分组竖向柱状。</summary>
        public bool Horizontal;

        /// <summary>值是不是百分比（0~1）。</summary>
        public bool Percent;

        /// <summary>参考线的位置，如 1.0 表示 100% 达标线。ShowReference 为 false 时不画。</summary>
        public double Reference;
        public bool ShowReference;

        /// <summary>横向条形最多显示多少条（从高到低）。0 表示全部显示。</summary>
        public int MaxBars;

        /// <summary>true 时 MaxBars 取"最落后的几条"而不是"最靠前的几条"。</summary>
        public bool TailBars;
    }

    /// <summary>
    /// 用 System.Drawing 画图，不引入任何第三方库。
    /// 输出 PNG，既能放进界面预览，也能直接发企业微信群、在手机上看。
    /// </summary>
    public static class ChartRenderer
    {
        static readonly Color Ink = Color.FromArgb(38, 42, 48);
        static readonly Color Muted = Color.FromArgb(120, 128, 138);
        static readonly Color Grid = Color.FromArgb(226, 230, 236);
        static readonly Color Good = Color.FromArgb(46, 155, 92);
        static readonly Color Warn = Color.FromArgb(232, 152, 40);
        static readonly Color Bad = Color.FromArgb(214, 69, 69);
        static readonly Color Base = Color.FromArgb(52, 116, 214);

        /// <summary>达标了就绿、接近就黄、差得多就红。用颜色替代看数字。</summary>
        public static Color RateColor(double rate)
        {
            if (rate >= 1.0) return Good;
            if (rate >= 0.8) return Warn;
            return Bad;
        }

        public static Bitmap Render(ChartData d, int width, int height)
        {
            var bmp = new Bitmap(width, height);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                g.Clear(Color.White);

                using (var titleFont = new Font("Microsoft YaHei UI", 15F, FontStyle.Bold))
                using (var subFont = new Font("Microsoft YaHei UI", 9F))
                using (var labelFont = new Font("Microsoft YaHei UI", 9F))
                using (var valueFont = new Font("Microsoft YaHei UI", 8.5F))
                {
                    float y = 16;
                    g.DrawString(d.Title, titleFont, new SolidBrush(Ink), 20, y);
                    y += 28;
                    if (!string.IsNullOrEmpty(d.Subtitle))
                    {
                        g.DrawString(d.Subtitle, subFont, new SolidBrush(Muted), 20, y);
                    }

                    if (d.Series.Count == 0 || d.Categories == null || d.Categories.Length == 0)
                    {
                        g.DrawString("没有可画的数据", subFont, new SolidBrush(Muted), 20, y + 10);
                        return bmp;
                    }

                    var area = new Rectangle(0, (int)y + 34, width, height - (int)y - 34);
                    if (d.Horizontal)
                        DrawHorizontal(g, d, area, labelFont, valueFont);
                    else
                        DrawGrouped(g, d, area, labelFont, valueFont);
                }
            }
            return bmp;
        }

        // ---------------- 横向条形 ----------------

        static void DrawHorizontal(Graphics g, ChartData d, Rectangle area, Font labelFont, Font valueFont)
        {
            // 只画一条系列：横向条形本来就是"每个类别一个值"
            var vals = d.Series[0].Values;
            var names = d.Categories;
            int n = Math.Min(vals.Length, names.Length);
            if (n == 0) return;

            // 全部画；MaxBars 只限制"最低的那几行"的显示条数，避免图太挤
            int count = n;
            int start = 0;
            if (d.MaxBars > 0 && n > d.MaxBars)
            {
                count = d.MaxBars;
                if (d.TailBars) start = n - count;   // 取最落后的几条
            }

            float left = 138;                       // 左边留给类别名
            float right = area.Width - 78;          // 右边留给数值
            float top = area.Y + 8;
            float bottom = area.Bottom - 30;

            double max = 0;
            for (int i = start; i < start + count; i++) if (vals[i] > max) max = vals[i];
            if (d.ShowReference && d.Reference > max) max = d.Reference;
            if (max <= 0) max = 1;
            double axisMax = NiceCeil(max);
            double cap = AxisCap(d);
            if (axisMax > cap) axisMax = cap;   // 异常值不把整张图拉扁

            DrawXAxis(g, d, left, right, top, bottom, axisMax, labelFont);

            float rowH = (bottom - top) / count;
            float barH = Math.Min(rowH * 0.68f, 26f);

            for (int i = start; i < start + count; i++)
            {
                float cy = top + rowH * (i - start) + rowH / 2f;
                double v = vals[i];
                float len = (float)((right - left) * Math.Min(v / axisMax, 1.0));

                Color c = d.Percent ? RateColor(v) : Base;
                var bar = new RectangleF(left, cy - barH / 2f, Math.Max(len, 1.5f), barH);
                FillRound(g, bar, c, 3);

                // 超出坐标轴上限：条形画满，末尾加一个箭头，值标签仍显示真实数值
                if (v > axisMax)
                {
                    using (var tri = new SolidBrush(c))
                    {
                        var pts = new[]
                        {
                            new PointF(right + 2, cy - 5),
                            new PointF(right + 10, cy),
                            new PointF(right + 2, cy + 5)
                        };
                        g.FillPolygon(tri, pts);
                    }
                }

                using (var nameBrush = new SolidBrush(Ink))
                {
                    var sf = new StringFormat();
                    sf.Alignment = StringAlignment.Far;
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Trimming = StringTrimming.EllipsisCharacter;
                    g.DrawString(Shorten(names[i], 12), labelFont, nameBrush,
                                 new RectangleF(6, cy - rowH / 2f, left - 14, rowH), sf);
                }

                // 数值标签：超过轴上限（例如 229.9%）时贴在图内右端
                string text = FormatValue(v, d.Percent);
                float tx = Math.Min(left + len + 6, right + 6);
                using (var vb = new SolidBrush(Ink))
                using (var sf2 = new StringFormat())
                {
                    sf2.LineAlignment = StringAlignment.Center;
                    g.DrawString(text, valueFont, vb, new RectangleF(tx, cy - rowH / 2f, 70, rowH), sf2);
                }
            }
        }

        // ---------------- 分组竖向柱状 ----------------

        static void DrawGrouped(Graphics g, ChartData d, Rectangle area, Font labelFont, Font valueFont)
        {
            int cats = d.Categories.Length;
            int series = d.Series.Count;
            if (cats == 0 || series == 0) return;

            float left = area.X + 64;
            float right = area.Right - 24;
            float top = area.Y + 46;          // 顶部留给图例
            float bottom = area.Bottom - 42;

            double max = 0;
            foreach (var s in d.Series)
                foreach (var v in s.Values)
                    if (v > max) max = v;
            if (d.ShowReference && d.Reference > max) max = d.Reference;
            if (max <= 0) max = 1;
            double axisMax = NiceCeil(max);
            double cap = AxisCap(d);
            if (axisMax > cap) axisMax = cap;

            DrawLegend(g, d, left, area.Y + 6, labelFont);
            DrawYAxis(g, d, left, right, top, bottom, axisMax, labelFont);

            float slotW = (right - left) / cats;
            float barW = Math.Max(4f, slotW * 0.72f / series);
            float groupW = barW * series;

            for (int i = 0; i < cats; i++)
            {
                float cx = left + slotW * i + slotW / 2f;

                for (int k = 0; k < series; k++)
                {
                    double v = k < d.Series[k].Values.Length ? d.Series[k].Values[i] : 0;
                    float h = (float)((bottom - top) * Math.Min(v / axisMax, 1.0));
                    float x = cx - groupW / 2f + barW * k;
                    var bar = new RectangleF(x, bottom - h, barW - 2, Math.Max(h, 1.5f));
                    FillRound(g, bar, d.Series[k].Color, 2);
                }

                var sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Near;
                using (var b = new SolidBrush(Ink))
                    g.DrawString(Shorten(d.Categories[i], 6), labelFont, b,
                                 new RectangleF(cx - slotW / 2f, bottom + 8, slotW, 26), sf);
            }
        }

        // ---------------- 轴与图例 ----------------

        static void DrawXAxis(Graphics g, ChartData d, float left, float right, float top, float bottom,
                              double axisMax, Font font)
        {
            using (var pen = new Pen(Grid, 1))
            {
                // 竖向网格 + 刻度（横向条形时，X 轴是数值）
                const int steps = 4;
                for (int i = 0; i <= steps; i++)
                {
                    double v = axisMax * i / steps;
                    float x = left + (right - left) * i / steps;
                    g.DrawLine(pen, x, top, x, bottom);

                    var sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    using (var b = new SolidBrush(Muted))
                        g.DrawString(FormatValue(v, d.Percent), font, b,
                                     new RectangleF(x - 34, bottom + 4, 68, 20), sf);
                }
            }
            DrawReferenceVertical(g, d, left, right, top, bottom, axisMax, font);
        }

        static void DrawYAxis(Graphics g, ChartData d, float left, float right, float top, float bottom,
                              double axisMax, Font font)
        {
            using (var pen = new Pen(Grid, 1))
            {
                const int steps = 4;
                for (int i = 0; i <= steps; i++)
                {
                    double v = axisMax * i / steps;
                    float y = bottom - (bottom - top) * i / steps;
                    g.DrawLine(pen, left, y, right, y);

                    var sf = new StringFormat();
                    sf.Alignment = StringAlignment.Far;
                    sf.LineAlignment = StringAlignment.Center;
                    using (var b = new SolidBrush(Muted))
                        g.DrawString(FormatValue(v, d.Percent), font, b,
                                     new RectangleF(6, y - 9, left - 12, 18), sf);
                }
            }
            DrawReferenceHorizontal(g, d, left, right, top, bottom, axisMax, font);
        }

        static void DrawReferenceVertical(Graphics g, ChartData d, float left, float right,
                                          float top, float bottom, double axisMax, Font font)
        {
            if (!d.ShowReference) return;
            float x = left + (float)((right - left) * d.Reference / axisMax);
            using (var pen = new Pen(Color.FromArgb(150, 90, 90, 110), 1.5f))
            {
                pen.DashStyle = DashStyle.Dash;
                g.DrawLine(pen, x, top - 4, x, bottom);
            }
            var sf = new StringFormat();
            sf.Alignment = StringAlignment.Center;
            using (var b = new SolidBrush(Color.FromArgb(110, 90, 90, 110)))
            {
                float tx = x - 46;
                if (tx + 92 > right) tx = right - 92;   // 别被右边界切掉
                if (tx < left) tx = left;
                g.DrawString(FormatValue(d.Reference, d.Percent) + " 达标线", font, b,
                             new RectangleF(tx, top - 20, 92, 18), sf);
            }
        }

        static void DrawReferenceHorizontal(Graphics g, ChartData d, float left, float right,
                                            float top, float bottom, double axisMax, Font font)
        {
            if (!d.ShowReference) return;
            float y = bottom - (float)((bottom - top) * d.Reference / axisMax);
            using (var pen = new Pen(Color.FromArgb(150, 90, 90, 110), 1.5f))
            {
                pen.DashStyle = DashStyle.Dash;
                g.DrawLine(pen, left, y, right, y);
            }
            var sf = new StringFormat();
            sf.Alignment = StringAlignment.Far;
            sf.LineAlignment = StringAlignment.Far;
            using (var b = new SolidBrush(Color.FromArgb(110, 90, 90, 110)))
                g.DrawString("100% 达标线", font, b, new RectangleF(right - 82, y - 20, 80, 18), sf);
        }

        static void DrawLegend(Graphics g, ChartData d, float x, float y, Font font)
        {
            float cx = x;
            foreach (var s in d.Series)
            {
                var sw = new RectangleF(cx, y + 4, 12, 12);
                FillRound(g, sw, s.Color, 2);
                var sz = g.MeasureString(s.Name, font);
                using (var b = new SolidBrush(Ink))
                    g.DrawString(s.Name, font, b, cx + 17, y);
                cx += 17 + sz.Width + 18;
            }
        }

        // ---------------- 小工具 ----------------

        /// <summary>这张图自然需要多高。横向条形按条数算，柱状固定。</summary>
        public static int HeightOf(ChartData d)
        {
            if (!d.Horizontal) return 420;
            int n = d.MaxBars > 0 ? Math.Min(d.Categories.Length, d.MaxBars) : d.Categories.Length;
            return Math.Max(200, 118 + 26 * n);
        }

        public static void Save(ChartData d, string path)
        {
            using (var bmp = Render(d, 900, HeightOf(d)))
            {
                string dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                bmp.Save(path, ImageFormat.Png);
            }
        }

        /// <summary>把几张图竖着拼成一张，适合一次性发到群里看。</summary>
        public static void SaveStack(List<ChartData> charts, string path, int width)
        {
            if (charts == null || charts.Count == 0) return;

            var bitmaps = new List<Bitmap>();
            int totalH = 0;
            try
            {
                foreach (var original in charts)
                {
                    // 一页纸里，横向条形最多放 14 条，否则整页被一张图撑爆。
                    var c = original;
                    if (original.Horizontal && original.Categories.Length > 14 && original.MaxBars == 0)
                    {
                        // 日报是督办用的：一页纸里只看最落后的那 14 个
                        c = new ChartData
                        {
                            Title = original.Title,
                            Subtitle = original.Subtitle + " · 一页纸只画最落后的 14 个，完整见单独的图",
                            Categories = original.Categories,
                            Horizontal = true,
                            Percent = original.Percent,
                            Reference = original.Reference,
                            ShowReference = original.ShowReference,
                            MaxBars = 14,
                            TailBars = true
                        };
                        c.Series.Add(original.Series[0]);
                    }
                    var bmp = Render(c, width, HeightOf(c));
                    bitmaps.Add(bmp);
                    totalH += bmp.Height;
                }
                totalH += 24 * (bitmaps.Count - 1);

                using (var canvas = new Bitmap(width, totalH))
                using (var g = Graphics.FromImage(canvas))
                {
                    g.Clear(Color.FromArgb(244, 246, 249));
                    int y = 0;
                    foreach (var b in bitmaps)
                    {
                        g.DrawImage(b, 0, y);
                        y += b.Height + 24;
                    }
                    string dir = Path.GetDirectoryName(path);
                    if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                    canvas.Save(path, ImageFormat.Png);
                }
            }
            finally
            {
                foreach (var b in bitmaps) b.Dispose();
            }
        }

        /// <summary>
        /// 坐标轴上限。有达标线的百分比图最多画到 200%：
        /// 一个 416% 的异常网格不该把其余三十几个压成一条线。
        /// </summary>
        static double AxisCap(ChartData d)
        {
            if (d.Percent && d.ShowReference && d.Reference > 0) return d.Reference * 2;
            return double.MaxValue;
        }

        static void FillRound(Graphics g, RectangleF r, Color c, int radius)
        {
            using (var path = new GraphicsPath())
            {
                float d2 = radius * 2;
                if (r.Width < d2) d2 = r.Width;
                path.AddArc(r.X, r.Y, d2, d2, 180, 90);
                path.AddArc(r.Right - d2, r.Y, d2, d2, 270, 90);
                path.AddArc(r.Right - d2, r.Bottom - d2, d2, d2, 0, 90);
                path.AddArc(r.X, r.Bottom - d2, d2, d2, 90, 90);
                path.CloseFigure();
                using (var b = new SolidBrush(c)) g.FillPath(b, path);
            }
        }

        /// <summary>把轴上限收成好看的整数：120% / 150% / 240% 这种。</summary>
        static double NiceCeil(double v)
        {
            if (v <= 0) return 1;
            double[] steps = { 1, 1.2, 1.5, 2, 2.4, 2.5, 3, 4, 5, 6, 8, 10 };
            // 百分比按 0.2 的粒度向上取；普通数按量级取
            if (v <= 1)
            {
                double[] pct = { 0.2, 0.4, 0.6, 0.8, 1.0, 1.2, 1.5, 2.0, 2.5, 3.0 };
                foreach (var s in pct) if (v <= s) return s;
                return Math.Ceiling(v / 0.5) * 0.5;
            }
            foreach (var s in steps) if (v <= s) return s;
            double mag = Math.Pow(10, Math.Floor(Math.Log10(v)));
            return Math.Ceiling(v / mag) * mag;
        }

        public static string FormatValue(double v, bool percent)
        {
            if (percent) return (v * 100).ToString("0.#") + "%";
            if (Math.Abs(v - Math.Round(v)) < 1e-9) return ((long)Math.Round(v)).ToString("#,0");
            return v.ToString("#,0.#");
        }

        static string Shorten(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Replace("\r", "").Replace("\n", "").Trim();
            return s.Length <= max ? s : s.Substring(0, max) + "…";
        }
    }
}