using System;
using System.Drawing;
using System.Windows.Forms;
using ReportTool.App;

namespace ReportTool.Tools
{
    /// <summary>离屏渲染主窗口，用于无人值守时检查界面布局。</summary>
    public static class UiShot
    {
        [STAThread]
        public static int Main(string[] args)
        {
            string outPath = args.Length > 0 ? args[0] : "ui.png";
            Application.EnableVisualStyles();

            // 带 --data 时先跑一遍真实数据，把结果表填进界面再截图
            BatchResult data = null;
            var inputs = new System.Collections.Generic.List<string>();
            for (int i = 1; i < args.Length; i++) inputs.Add(args[i]);
            if (inputs.Count > 0)
            {
                data = Pipeline.Run(inputs, System.IO.Path.Combine(System.IO.Path.GetTempPath(), "reporttool_ui"),
                                    delegate(string s) { Console.WriteLine(s); });
            }

            using (var f = new MainForm())
            {
                if (data != null) f.ShowResult(data);
                var timer = new Timer();
                timer.Interval = 1200;
                timer.Tick += delegate
                {
                    timer.Stop();
                    try
                    {
                        using (var bmp = new Bitmap(f.Width, f.Height))
                        {
                            f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                            bmp.Save(outPath, System.Drawing.Imaging.ImageFormat.Png);
                        }
                        // DrawToBitmap 抓不到 DataGridView 的内容，单独抓一份看看
                        try
                        {
                            int n = 0;
                            foreach (Control c in f.Controls)
                            {
                                var sp = c as SplitContainer;
                                if (sp == null) continue;
                                foreach (Control pc in sp.Panel2.Controls)
                                {
                                    var tc = pc as TabControl;
                                    if (tc == null) continue;
                                    foreach (TabPage page in tc.TabPages)
                                    {
                                        foreach (Control gc in page.Controls)
                                        {
                                            var grid = gc as DataGridView;
                                            if (grid == null) continue;
                                            using (var gb = new Bitmap(grid.Width, grid.Height))
                                            {
                                                grid.DrawToBitmap(gb, new Rectangle(0, 0, grid.Width, grid.Height));
                                                gb.Save(outPath + ".grid" + (n++) + ".png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                            if (n >= 2) break;
                                        }
                                        if (n >= 2) break;
                                    }
                                }
                            }
                            Console.WriteLine("grid shots: " + n);
                        }
                        catch (Exception gex) { Console.WriteLine("grid shot failed: " + gex.Message); }
                        Console.WriteLine("saved " + outPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("capture failed: " + ex.Message);
                    }
                    Dump(f, 0);
                    f.Close();
                };
                timer.Start();
                Application.Run(f);
            }
            return 0;
        }

        static void Dump(Control c, int depth)
        {
            if (depth > 4) return;
            Console.WriteLine(new string(' ', depth * 2) + c.GetType().Name +
                              " bounds=" + c.Bounds + " dock=" + c.Dock + " visible=" + c.Visible);
            var tc = c as TabControl;
            if (tc != null)
            {
                foreach (TabPage p in tc.TabPages)
                {
                    Console.WriteLine("  [TAB] " + p.Text);
                    foreach (Control pc in p.Controls)
                    {
                        var g = pc as DataGridView;
                        if (g == null) continue;
                        Console.WriteLine("    cols=" + g.Columns.Count + " rows=" + g.Rows.Count);
                        var head = new System.Text.StringBuilder("    HDR: ");
                        foreach (DataGridViewColumn col in g.Columns) head.Append(col.HeaderText + " | ");
                        Console.WriteLine(head.ToString());
                        for (int r = 0; r < Math.Min(2, g.Rows.Count); r++)
                        {
                            var line = new System.Text.StringBuilder("    R" + r + ": ");
                            foreach (DataGridViewCell cell in g.Rows[r].Cells) line.Append(cell.Value + " | ");
                            Console.WriteLine(line.ToString());
                        }
                    }
                }
            }
            foreach (Control ch in c.Controls) Dump(ch, depth + 1);
        }
    }
}
