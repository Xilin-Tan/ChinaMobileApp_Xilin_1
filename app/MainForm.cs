using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ReportTool.App
{
    public class MainForm : Form
    {
        readonly Panel _drop;
        readonly Label _dropText;
        readonly TextBox _log;
        readonly TabControl _tabs;
        readonly ToolStripStatusLabel _status;
        readonly ToolStripButton _btnOpen, _btnClear, _btnPick, _btnScan;
        readonly TabPage _chartPage;
        readonly PictureBox _chartBox;
        readonly Label _chartHint;
        readonly FlowLayoutPanel _chartBar;

        readonly List<string> _queue = new List<string>();
        string _chartPath;
        bool _busy;
        string _lastOutput;

        public MainForm()
        {
            Text = "报表生成工具";
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(1180, 760);
            MinimumSize = new Size(820, 520);
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.WhiteSmoke;

            _drop = new Panel();
            _drop.Dock = DockStyle.Top;
            _drop.Height = 108;
            _drop.AllowDrop = true;
            _drop.BackColor = Color.White;
            _drop.Padding = new Padding(12);
            _drop.Paint += DropPaint;
            _drop.DragEnter += DropDragEnter;
            _drop.DragDrop += DropDragDrop;
            _drop.DragLeave += DropDragLeave;
            _drop.Cursor = Cursors.Hand;
            _drop.Click += PickFiles;

            _dropText = new Label();
            _dropText.Dock = DockStyle.Fill;
            _dropText.TextAlign = ContentAlignment.MiddleCenter;
            _dropText.Font = new Font("Microsoft YaHei UI", 12F);
            _dropText.ForeColor = Color.FromArgb(70, 70, 70);
            _dropText.Text = "把 xlsx 文件（或放文件的文件夹）拖到这里\r\n区县 / 网格 / 重点业务，一次拖三个即可";
            _dropText.AllowDrop = true;
            _dropText.DragEnter += DropDragEnter;
            _dropText.DragDrop += DropDragDrop;
            _dropText.Cursor = Cursors.Hand;
            _dropText.Click += PickFiles;
            _drop.Controls.Add(_dropText);

            var go = new Button();
            go.Text = "一键取企业微信里的报表并生成";
            go.Height = 30;
            go.Dock = DockStyle.Bottom;
            go.FlatStyle = FlatStyle.System;
            go.Click += ScanWeChat;
            _drop.Controls.Add(go);

            var split = new SplitContainer();
            split.Dock = DockStyle.Fill;
            split.Orientation = Orientation.Vertical;
            // 必须先给足尺寸，否则设置 SplitterDistance / PanelMinSize 会直接抛异常。
            split.Size = new Size(1180, 600);
            split.SplitterWidth = 8;
            split.Panel1MinSize = 240;
            split.Panel2MinSize = 300;
            split.SplitterDistance = 430;
            split.BackColor = Color.FromArgb(200, 205, 212);
            split.Panel1.BackColor = Color.FromArgb(30, 30, 30);
            split.Panel2.BackColor = Color.White;

            _log = new TextBox();
            _log.Dock = DockStyle.Fill;
            _log.Multiline = true;
            _log.ReadOnly = true;
            _log.ScrollBars = ScrollBars.Vertical;
            _log.BackColor = Color.FromArgb(30, 30, 30);
            _log.ForeColor = Color.Gainsboro;
            _log.Font = new Font("Consolas", 9.5F);
            _log.BorderStyle = BorderStyle.None;
            split.Panel1.Controls.Add(_log);

            _tabs = new TabControl();
            _tabs.Dock = DockStyle.Fill;
            split.Panel2.Controls.Add(_tabs);

            // 「图表」页在构造时就建好并立刻加入 TabPages，全程常驻。
            // 注意：TabPages.Insert 在控件句柄还没创建时会被静默丢弃 ——
            // 页面会留在 Controls 里、却不在 TabPages 中，标签页就"消失"了（踩过，别再改回去）。
            _chartBox = new PictureBox();
            _chartBox.Dock = DockStyle.Fill;
            _chartBox.SizeMode = PictureBoxSizeMode.Zoom;
            _chartBox.BackColor = Color.FromArgb(244, 246, 249);
            _chartBox.Visible = false;   // 还没图时让提示文字露出来
            _chartBox.Cursor = Cursors.Hand;
            _chartBox.DoubleClick += OpenChartFile;

            _chartHint = new Label();
            _chartHint.Dock = DockStyle.Fill;
            _chartHint.TextAlign = ContentAlignment.MiddleCenter;
            _chartHint.ForeColor = Color.FromArgb(125, 133, 145);
            _chartHint.Text = "处理完成后，当天的报表图会显示在这里。";

            var imgHost = new Panel();
            imgHost.Dock = DockStyle.Fill;
            imgHost.BackColor = Color.FromArgb(244, 246, 249);
            imgHost.Controls.Add(_chartHint);
            imgHost.Controls.Add(_chartBox);

            // 一页纸整张太长，窗口里看不清，所以默认显示单张，上面一排按钮切换
            _chartBar = new FlowLayoutPanel();
            _chartBar.Dock = DockStyle.Fill;
            _chartBar.FlowDirection = FlowDirection.LeftToRight;
            _chartBar.WrapContents = false;
            _chartBar.AutoScroll = true;
            _chartBar.BackColor = Color.White;
            _chartBar.Visible = false;

            var chartLayout = new TableLayoutPanel();
            chartLayout.Dock = DockStyle.Fill;
            chartLayout.ColumnCount = 1;
            chartLayout.RowCount = 2;
            chartLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            chartLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
            chartLayout.Controls.Add(imgHost, 0, 0);
            chartLayout.Controls.Add(_chartBar, 0, 1);

            _chartPage = new TabPage("图表");
            _chartPage.Padding = new Padding(4);
            _chartPage.BackColor = Color.FromArgb(244, 246, 249);
            _chartPage.Controls.Add(chartLayout);

            _tabs.TabPages.Add(_chartPage);

            var strip = new StatusStrip();
            strip.BackColor = Color.White;
            _status = new ToolStripStatusLabel("就绪：把文件拖进上面的框，会自动开始。");
            _status.Spring = true;
            _status.TextAlign = ContentAlignment.MiddleLeft;
            strip.Items.Add(_status);

            _btnPick = new ToolStripButton("选择文件…");
            _btnPick.Click += PickFiles;
            strip.Items.Add(_btnPick);

            _btnScan = new ToolStripButton("扫描企业微信");
            _btnScan.Click += ScanWeChat;
            strip.Items.Add(_btnScan);

            _btnOpen = new ToolStripButton("打开输出文件夹");
            _btnOpen.Enabled = false;
            _btnOpen.Click += OpenOutput;
            strip.Items.Add(_btnOpen);

            _btnClear = new ToolStripButton("清空");
            _btnClear.Click += ClearAll;
            strip.Items.Add(_btnClear);

            // 停靠顺序：先 Fill，再 Bottom，最后 Top（与设计器生成的顺序一致）。
            Controls.Add(split);
            Controls.Add(strip);
            Controls.Add(_drop);
        }

        void DropPaint(object sender, PaintEventArgs e)
        {
            var p = (Panel)sender;
            using (var pen = new Pen(Color.FromArgb(180, 190, 205), 2))
            {
                pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                var r = new Rectangle(1, 1, p.Width - 3, p.Height - 3);
                e.Graphics.DrawRectangle(pen, r);
            }
        }

        void DropDragEnter(object sender, DragEventArgs e)
        {
            if (!_busy && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
                _drop.BackColor = Color.FromArgb(232, 242, 255);
            }
        }

        void DropDragLeave(object sender, EventArgs e)
        {
            _drop.BackColor = Color.White;
        }

        void DropDragDrop(object sender, DragEventArgs e)
        {
            _drop.BackColor = Color.White;
            if (_busy) return;
            var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (paths == null) return;
            AddAndRun(paths);
        }

        void PickFiles(object sender, EventArgs e)
        {
            if (_busy) return;
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "选择要生成的报表（可多选）";
                dlg.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                dlg.Multiselect = true;
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    AddAndRun(dlg.FileNames);
            }
        }

        void ScanWeChat(object sender, EventArgs e)
        {
            if (_busy) return;
            var folders = Inbox.FindWeChatFolders();
            if (folders.Count == 0)
            {
                MessageBox.Show(this,
                    "没有找到企业微信的接收目录。\r\n\r\n" +
                    "请确认企业微信已登录、并且至少收到过一个文件。\r\n" +
                    "也可以直接点「选择文件…」手动选。",
                    "没找到", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 当月 + 上月一起扫：月初那几小时，最新一份可能还落在上个月的目录里
            var all = Inbox.ListAll(folders);
            if (all.Count == 0)
            {
                MessageBox.Show(this, "这些目录里没有能识别的报表：\r\n" + string.Join("\r\n", folders.ToArray()),
                    "没有报表", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _queue.Clear();
            ClearTables();
            ClearChart();
            _log.Clear();
            foreach (var fd in folders) Append("扫描目录：" + fd);
            Append("共找到 " + all.Count + " 份可识别的报表，正在读取每份的通报时间…");
            // 文件名序号和落盘时间都不可信，只能读内容里的通报时间
            Inbox.ProbeTimes(all, Append);

            var latest = Inbox.LatestPerKind(all);
            Append("");
            Append("取每一类里最新的一份：");
            foreach (var f in latest)
            {
                string when = f.ReportTime.HasValue
                    ? f.ReportTime.Value.ToString("MM-dd HH:mm")
                    : "时间读不到，按修改时间 " + f.Modified.ToString("MM-dd HH:mm");
                string note = "";
                // 企业微信可能正在写这个文件，等它落盘
                if (!Inbox.WaitUntilReady(f.Path, 3000)) note = "  [可能还在下载]";
                Append("  " + KindNames.Of(f.Kind) + "：" + Path.GetFileName(f.Path) +
                       "（通报 " + when + "，" + (f.Size / 1024) + " KB）" + note);
                _queue.Add(f.Path);
            }
            Run();
        }

        void AddAndRun(string[] paths)
        {
            var added = new List<string>();
            foreach (var p in paths)
            {
                foreach (var f in Expand(p))
                {
                    if (_queue.Contains(f)) continue;
                    _queue.Add(f);
                    added.Add(f);
                }
            }
            if (added.Count == 0)
            {
                if (_queue.Count == 0)
                {
                    MessageBox.Show(this,
                        "这个文件夹里没有找到可识别的报表。\r\n\r\n" +
                        "文件名需要以「区县正向销售完成情况日预测通报」「网格正向销售完成情况日预测通报」" +
                        "「重点业务日调度通报」开头。\r\n" +
                        "也可以点「一键取企业微信里的报表并生成」，由程序自动挑最新的一份。",
                        "没有找到报表", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                Append("这些文件已经在队列里了，直接开始处理。");
                Run();
                return;
            }
            foreach (var f in added)
                Append("已加入：" + Path.GetFileName(f) + "  [" + KindNames.Of(Pipeline.Peek(f)) + "]");
            Run();
        }

        static IEnumerable<string> Expand(string path)
        {
            if (Directory.Exists(path))
            {
                var list = new List<string>();
                try
                {
                    foreach (var f in Directory.GetFiles(path, "*.xlsx", SearchOption.TopDirectoryOnly))
                        if (!Path.GetFileName(f).StartsWith("~$")) list.Add(f);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("读取文件夹失败：" + ex.Message);
                }
                list.Sort(StringComparer.OrdinalIgnoreCase);

                // 拖进来的多半是整个企业微信目录，里面有上百个无关文件、还有 14 MB 的大表。
                // 只挑出三份目标报表 —— 多读一个无关大表就可能卡几分钟。
                var found = Inbox.ListAll(new[] { path });
                if (found.Count == 0) return new string[0];

                // 目录里同一类往往存着十几份历史文件，每类只取最新的一份，
                // 否则会把历史报表也当成"重复拖入"逐份报错。
                Inbox.ProbeTimes(found, null);
                var latest = Inbox.LatestPerKind(found);
                var picked = new List<string>();
                foreach (var f in latest) picked.Add(f.Path);
                return picked;
            }
            if (File.Exists(path) && path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                && !Path.GetFileName(path).StartsWith("~$"))
                return new[] { path };
            return new string[0];
        }

        void Run()
        {
            if (_busy || _queue.Count == 0) return;
            _busy = true;
            _btnPick.Enabled = false;
            _drop.AllowDrop = false;
            _dropText.AllowDrop = false;
            _dropText.Text = "正在生成，请稍候…";
            ClearTables();
            ClearChart();
            SetStatus("正在处理 " + _queue.Count + " 份文件…");

            var inputs = new List<string>(_queue);
            var worker = new Thread(delegate()
            {
                BatchResult result = null;
                string fatal = null;
                try
                {
                    result = Pipeline.Run(inputs, null, delegate(string line) { Post(delegate { Append(line); }); });
                }
                catch (Exception ex)
                {
                    fatal = ex.Message;
                }

                Post(delegate
                {
                    _busy = false;
                    _btnPick.Enabled = true;
                    _drop.AllowDrop = true;
                    _dropText.AllowDrop = true;
                    _dropText.Text = "把 xlsx 文件（或放文件的文件夹）拖到这里\r\n区县 / 网格 / 重点业务，一次拖三个即可";
                    if (fatal != null)
                    {
                        Append("错误：" + fatal);
                        SetStatus("失败：" + fatal);
                        return;
                    }
                    _queue.Clear();
                    ShowResult(result);
                });
            });
            worker.IsBackground = true;
            worker.Start();
        }

        internal void ShowResult(BatchResult r)
        {
            if (r == null) { SetStatus("没有结果。"); return; }
            _lastOutput = r.OutputDir;
            _btnOpen.Enabled = Directory.Exists(r.OutputDir);

            int ok = 0, bad = 0;
            foreach (var f in r.Files)
            {
                if (f.Error == null) ok++; else bad++;
            }

            // 图表放第一个标签页：先看图，再看数
            bool hasChart = ShowCharts(r);

            foreach (var t in r.Tables) AddTable(t);

            // 全部加完再选中，否则后面 AddTable 会把选中态抢走
            if (hasChart) _tabs.SelectedTab = _chartPage;

            if (bad == 0)
            {
                SetStatus("完成：" + ok + " 份，输出在 " + r.OutputDir);
                Append("");
                Append("全部完成。文件在：" + r.OutputDir);
            }
            else
            {
                SetStatus("完成 " + ok + " 份，失败 " + bad + " 份（详见左侧日志）。");
            }
        }

        /// <summary>
        /// 按「一页纸 + 每张单图」生成切换按钮，默认显示第一张。
        /// 返回是否至少显示了一张图。
        /// </summary>
        bool ShowCharts(BatchResult r)
        {
            _chartBar.Controls.Clear();
            var labels = new List<string>();
            var paths = new List<string>();

            bool hasCombined = false;
            if (!string.IsNullOrEmpty(r.CombinedChartPath) && File.Exists(r.CombinedChartPath))
            {
                labels.Add("一页纸（全部）");
                paths.Add(r.CombinedChartPath);
                hasCombined = true;
            }
            if (r.ChartFiles != null)
            {
                foreach (var f in r.ChartFiles)
                {
                    if (!File.Exists(f)) continue;
                    labels.Add(ShortChartName(f));
                    paths.Add(f);
                }
            }

            if (paths.Count == 0)
            {
                _chartBar.Visible = false;
                _chartBox.Visible = false;
                _chartHint.Text = "这次没有生成图表（详见左侧日志）。";
                _chartHint.Visible = true;
                return false;
            }

            for (int i = 0; i < paths.Count; i++)
            {
                var btn = new Button();
                btn.Text = labels[i];
                btn.AutoSize = true;
                btn.Height = 26;
                btn.Margin = new Padding(2, 5, 2, 1);
                btn.FlatStyle = FlatStyle.System;
                string path = paths[i];
                btn.Click += delegate { ShowChart(path); };
                _chartBar.Controls.Add(btn);
            }

            // 只有一张图（比如只拖了区县表）就不必显示切换栏
            _chartBar.Visible = paths.Count > 1;
            // 默认选第一张单图（比一页纸看得清）；只有一页纸时才显示整页
            int start = hasCombined && paths.Count > 1 ? 1 : 0;
            return ShowChart(paths[start]);
        }

        /// <summary>「图3_网格收入完成率.png」→「网格收入完成率」。</summary>
        static string ShortChartName(string file)
        {
            string name = Path.GetFileNameWithoutExtension(file);
            int us = name.IndexOf('_');
            if (us > 0 && us < name.Length - 1) name = name.Substring(us + 1);
            return name;
        }

        /// <summary>
        /// 把 PNG 显示到「图表」页。必须先读进内存再解码：
        /// Image.FromStream 会一直引用那个流，直接读文件会锁住 PNG 不放。
        /// </summary>
        bool ShowChart(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return false;
            try
            {
                var old = _chartBox.Image;
                var bytes = File.ReadAllBytes(path);
                using (var ms = new MemoryStream(bytes))
                using (var img = Image.FromStream(ms))
                {
                    _chartBox.Image = new Bitmap(img);   // 复制一份，脱离流
                }
                if (old != null) old.Dispose();

                if (!_tabs.TabPages.Contains(_chartPage))
                    _tabs.TabPages.Add(_chartPage);
                _chartBox.Visible = true;
                _chartHint.Visible = false;
                _chartPath = path;
                return true;
            }
            catch (Exception ex)
            {
                Append("图表加载失败：" + ex.Message);
                return false;
            }
        }
        /// <summary>把图表页恢复成"还没有图"的状态，并释放旧图占用的内存。</summary>
        void ClearChart()
        {
            var img = _chartBox.Image;
            _chartBox.Image = null;
            if (img != null) img.Dispose();
            _chartPath = null;
            _chartBar.Controls.Clear();
            _chartBar.Visible = false;
            _chartBox.Visible = false;
            _chartHint.Text = "处理完成后，当天的报表图会显示在这里。";
            _chartHint.Visible = true;
        }

        /// <summary>双击图片 → 用系统看图程序打开原尺寸 PNG。</summary>
        void OpenChartFile(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_chartPath) || !File.Exists(_chartPath)) return;
            try { Process.Start("explorer.exe", "\"" + _chartPath + "\""); }
            catch (Exception ex) { Append("打开图片失败：" + ex.Message); }
        }

        /// <summary>
        /// 清掉上一轮的数据表页，保留常驻的「图表」页。
        /// 不能用 TabPages.Clear()：图表页一离开 TabControl，就要靠重新 Add 找回来。
        /// </summary>
        void ClearTables()
        {
            for (int i = _tabs.TabPages.Count - 1; i >= 0; i--)
            {
                var p = _tabs.TabPages[i];
                if (p == _chartPage) continue;
                _tabs.TabPages.RemoveAt(i);
                p.Dispose();
            }
        }

        void AddTable(PreviewTable t)
        {
            var page = new TabPage(t.Title.Length > 18 ? t.Title.Substring(0, 18) + "…" : t.Title);
            page.Padding = new Padding(4);

            var grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.RowHeadersVisible = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = false;
            grid.BackgroundColor = Color.White;
            grid.BorderStyle = BorderStyle.None;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font(Font, FontStyle.Bold);
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 244, 250);
            grid.EnableHeadersVisualStyles = false;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 250, 252);

            foreach (var c in t.Columns)
            {
                int i = grid.Columns.Add(c, c);
                grid.Columns[i].SortMode = DataGridViewColumnSortMode.Automatic;
                if (i > 0) grid.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            foreach (var row in t.Rows) grid.Rows.Add(row);

            page.Controls.Add(grid);
            _tabs.TabPages.Add(page);
        }

        void OpenOutput(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_lastOutput) || !Directory.Exists(_lastOutput)) return;
            Process.Start("explorer.exe", "\"" + _lastOutput + "\"");
        }

        void ClearAll(object sender, EventArgs e)
        {
            if (_busy) return;
            _queue.Clear();
            ClearTables();
            ClearChart();
            _log.Clear();
            SetStatus("就绪：把文件拖进上面的框，会自动开始。");
        }

        void Append(string line)
        {
            _log.AppendText(line + Environment.NewLine);
        }

        void SetStatus(string text)
        {
            _status.Text = text;
        }

        void Post(MethodInvoker action)
        {
            if (IsDisposed) return;
            try { BeginInvoke(action); }
            catch (Exception) { }
        }
    }
}