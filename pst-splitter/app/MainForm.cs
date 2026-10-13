using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace PstSplitter
{
    public class MainForm : Form
    {
        static readonly Color Accent = Color.FromArgb(0, 84, 147);

        class ModeInfo
        {
            public SplitMode Mode;
            public string Text;
            public string Help;
            public RadioButton Radio;
            public Panel Panel;
        }

        readonly ModeInfo[] modes =
        {
            new ModeInfo { Mode = SplitMode.Size, Text = "Size", Help = "Starts a new file each time the current one reaches the size below. Files are named _part001, _part002, and so on." },
            new ModeInfo { Mode = SplitMode.Year, Text = "Year", Help = "One file per year, using the date each item was received or sent." },
            new ModeInfo { Mode = SplitMode.Month, Text = "Month", Help = "One file per month, using the date each item was received or sent." },
            new ModeInfo { Mode = SplitMode.Date, Text = "Date range", Help = "One file with the items between these dates (both days included)." },
            new ModeInfo { Mode = SplitMode.Folder, Text = "Folder", Help = "One file per main folder: Inbox, Sent Items, Calendar and so on, with their subfolders. If everything is inside one folder, it splits on that folder's subfolders." },
            new ModeInfo { Mode = SplitMode.Sender, Text = "Email address", Help = "One file per email address. List the addresses to pull out, or leave the box empty for one file per sender." },
        };

        TextBox txtSource, txtOutput, txtEmail, txtLog;
        Label lblSourceInfo, lblStatus, lblElapsed;
        readonly System.Windows.Forms.Timer ticker = new System.Windows.Forms.Timer { Interval = 1000 };
        readonly Stopwatch runClock = new Stopwatch();
        GroupBox gbFiles, gbSplit;
        DateTimePicker dtStart, dtEnd;
        CheckBox chkOtherDate, chkOtherSender, chkSize, chkSkip;
        ComboBox cboMatch;
        NumericUpDown numSize;
        Button btnPreview, btnSplit, btnStop, btnOpen;
        ProgressBar bar;
        TabControl tabs;
        TabPage tabResults, tabLog;
        ListView list;

        volatile bool cancel;
        bool running, closeAfter;
        string lastOutput;

        public MainForm()
        {
            Text = "PST Splitter";
            Font = new Font("Segoe UI", 9F);
            ClientSize = new Size(720, 740);
            MinimumSize = new Size(736, 700);
            StartPosition = FormStartPosition.CenterScreen;
            AllowDrop = true;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            BuildUi();
            UpdateModeView();
        }

        // ---- layout ---------------------------------------------------------

        static T Add<T>(Control parent, int x, int y, int w, int h, string text = null) where T : Control, new()
        {
            var c = new T { Location = new Point(x, y) };
            if (w > 0) c.Width = w;
            if (h > 0) c.Height = h;
            if (text != null) c.Text = text;
            parent.Controls.Add(c);
            return c;
        }

        static Label AddLabel(Control parent, string text, int x, int y, bool hint = false)
        {
            var l = Add<Label>(parent, x, y + 3, 0, 0, text);
            l.AutoSize = true;
            if (hint) l.ForeColor = Color.DimGray;
            return l;
        }

        void BuildUi()
        {
            var header = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = Accent };
            Controls.Add(header);
            var title = AddLabel(header, "PST Splitter", 16, 6);
            title.Font = new Font("Segoe UI Semibold", 16F);
            title.ForeColor = Color.White;
            var sub = AddLabel(header, "Split a large Outlook data file into smaller ones. The original is not changed.", 18, 38);
            sub.ForeColor = Color.FromArgb(210, 228, 245);

            // 1. Files
            gbFiles = Add<GroupBox>(this, 12, 74, 696, 104, "1. Files");
            gbFiles.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            AddLabel(gbFiles, "PST file to split", 12, 24);
            txtSource = Add<TextBox>(gbFiles, 130, 24, 456, 0);
            txtSource.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtSource.Leave += (s, e) => ShowSourceInfo();
            var btnSource = Add<Button>(gbFiles, 594, 23, 90, 26, "Browse...");
            btnSource.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnSource.Click += (s, e) => BrowseSource();
            lblSourceInfo = AddLabel(gbFiles, "Or drag a .pst file onto this window.", 130, 46, true);
            AddLabel(gbFiles, "Save split files to", 12, 70);
            txtOutput = Add<TextBox>(gbFiles, 130, 70, 456, 0);
            txtOutput.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            var btnOutput = Add<Button>(gbFiles, 594, 69, 90, 26, "Browse...");
            btnOutput.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnOutput.Click += (s, e) => BrowseOutput();

            // 2. Split by
            gbSplit = Add<GroupBox>(this, 12, 186, 696, 240, "2. Split by");
            gbSplit.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            for (int i = 0; i < modes.Length; i++)
            {
                var m = modes[i];
                m.Radio = Add<RadioButton>(gbSplit, 16, 26 + 27 * i, 150, 24, m.Text);
                m.Radio.CheckedChanged += (s, e) => UpdateModeView();
                m.Panel = Add<Panel>(gbSplit, 180, 20, 504, 150);
                m.Panel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                m.Panel.Visible = false;
                var help = AddLabel(m.Panel, m.Help, 0, 0, true);
                help.AutoSize = false;
                help.Size = new Size(500, 40);
            }
            modes[0].Radio.Checked = true;
            var sep = Add<Label>(gbSplit, 172, 22, 2, 150, "");
            sep.BorderStyle = BorderStyle.Fixed3D;

            var pDate = ModePanel(SplitMode.Date);
            AddLabel(pDate, "From", 0, 48);
            dtStart = Add<DateTimePicker>(pDate, 44, 48, 120, 0);
            dtStart.Format = DateTimePickerFormat.Short;
            dtStart.Value = DateTime.Today.AddYears(-1);
            AddLabel(pDate, "to", 174, 48);
            dtEnd = Add<DateTimePicker>(pDate, 198, 48, 120, 0);
            dtEnd.Format = DateTimePickerFormat.Short;
            chkOtherDate = Add<CheckBox>(pDate, 0, 84, 400, 24, "Put items outside the dates into an \"_Other\" file");

            var pSender = ModePanel(SplitMode.Sender);
            AddLabel(pSender, "Addresses", 0, 44);
            txtEmail = Add<TextBox>(pSender, 70, 44, 430, 0);
            txtEmail.Multiline = true;
            txtEmail.AcceptsReturn = true;
            txtEmail.ScrollBars = ScrollBars.Vertical;
            txtEmail.Height = 46;
            txtEmail.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            AddLabel(pSender, "One per line, or separated by commas.", 70, 90, true);
            AddLabel(pSender, "Match on", 0, 116);
            cboMatch = Add<ComboBox>(pSender, 70, 116, 150, 0);
            cboMatch.DropDownStyle = ComboBoxStyle.DropDownList;
            cboMatch.Items.AddRange(new object[] { "Sender", "Recipients", "Both" });
            cboMatch.SelectedIndex = 0;
            chkOtherSender = Add<CheckBox>(pSender, 236, 116, 260, 24, "Unmatched items go to an \"_Other\" file");

            var line = Add<Label>(gbSplit, 12, 180, 672, 2, "");
            line.BorderStyle = BorderStyle.Fixed3D;
            line.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            chkSize = Add<CheckBox>(gbSplit, 16, 190, 190, 24, "Limit each file to");
            chkSize.CheckedChanged += (s, e) => numSize.Enabled = chkSize.Checked;
            numSize = Add<NumericUpDown>(gbSplit, 206, 190, 90, 0);
            numSize.Minimum = 10;
            numSize.Maximum = 50000;
            numSize.Increment = 256;
            numSize.Value = 2048;
            numSize.ThousandsSeparator = true;
            AddLabel(gbSplit, "MB   (set about 10% under the size you need)", 300, 190, true);
            chkSkip = Add<CheckBox>(gbSplit, 16, 212, 460, 24, "Skip Deleted Items, Junk Email and Outlook sync folders");
            chkSkip.Checked = true;

            // 3. Run
            btnPreview = Add<Button>(this, 12, 438, 110, 32, "Preview");
            btnPreview.Click += (s, e) => StartSplit(true);
            btnSplit = Add<Button>(this, 130, 438, 110, 32, "Split");
            btnSplit.BackColor = Accent;
            btnSplit.ForeColor = Color.White;
            btnSplit.FlatStyle = FlatStyle.Flat;
            btnSplit.FlatAppearance.BorderSize = 0;
            btnSplit.Font = new Font("Segoe UI Semibold", 9F);
            btnSplit.Click += (s, e) => StartSplit(false);
            btnStop = Add<Button>(this, 248, 438, 90, 32, "Stop");
            btnStop.Enabled = false;
            btnStop.Click += (s, e) =>
            {
                cancel = true;
                btnStop.Enabled = false;
                lblStatus.Text = "Stopping after the current item...";
            };
            btnOpen = Add<Button>(this, 568, 438, 140, 32, "Open output folder");
            btnOpen.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnOpen.Enabled = false;
            btnOpen.Click += (s, e) => { if (lastOutput != null) Process.Start("explorer.exe", "\"" + lastOutput + "\""); };
            var tip = new ToolTip();
            tip.SetToolTip(btnPreview, "Shows the files that would be made, without writing anything");
            tip.SetToolTip(btnSplit, "Writes the new PST files");

            bar = Add<ProgressBar>(this, 12, 480, 696, 18);
            bar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblStatus = AddLabel(this, "Choose a PST file to start.", 12, 500);
            lblStatus.AutoSize = false;
            lblStatus.Size = new Size(566, 20);
            lblStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblStatus.AutoEllipsis = true;
            // A clock that ticks while a split runs, so it's clear the program hasn't frozen.
            lblElapsed = AddLabel(this, "", 588, 500, true);
            lblElapsed.AutoSize = false;
            lblElapsed.Size = new Size(120, 20);
            lblElapsed.TextAlign = ContentAlignment.TopRight;
            lblElapsed.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            ticker.Tick += (s, e) => lblElapsed.Text = "Elapsed " + FormatTime(runClock.Elapsed);

            tabs = Add<TabControl>(this, 12, 526, 696, 202);
            tabs.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            tabResults = new TabPage("Results");
            tabLog = new TabPage("Log");
            tabs.TabPages.AddRange(new[] { tabResults, tabLog });

            list = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true };
            list.Columns.Add("File", 420);
            list.Columns.Add("Items", 100, HorizontalAlignment.Right);
            list.Columns.Add("Size (MB)", 120, HorizontalAlignment.Right);
            list.DoubleClick += (s, e) =>
            {
                if (list.SelectedItems.Count == 0) return;
                var path = list.SelectedItems[0].Tag as string;
                if (path != null && File.Exists(path)) Process.Start("explorer.exe", "/select,\"" + path + "\"");
            };
            tabResults.Controls.Add(list);

            txtLog = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, Font = new Font("Consolas", 9F) };
            tabLog.Controls.Add(txtLog);

            AcceptButton = btnPreview;
        }

        Panel ModePanel(SplitMode mode) => modes.First(m => m.Mode == mode).Panel;

        SplitMode SelectedMode => modes.First(m => m.Radio.Checked).Mode;

        // ---- behaviour ------------------------------------------------------

        void UpdateModeView()
        {
            if (chkSize == null) return;
            var mode = SelectedMode;
            foreach (var m in modes) m.Panel.Visible = m.Mode == mode;
            if (mode == SplitMode.Size)
            {
                chkSize.Checked = true;
                chkSize.Enabled = false;
            }
            else
            {
                chkSize.Enabled = true;
            }
            numSize.Enabled = chkSize.Checked;
        }

        void BrowseSource()
        {
            using (var d = new OpenFileDialog { Filter = "Outlook data files (*.pst)|*.pst|All files (*.*)|*.*" })
                if (d.ShowDialog(this) == DialogResult.OK) SetSource(d.FileName);
        }

        void BrowseOutput()
        {
            using (var d = new FolderBrowserDialog { Description = "Choose where to save the split PST files" })
            {
                if (Directory.Exists(txtOutput.Text)) d.SelectedPath = txtOutput.Text;
                if (d.ShowDialog(this) == DialogResult.OK) txtOutput.Text = d.SelectedPath;
            }
        }

        void SetSource(string path)
        {
            txtSource.Text = path;
            if (txtOutput.Text.Trim().Length == 0 && path.Length > 0)
            {
                try
                {
                    txtOutput.Text = Path.Combine(Path.GetDirectoryName(path), Path.GetFileNameWithoutExtension(path) + "_split");
                }
                catch { }
            }
            ShowSourceInfo();
        }

        void ShowSourceInfo()
        {
            var path = txtSource.Text.Trim();
            bool exists = false;
            try { exists = path.Length > 0 && File.Exists(path); } catch { }
            if (exists)
            {
                lblSourceInfo.Text = string.Format("{0:N0} MB. Click Preview to see how it would be split.", new FileInfo(path).Length / 1048576.0);
                lblStatus.Text = "Ready.";
            }
            else
            {
                lblSourceInfo.Text = path.Length > 0 ? "File not found." : "Or drag a .pst file onto this window.";
            }
        }

        protected override void OnDragEnter(DragEventArgs e)
        {
            base.OnDragEnter(e);
            if (!running && e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        protected override void OnDragDrop(DragEventArgs e)
        {
            base.OnDragDrop(e);
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0) SetSource(files[0]);
        }

        SplitOptions ReadOptions()
        {
            var mode = SelectedMode;
            var opts = new SplitOptions
            {
                Source = txtSource.Text.Trim(),
                Output = txtOutput.Text.Trim(),
                Mode = mode,
                MaxSizeMB = chkSize.Checked ? (int)numSize.Value : 0,
                StartDate = dtStart.Value.Date,
                EndDate = dtEnd.Value.Date,
                EmailAddresses = SplitOptions.ParseAddresses(txtEmail.Text),
                MatchOn = (MatchOn)cboMatch.SelectedIndex,
                IncludeOther = (mode == SplitMode.Date && chkOtherDate.Checked) || (mode == SplitMode.Sender && chkOtherSender.Checked),
            };
            if (chkSkip.Checked) opts.ExcludeFolders.AddRange(SplitOptions.DefaultExcludes);
            return opts;
        }

        void StartSplit(bool preview)
        {
            var opts = ReadOptions();
            var problem = opts.Validate();
            if (problem != null)
            {
                MessageBox.Show(this, problem, "PST Splitter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!preview)
            {
                if (Directory.Exists(opts.Output) && Directory.EnumerateFiles(opts.Output, "*.pst").Any())
                {
                    MessageBox.Show(this, "The output folder already has PST files in it.\n\nChoose an empty folder.", "PST Splitter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var msg = "Split " + Path.GetFileName(opts.Source) + " into " + opts.Output + "?\n\n" +
                          "The original PST is only read, never changed. If Outlook has it open, close Outlook first.";
                if (MessageBox.Show(this, msg, "PST Splitter", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
            }

            txtLog.Clear();
            list.Items.Clear();
            bar.Value = 0;
            cancel = false;
            SetRunning(true);

            var splitter = new Splitter
            {
                Log = m => BeginInvoke((Action)(() => WriteLog(m))),
                ShouldCancel = () => cancel,
            };
            var clock = Stopwatch.StartNew();
            splitter.Progress = (done, total, folder) =>
            {
                // Update the window at most ~10 times a second while counting.
                if (total > 0 && done != total && clock.ElapsedMilliseconds < 100) return;
                clock.Restart();
                BeginInvoke((Action)(() => ShowProgress(done, total, folder)));
            };

            // Split off the UI thread so the window stays responsive.
            var worker = new Thread(() =>
            {
                SplitResult result = null;
                Exception error = null;
                try { result = splitter.Run(opts, preview); }
                catch (Exception ex) { error = ex; }
                BeginInvoke((Action)(() => Finish(result, error)));
            });
            worker.IsBackground = true;
            worker.Start();
        }

        void ShowProgress(int done, int total, string label)
        {
            if (total <= 0)
            {
                // Total not known yet: show a moving bar.
                if (bar.Style != ProgressBarStyle.Marquee) { bar.Style = ProgressBarStyle.Marquee; bar.MarqueeAnimationSpeed = 30; }
            }
            else
            {
                if (bar.Style != ProgressBarStyle.Continuous) bar.Style = ProgressBarStyle.Continuous;
                bar.Maximum = total;
                bar.Value = Math.Min(done, total);
            }
            lblStatus.Text = label;
        }

        static string FormatTime(TimeSpan t) =>
            t.TotalHours >= 1 ? string.Format("{0}:{1:00}:{2:00}", (int)t.TotalHours, t.Minutes, t.Seconds) : string.Format("{0}:{1:00}", t.Minutes, t.Seconds);

        void WriteLog(string msg)
        {
            txtLog.AppendText(msg + "\r\n");
            if (msg.Length > 0 && !msg.StartsWith("  ") && !msg.EndsWith(" MB")) lblStatus.Text = msg;
        }

        void Finish(SplitResult result, Exception error)
        {
            SetRunning(false);
            UpdateModeView();
            if (error != null)
            {
                var message = error is System.Reflection.TargetInvocationException && error.InnerException != null ? error.InnerException.Message : error.Message;
                WriteLog("ERROR: " + message);
                lblStatus.Text = "Something went wrong. See the Log tab.";
                tabs.SelectedTab = tabLog;
                if (!closeAfter) MessageBox.Show(this, message, "PST Splitter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                foreach (var f in result.Files)
                {
                    var row = new ListViewItem(f.Name) { Tag = f.Path };
                    row.SubItems.Add(f.Items.ToString("N0"));
                    row.SubItems.Add((f.Bytes / 1048576.0).ToString("N1"));
                    list.Items.Add(row);
                }
                tabs.SelectedTab = tabResults;
                var summary = string.Format("{0:N0} items in {1} file(s)", result.Copied, result.Files.Count);
                if (result.Skipped > 0) summary += string.Format(", {0:N0} skipped", result.Skipped);
                if (result.Failed > 0) summary += string.Format(", {0:N0} could not be copied (see Log)", result.Failed);
                if (result.Cancelled) lblStatus.Text = "Stopped. " + summary + ".";
                else if (result.Preview) lblStatus.Text = "Preview: " + summary + ". Nothing was written.";
                else
                {
                    bar.Value = bar.Maximum;
                    lblStatus.Text = "Finished: " + summary + ".";
                }
                if (!result.Preview && result.Files.Count > 0)
                {
                    lastOutput = result.Output;
                    btnOpen.Enabled = true;
                }
            }
            if (closeAfter) Close();
        }

        void SetRunning(bool value)
        {
            running = value;
            if (value)
            {
                runClock.Restart();
                lblElapsed.Text = "Elapsed 0:00";
                ticker.Start();
            }
            else
            {
                runClock.Stop();
                ticker.Stop();
                lblElapsed.Text = "Took " + FormatTime(runClock.Elapsed);
                bar.Style = ProgressBarStyle.Continuous;
            }
            gbFiles.Enabled = gbSplit.Enabled = btnPreview.Enabled = btnSplit.Enabled = !value;
            btnStop.Enabled = value;
            if (value) btnOpen.Enabled = false;
            UseWaitCursor = value;
            btnStop.UseWaitCursor = false;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (running)
            {
                var answer = MessageBox.Show(this, "A split is running. Stop it and close?", "PST Splitter", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (answer == DialogResult.Yes)
                {
                    cancel = true;
                    closeAfter = true;
                    btnStop.Enabled = false;
                    lblStatus.Text = "Stopping after the current item...";
                }
                e.Cancel = true;
            }
            base.OnFormClosing(e);
        }
    }
}
