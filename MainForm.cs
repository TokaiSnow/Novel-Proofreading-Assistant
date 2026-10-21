using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace NovelProofreader
{
    public class MainForm : Form
    {
        private const string EmptySteps = "① 选择 Word 文档\r\n② 自动扫描并打开批注版\r\n③ 修改后点“写回源文件”";
        private const string EmptyNoIssues = "未发现明显问题\r\n临时文件已生成，可在 Word 中查看";
        private const string EmptyApplied = "选中的修改已全部应用\r\n可点【写回源文件】完成";

        private readonly Font uiFont;
        private readonly Font boldFont;

        private TextBox txtSource;
        private Button btnSelect;
        private Button btnRescan;
        private Button btnOpen;
        private Button btnApply;
        private Button btnWrite;
        private CheckBox chkPunctuation;
        private CheckBox chkAdvanced;
        private CheckBox chkBackup;
        private CheckBox chkOnline;
        private CheckBox chkSelectAll;
        private LinkLabel linkApi;
        private ListView listIssues;
        private Label lblEmpty;
        private Label lblSummary;
        private Label lblStatus;
        private ProgressBar progress;
        private ColumnHeader colNote;

        private string sourcePath;
        private string tempPath;
        private string initialFile;
        private DateTime tempCreatedUtc;
        private DateTime tempOpenedUtc;
        private DateTime sourceWriteUtc;
        private long sourceLength;
        private ScanResult scanResult;
        private DeepSeekConfig deepSeekConfig;
        private bool busy;
        private bool updatingSelectAll;
        private bool waitingForWordOpen;
        private DateTime wordOpenWaitStart;
        private System.Windows.Forms.Timer wordOpenTimer;
        private BackgroundWorker worker;

        public MainForm() : this(null)
        {
        }

        public MainForm(string initialFile)
        {
            this.initialFile = initialFile;
            deepSeekConfig = DeepSeekConfig.LoadOrCreate();
            uiFont = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
            boldFont = new Font(uiFont, FontStyle.Bold);
            InitializeUi();
            this.Load += delegate
            {
                UpdateButtons();
                if (deepSeekConfig.Enabled && !deepSeekConfig.HasKey)
                    lblStatus.Text = "DeepSeek 已启用但未填写 API Key，请点【API 配置】。";
                if (!string.IsNullOrEmpty(this.initialFile))
                {
                    string file = this.initialFile;
                    this.initialFile = null;
                    LoadSource(file);
                }
            };
            this.FormClosing += OnFormClosing;
        }

        private void InitializeUi()
        {
            this.SuspendLayout();
            this.Text = "小说校对助手";
            this.ClientSize = new Size(860, 630);
            this.MinimumSize = new Size(760, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = uiFont;
            this.BackColor = Color.FromArgb(246, 248, 250);
            this.KeyPreview = true;
            this.AllowDrop = true;
            this.AutoScaleMode = AutoScaleMode.Dpi;

            try
            {
                this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { }

            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.Padding = new Padding(12, 10, 12, 6);
            root.ColumnCount = 1;
            root.RowCount = 5;
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 84));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));

            root.Controls.Add(BuildTopPanel(), 0, 0);
            root.Controls.Add(BuildOptionsPanel(), 0, 1);
            root.Controls.Add(BuildListPanel(), 0, 2);
            root.Controls.Add(BuildBottomPanel(), 0, 3);
            root.Controls.Add(BuildStatusPanel(), 0, 4);
            this.Controls.Add(root);

            ToolTip tip = new ToolTip();
            tip.SetToolTip(txtSource, "当前要校对的源文件");
            tip.SetToolTip(btnSelect, "选择工作区中的 Word 文档（Ctrl+O）");
            tip.SetToolTip(btnRescan, "用当前选项重新扫描并生成新的临时文件（F5）");
            tip.SetToolTip(btnOpen, "再次打开已生成的批注版临时文件");
            tip.SetToolTip(btnApply, "把勾选的建议直接写入临时文件；不会改动源文件");
            tip.SetToolTip(btnWrite, "保存并关闭临时文件，将其正文写回源文件，然后删除临时文件（Ctrl+S）");
            tip.SetToolTip(chkBackup, "覆盖源文件前，在同目录保留一份“文件名.校对备份.docx”");
            tip.SetToolTip(chkOnline, "开启后把正文发送到 DeepSeek API 做在线校对；需要先在【API 配置】中填写 ApiKey");
            tip.SetToolTip(chkSelectAll, "勾选/取消勾选列表中的全部修改项");
            tip.SetToolTip(linkApi, "打开 小说校对助手.ini 填写 DeepSeek API Key");

            this.DragEnter += OnDragEnter;
            this.DragDrop += OnDragDrop;
            this.KeyDown += OnKeyDown;
            this.FormClosed += delegate
            {
                if (wordOpenTimer != null) wordOpenTimer.Stop();
            };

            wordOpenTimer = new System.Windows.Forms.Timer();
            wordOpenTimer.Interval = 250;
            wordOpenTimer.Tick += delegate { CheckWordOpenTimer(); };

            this.ResumeLayout(false);
        }

        private Control BuildTopPanel()
        {
            Panel panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.Margin = new Padding(0);

            TableLayoutPanel grid = new TableLayoutPanel();
            grid.Dock = DockStyle.Fill;
            grid.ColumnCount = 4;
            grid.RowCount = 2;
            grid.Margin = new Padding(0);
            grid.Padding = new Padding(0);
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 168));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

            Label label = new Label();
            label.Text = "源文件";
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.ForeColor = Color.FromArgb(52, 73, 94);
            grid.Controls.Add(label, 0, 0);

            txtSource = new TextBox();
            txtSource.Dock = DockStyle.Fill;
            txtSource.ReadOnly = true;
            txtSource.BorderStyle = BorderStyle.FixedSingle;
            txtSource.BackColor = Color.White;
            txtSource.Margin = new Padding(0, 6, 8, 6);
            txtSource.TabStop = false;
            grid.Controls.Add(txtSource, 1, 0);

            btnSelect = MakeButton("选择 Word 文件…", true);
            btnSelect.Dock = DockStyle.Fill;
            btnSelect.Margin = new Padding(0, 4, 8, 4);
            btnSelect.Click += delegate { SelectFile(); };
            grid.Controls.Add(btnSelect, 2, 0);

            btnRescan = MakeButton("重新扫描", false);
            btnRescan.Dock = DockStyle.Fill;
            btnRescan.Margin = new Padding(0, 4, 0, 4);
            btnRescan.Click += delegate { Rescan(); };
            grid.Controls.Add(btnRescan, 3, 0);

            lblSummary = new Label();
            lblSummary.Text = "尚未选择文件";
            lblSummary.Dock = DockStyle.Fill;
            lblSummary.TextAlign = ContentAlignment.MiddleLeft;
            lblSummary.AutoEllipsis = true;
            lblSummary.ForeColor = Color.FromArgb(86, 101, 115);
            grid.SetColumnSpan(lblSummary, 4);
            grid.Controls.Add(lblSummary, 0, 1);

            panel.Controls.Add(grid);
            return panel;
        }

        private Control BuildOptionsPanel()
        {
            FlowLayoutPanel panel = new FlowLayoutPanel();
            panel.Dock = DockStyle.Fill;
            panel.FlowDirection = FlowDirection.LeftToRight;
            panel.WrapContents = false;
            panel.Padding = new Padding(0, 4, 0, 0);
            panel.Margin = new Padding(0);

            chkPunctuation = new CheckBox();
            chkPunctuation.Text = "标点/格式检查";
            chkPunctuation.Checked = true;
            chkPunctuation.AutoSize = true;
            chkPunctuation.Margin = new Padding(0, 4, 12, 0);
            chkPunctuation.CheckedChanged += delegate { MarkOptionsDirty(); };
            panel.Controls.Add(chkPunctuation);

            chkAdvanced = new CheckBox();
            chkAdvanced.Text = "“的/地/得”疑似";
            chkAdvanced.Checked = true;
            chkAdvanced.AutoSize = true;
            chkAdvanced.Margin = new Padding(0, 4, 12, 0);
            chkAdvanced.CheckedChanged += delegate { MarkOptionsDirty(); };
            panel.Controls.Add(chkAdvanced);

            chkBackup = new CheckBox();
            chkBackup.Text = "写回前备份";
            chkBackup.Checked = true;
            chkBackup.AutoSize = true;
            chkBackup.Margin = new Padding(0, 4, 10, 0);
            panel.Controls.Add(chkBackup);

            chkOnline = new CheckBox();
            chkOnline.Text = "DeepSeek 在线";
            chkOnline.Checked = deepSeekConfig.Enabled && deepSeekConfig.HasKey;
            chkOnline.AutoSize = true;
            chkOnline.Margin = new Padding(0, 4, 4, 0);
            chkOnline.CheckedChanged += delegate { MarkOptionsDirty(); };
            panel.Controls.Add(chkOnline);

            linkApi = new LinkLabel();
            linkApi.Text = "API 配置";
            linkApi.AutoSize = true;
            linkApi.Margin = new Padding(0, 5, 6, 0);
            linkApi.LinkColor = Color.FromArgb(40, 116, 166);
            linkApi.LinkClicked += delegate { OpenApiConfig(); };
            panel.Controls.Add(linkApi);

            LinkLabel help = new LinkLabel();
            help.Text = "使用说明";
            help.AutoSize = true;
            help.Margin = new Padding(0, 5, 0, 0);
            help.LinkColor = Color.FromArgb(40, 116, 166);
            help.LinkClicked += delegate { ShowHelp(); };
            panel.Controls.Add(help);

            return panel;
        }

        private Control BuildListPanel()
        {
            Panel outer = new Panel();
            outer.Dock = DockStyle.Fill;
            outer.BackColor = Color.White;
            outer.BorderStyle = BorderStyle.FixedSingle;
            outer.Margin = new Padding(0, 6, 0, 6);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 2;
            layout.Margin = new Padding(0);
            layout.Padding = new Padding(0);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            FlowLayoutPanel header = new FlowLayoutPanel();
            header.Dock = DockStyle.Fill;
            header.FlowDirection = FlowDirection.LeftToRight;
            header.WrapContents = false;
            header.BackColor = Color.White;
            header.Padding = new Padding(4, 5, 0, 0);
            header.Margin = new Padding(0);

            chkSelectAll = new CheckBox();
            chkSelectAll.Text = "全选";
            chkSelectAll.AutoSize = true;
            chkSelectAll.Enabled = false;
            chkSelectAll.Margin = new Padding(0, 2, 10, 0);
            chkSelectAll.CheckedChanged += delegate { SelectAllChanged(); };
            header.Controls.Add(chkSelectAll);

            Label selectHint = new Label();
            selectHint.Text = "勾选要应用的修改项，然后点【应用选中修改】";
            selectHint.AutoSize = true;
            selectHint.ForeColor = Color.FromArgb(127, 140, 141);
            selectHint.Margin = new Padding(0, 4, 0, 0);
            header.Controls.Add(selectHint);

            Panel host = new Panel();
            host.Dock = DockStyle.Fill;
            host.BackColor = Color.White;
            host.Margin = new Padding(0);

            listIssues = new ListView();
            listIssues.Dock = DockStyle.Fill;
            listIssues.View = View.Details;
            listIssues.FullRowSelect = true;
            listIssues.HideSelection = false;
            listIssues.MultiSelect = false;
            listIssues.GridLines = false;
            listIssues.BorderStyle = BorderStyle.None;
            listIssues.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            listIssues.ShowItemToolTips = true;
            listIssues.CheckBoxes = true;
            listIssues.Font = uiFont;
            listIssues.Columns.Add("类型", 96);
            listIssues.Columns.Add("原文", 140);
            listIssues.Columns.Add("建议", 165);
            colNote = listIssues.Columns.Add("说明", 300);
            listIssues.Columns.Add("位置", 100);
            listIssues.Resize += delegate { AdjustColumns(); };
            listIssues.ItemChecked += delegate { UpdateSelectAllState(); UpdateApplyButton(); };

            lblEmpty = new Label();
            lblEmpty.Dock = DockStyle.Fill;
            lblEmpty.TextAlign = ContentAlignment.MiddleCenter;
            lblEmpty.BackColor = Color.White;
            lblEmpty.ForeColor = Color.FromArgb(149, 165, 166);
            lblEmpty.Font = new Font(uiFont.FontFamily, 10.5F, FontStyle.Regular);
            lblEmpty.Text = EmptySteps;
            lblEmpty.BringToFront();

            host.Controls.Add(listIssues);
            host.Controls.Add(lblEmpty);
            lblEmpty.BringToFront();

            layout.Controls.Add(header, 0, 0);
            layout.Controls.Add(host, 0, 1);
            outer.Controls.Add(layout);
            return outer;
        }

        private Control BuildBottomPanel()
        {
            TableLayoutPanel panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Fill;
            panel.ColumnCount = 2;
            panel.RowCount = 1;
            panel.Margin = new Padding(0);
            panel.Padding = new Padding(0);
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Fill;
            buttons.FlowDirection = FlowDirection.LeftToRight;
            buttons.WrapContents = false;
            buttons.Margin = new Padding(0);
            buttons.Padding = new Padding(0);

            btnOpen = MakeButton("打开临时文件", false);
            btnOpen.Width = 120;
            btnOpen.Height = 34;
            btnOpen.Margin = new Padding(0, 8, 8, 0);
            btnOpen.Click += delegate { OpenTemp(); };
            buttons.Controls.Add(btnOpen);

            btnApply = MakeButton("应用选中修改", false);
            btnApply.Width = 140;
            btnApply.Height = 34;
            btnApply.Margin = new Padding(0, 8, 8, 0);
            btnApply.Click += delegate { ApplySelected(); };
            buttons.Controls.Add(btnApply);

            btnWrite = MakeButton("写回源文件并删除临时文件", true);
            btnWrite.Width = 238;
            btnWrite.Height = 34;
            btnWrite.Margin = new Padding(0, 8, 0, 0);
            btnWrite.Font = boldFont;
            btnWrite.Click += delegate { WriteBack(); };
            buttons.Controls.Add(btnWrite);

            panel.Controls.Add(buttons, 0, 0);

            progress = new ProgressBar();
            progress.Width = 130;
            progress.Height = 12;
            progress.Style = ProgressBarStyle.Marquee;
            progress.Visible = false;
            progress.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            progress.Margin = new Padding(0, 19, 8, 0);
            panel.Controls.Add(progress, 1, 0);

            return panel;
        }

        private Control BuildStatusPanel()
        {
            Panel panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.BackColor = Color.FromArgb(234, 242, 248);
            panel.Margin = new Padding(0);

            lblStatus = new Label();
            lblStatus.Dock = DockStyle.Fill;
            lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            lblStatus.Padding = new Padding(8, 0, 8, 0);
            lblStatus.AutoEllipsis = true;
            lblStatus.ForeColor = Color.FromArgb(44, 62, 80);
            lblStatus.Text = "就绪：请选择一个 .docx 文件。";
            panel.Controls.Add(lblStatus);
            return panel;
        }

        private Button MakeButton(string text, bool primary)
        {
            Button button = new Button();
            button.Text = text;
            button.AutoSize = false;
            button.Height = 32;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = primary ? Color.FromArgb(31, 97, 141) : Color.FromArgb(189, 195, 199);
            button.BackColor = primary ? Color.FromArgb(41, 128, 185) : Color.FromArgb(250, 251, 252);
            button.ForeColor = primary ? Color.White : Color.FromArgb(44, 62, 80);
            button.UseVisualStyleBackColor = false;
            button.Cursor = Cursors.Hand;
            button.Margin = new Padding(0, 4, 8, 4);

            Color normal = button.BackColor;
            Color normalText = button.ForeColor;
            Color hover = primary ? Color.FromArgb(52, 152, 219) : Color.FromArgb(236, 240, 241);
            button.MouseEnter += delegate { if (button.Enabled) button.BackColor = hover; };
            button.MouseLeave += delegate { if (button.Enabled) button.BackColor = normal; };
            button.EnabledChanged += delegate
            {
                button.BackColor = button.Enabled ? normal : Color.FromArgb(221, 221, 221);
                button.ForeColor = button.Enabled ? normalText : Color.FromArgb(120, 120, 120);
            };
            return button;
        }

        private void AdjustColumns()
        {
            if (colNote == null || listIssues == null) return;
            int fixedWidth = 96 + 140 + 165 + 100 + 28;
            colNote.Width = Math.Max(160, listIssues.ClientSize.Width - fixedWidth);
        }

        private void SelectFile()
        {
            if (busy) return;
            if (tempPath != null && File.Exists(tempPath))
            {
                DialogResult discard = MessageBox.Show(this,
                    "当前临时文件还没有写回。\r\n选择新文件会丢弃它（包括你在 Word 中尚未写回的修改）。\r\n\r\n是否继续？",
                    "选择新文件", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (discard != DialogResult.Yes) return;
                if (!DiscardTemp(tempPath)) return;
            }

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "选择要校对的 Word 文档";
                dialog.Filter = "Word 文档 (*.docx)|*.docx|所有文件 (*.*)|*.*";
                dialog.CheckFileExists = true;
                dialog.Multiselect = false;
                string initial = null;
                if (!string.IsNullOrEmpty(sourcePath)) initial = Path.GetDirectoryName(sourcePath);
                if (string.IsNullOrEmpty(initial) || !Directory.Exists(initial)) initial = Application.StartupPath;
                dialog.InitialDirectory = initial;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                LoadSource(dialog.FileName);
            }
        }

        private void LoadSource(string path)
        {
            if (busy) return;
            if (string.IsNullOrEmpty(path)) return;
            if (!path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, "目前只支持 .docx 格式。\r\n请先在 Word 中把 .doc 文件另存为 .docx。",
                    "格式不支持", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!File.Exists(path))
            {
                MessageBox.Show(this, "文件不存在：" + path, "打开失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!DocxTool.IsValidDocx(path))
            {
                MessageBox.Show(this, "这个文件不是有效的 .docx 文档，或者文件已损坏。", "打开失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (tempPath != null && File.Exists(tempPath))
            {
                DialogResult discard = MessageBox.Show(this,
                    "当前临时文件还没有写回。\r\n选择新文件会丢弃它（包括你在 Word 中尚未写回的修改）。\r\n\r\n是否继续？",
                    "选择新文件", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (discard != DialogResult.Yes) return;
                if (!DiscardTemp(tempPath)) return;
            }

            sourcePath = Path.GetFullPath(path);
            tempPath = null;
            tempCreatedUtc = DateTime.MinValue;
            tempOpenedUtc = DateTime.MinValue;
            waitingForWordOpen = false;
            if (wordOpenTimer != null) wordOpenTimer.Stop();
            scanResult = null;
            listIssues.Items.Clear();
            lblEmpty.Text = EmptySteps;
            lblEmpty.Visible = true;
            lblSummary.Text = "正在准备扫描…";
            txtSource.Text = sourcePath;
            txtSource.SelectionStart = 0;
            txtSource.ScrollToCaret();
            UpdateButtons();
            StartScan(false);
        }

        private string BuildTempPath()
        {
            string dir = Path.GetDirectoryName(sourcePath);
            string name = Path.GetFileNameWithoutExtension(sourcePath);
            return Path.Combine(dir, name + ".校对临时.docx");
        }

        private void Rescan()
        {
            if (busy || string.IsNullOrEmpty(sourcePath)) return;
            if (waitingForWordOpen)
            {
                lblStatus.Text = "Word 正在打开临时文件，请稍等一两秒再点【重新扫描】。";
                return;
            }
            if (!File.Exists(sourcePath))
            {
                MessageBox.Show(this, "源文件不存在，可能已被移动或删除。", "重新扫描",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            StartScan(true);
        }

        private void StartScan(bool rescan)
        {
            if (busy || string.IsNullOrEmpty(sourcePath)) return;
            string temp = BuildTempPath();

            DeepSeekConfig cfg = DeepSeekConfig.LoadOrCreate();
            deepSeekConfig = cfg;
            bool useOnline = chkOnline.Checked;
            if (useOnline && !cfg.HasKey)
            {
                DialogResult setup = MessageBox.Show(this,
                    "你勾选了 DeepSeek 在线校对，但还没有填写 API Key。\r\n\r\n配置文件：\r\n" + cfg.ConfigPath +
                    "\r\n\r\n是否现在打开配置文件？填好并保存后，再点【重新扫描】。",
                    "未配置 API Key", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                if (setup == DialogResult.Yes) OpenApiConfig();
                return;
            }

            if (File.Exists(temp))
            {
                bool modified = true;
                if (rescan && tempPath != null && PathsEqual(temp, tempPath))
                    modified = File.GetLastWriteTimeUtc(temp) > tempCreatedUtc.AddSeconds(2);
                if (modified)
                {
                    DialogResult overwrite = MessageBox.Show(this,
                        "临时文件已存在：\r\n" + temp + "\r\n\r\n重新生成会覆盖它。是否继续？",
                        "重新扫描", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    if (overwrite != DialogResult.Yes) return;
                }
                if (!DiscardTemp(temp)) return;
            }

            ScanOptions options = new ScanOptions();
            options.Punctuation = chkPunctuation.Checked;
            options.Advanced = chkAdvanced.Checked;
            options.Style = true;
            options.Names = true;
            options.Online = useOnline;

            ScanRequest request = new ScanRequest();
            request.Source = sourcePath;
            request.Temp = temp;
            request.Options = options;
            request.DeepSeek = cfg;

            tempPath = null;
            waitingForWordOpen = false;
            if (wordOpenTimer != null) wordOpenTimer.Stop();
            SetBusy(true, useOnline ? "正在扫描，随后请求 DeepSeek…" : "正在扫描并生成批注版临时文件…");
            lblEmpty.Visible = false;
            lblSummary.Text = "正在扫描…";
            listIssues.Items.Clear();

            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs pe)
            {
                if (busy && pe.UserState is string) lblStatus.Text = (string)pe.UserState;
            };
            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                BackgroundWorker bw = (BackgroundWorker)sender;
                ScanRequest req = (ScanRequest)e.Argument;
                List<string> paragraphs = DocxTool.ReadParagraphs(req.Source);
                Scanner scanner = new Scanner();
                ScanResult result = scanner.Scan(paragraphs, req.Options);

                if (req.Options.Online && req.DeepSeek != null && req.DeepSeek.HasKey)
                {
                    OnlineProofreadResult online = DeepSeekProofreader.Proofread(paragraphs, req.DeepSeek,
                        delegate(int done, int total, string message) { bw.ReportProgress(0, message); });
                    if (online.Success)
                    {
                        result.Issues = Scanner.MergeIssues(result.Issues, online.Issues);
                        result.OnlineUsed = true;
                        result.OnlineIssueCount = online.Issues.Count;
                    }
                    else
                    {
                        result.OnlineError = online.Error;
                    }
                }

                string error;
                int annotated = DocxTool.Annotate(req.Source, req.Temp, result.Issues, out error);
                if (annotated < 0) throw new Exception(string.IsNullOrEmpty(error) ? "生成批注版失败。" : error);

                ScanWorkResult work = new ScanWorkResult();
                work.Scan = result;
                work.Annotated = annotated;
                work.Temp = req.Temp;
                work.CreatedUtc = DateTime.UtcNow;
                work.SourceLength = new FileInfo(req.Source).Length;
                work.SourceWriteUtc = File.GetLastWriteTimeUtc(req.Source);
                e.Result = work;
            };
            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                SetBusy(false, string.Empty);
                if (e.Error != null)
                {
                    try { if (File.Exists(temp)) File.Delete(temp); } catch { }
                    lblEmpty.Visible = true;
                    lblEmpty.Text = EmptySteps;
                    lblSummary.Text = "扫描失败";
                    lblStatus.Text = "扫描失败：" + e.Error.Message;
                    MessageBox.Show(this, "扫描失败：\r\n" + e.Error.Message, "扫描失败",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateButtons();
                    return;
                }

                ScanWorkResult work = (ScanWorkResult)e.Result;
                scanResult = work.Scan;
                tempPath = work.Temp;
                waitingForWordOpen = true;
                tempCreatedUtc = work.CreatedUtc;
                sourceLength = work.SourceLength;
                sourceWriteUtc = work.SourceWriteUtc;
                PopulateIssues(work.Scan);
                UpdateSummary(work.Scan, work.Annotated);
                UpdateButtons();

                if (!string.IsNullOrEmpty(work.Scan.OnlineError))
                {
                    lblStatus.Text = "DeepSeek 在线校对失败，已使用离线规则库结果。";
                    MessageBox.Show(this,
                        "DeepSeek 在线校对失败：\r\n" + work.Scan.OnlineError +
                        "\r\n\r\n已自动使用离线规则库结果，批注仍可正常使用。",
                        "在线校对失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (work.Scan.OnlineUsed)
                {
                    if (work.Scan.Issues.Count == 0)
                        lblStatus.Text = "DeepSeek 在线校对完成，未发现明显问题；已生成副本。";
                    else
                        lblStatus.Text = "已生成 " + work.Annotated + " 条批注（含 DeepSeek 在线建议）并打开临时文件。修改后可按 Ctrl+S 保存，再写回。";
                }
                else if (work.Scan.Issues.Count == 0)
                {
                    lblStatus.Text = "未发现明显问题；已生成副本，可按需查看或直接关闭。";
                }
                else
                {
                    lblStatus.Text = "已生成 " + work.Annotated + " 条批注并打开临时文件。修改后请按 Ctrl+S 保存，再点【写回源文件】。";
                }
                OpenTemp();
            };
            worker.RunWorkerAsync(request);
        }

        private void PopulateIssues(ScanResult result)
        {
            listIssues.BeginUpdate();
            try
            {
                listIssues.Items.Clear();
                if (result != null)
                {
                    foreach (Issue issue in result.Issues)
                    {
                        ListViewItem item = new ListViewItem(issue.Type);
                        item.SubItems.Add(Shorten(ShowWhitespace(issue.Original), 80));
                        item.SubItems.Add(Shorten(ShowWhitespace(issue.Suggestion), 80));
                        item.SubItems.Add(Shorten(issue.Note, 240));
                        item.SubItems.Add(issue.Location);
                        item.ForeColor = IssueColor(issue.Type);
                        item.ToolTipText = issue.Note;
                        item.Tag = issue;
                        item.Checked = false;
                        listIssues.Items.Add(item);
                    }
                }
            }
            finally
            {
                listIssues.EndUpdate();
            }
            updatingSelectAll = true;
            try { chkSelectAll.Checked = false; }
            finally { updatingSelectAll = false; }
            chkSelectAll.Enabled = listIssues.Items.Count > 0;
            lblEmpty.Text = result != null && result.Issues.Count == 0 ? EmptyNoIssues : EmptySteps;
            lblEmpty.Visible = listIssues.Items.Count == 0;
            AdjustColumns();
            UpdateApplyButton();
        }

        private static string Shorten(string text, int max)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            string s = text.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
            if (s.Length <= max) return s;
            return s.Substring(0, max - 1) + "…";
        }

        private static string ShowWhitespace(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return text.Replace(" ", "[空格]").Replace("\t", "[空格]").Replace("\r", string.Empty).Replace("\n", string.Empty);
        }

        private static Color IssueColor(string type)
        {
            if (type == Scanner.Typo) return Color.FromArgb(192, 57, 43);
            if (type == Scanner.Lex) return Color.FromArgb(185, 119, 14);
            if (type == Scanner.Redundant) return Color.FromArgb(160, 64, 0);
            if (type == Scanner.Name) return Color.FromArgb(125, 60, 152);
            if (type == Scanner.Punctuation) return Color.FromArgb(86, 101, 115);
            if (type == Scanner.Style) return Color.FromArgb(36, 113, 163);
            return Color.FromArgb(44, 62, 80);
        }

        private void UpdateSummary(ScanResult result, int annotated)
        {
            if (result == null)
            {
                lblSummary.Text = "尚未选择文件";
                return;
            }
            int count = result.Issues == null ? 0 : result.Issues.Count;
            StringBuilder sb = new StringBuilder();
            if (count == 0)
            {
                sb.Append("未发现明显问题");
            }
            else
            {
                sb.Append("共发现 ").Append(count).Append(" 处");
                sb.Append("（");
                string[] order = new string[] { Scanner.Typo, Scanner.Lex, Scanner.Redundant, Scanner.Name, Scanner.Punctuation, Scanner.Style };
                bool first = true;
                foreach (string type in order)
                {
                    int n = result.Issues.Count(i => i.Type == type);
                    if (n == 0) continue;
                    if (!first) sb.Append(" · ");
                    sb.Append(type).Append(' ').Append(n);
                    first = false;
                }
                sb.Append("）");
            }
            if (annotated != count)
            {
                sb.Append(" ｜ 已标注 ").Append(annotated).Append(" 处");
            }
            sb.Append(" ｜ 全文 ").Append(result.CharacterCount).Append(" 字");
            if (result.OnlineUsed) sb.Append(" ｜ DeepSeek 在线");
            else if (!string.IsNullOrEmpty(result.OnlineError)) sb.Append(" ｜ 在线失败，已用离线结果");
            lblSummary.Text = sb.ToString();
        }

        private void OpenTemp()
        {
            if (tempPath == null || !File.Exists(tempPath))
            {
                waitingForWordOpen = false;
                if (wordOpenTimer != null) wordOpenTimer.Stop();
                MessageBox.Show(this, "临时文件不存在，可能已被删除。", "打开临时文件",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                UpdateButtons();
                return;
            }
            try
            {
                waitingForWordOpen = true;
                lblStatus.Text = "正在用 Word 打开临时文件…";
                this.Refresh();
                UpdateButtons();

                string openError;
                bool openedByCom = WordComHelper.OpenDocument(tempPath, out openError);
                tempOpenedUtc = DateTime.UtcNow;

                if (openedByCom)
                {
                    waitingForWordOpen = false;
                    if (wordOpenTimer != null) wordOpenTimer.Stop();
                    lblStatus.Text = "已打开临时文件。修改后按 Ctrl+S 保存，再回到本窗口点【写回源文件】。";
                    UpdateButtons();
                    return;
                }

                ProcessStartInfo info = new ProcessStartInfo(tempPath);
                info.UseShellExecute = true;
                info.WorkingDirectory = Path.GetDirectoryName(tempPath);
                Process.Start(info);

                bool alreadyOpen = WordComHelper.IsDocumentOpen(tempPath) || WordComHelper.IsLocked(tempPath);
                if (alreadyOpen)
                {
                    waitingForWordOpen = false;
                    if (wordOpenTimer != null) wordOpenTimer.Stop();
                    lblStatus.Text = "已打开临时文件。修改后按 Ctrl+S 保存，再回到本窗口点【写回源文件】。";
                }
                else
                {
                    waitingForWordOpen = true;
                    wordOpenWaitStart = DateTime.UtcNow;
                    if (wordOpenTimer != null) wordOpenTimer.Start();
                    lblStatus.Text = "已启动 Word，正在等待临时文件打开…";
                }
                UpdateButtons();
            }
            catch (Exception ex)
            {
                waitingForWordOpen = false;
                if (wordOpenTimer != null) wordOpenTimer.Stop();
                MessageBox.Show(this, "无法自动打开临时文件：\r\n" + ex.Message + "\r\n\r\n文件位置：\r\n" + tempPath,
                    "打开失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lblStatus.Text = "临时文件已生成，但自动打开失败：" + tempPath;
                UpdateButtons();
            }
        }

        private void WriteBack()
        {
            if (busy || tempPath == null || sourcePath == null) return;
            if (waitingForWordOpen)
            {
                lblStatus.Text = "Word 正在打开临时文件，请稍等一两秒再点【写回源文件】。";
                return;
            }
            if (!File.Exists(tempPath))
            {
                MessageBox.Show(this, "临时文件不存在，可能已被删除或移动。", "无法写回",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tempPath = null;
                UpdateButtons();
                return;
            }
            if (!File.Exists(sourcePath))
            {
                MessageBox.Show(this, "源文件不存在，可能已被移动或删除。", "无法写回",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            WaitForTempOpenGrace();

            if (WordComHelper.IsLocked(sourcePath))
            {
                MessageBox.Show(this,
                    "源文件正在 Word 中打开。\r\n\r\n请先保存并关闭源文件，再点【写回源文件】。\r\n" + sourcePath,
                    "源文件被占用", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string closeError;
            CloseDocumentResult closeResult = WordComHelper.CloseIfOpenWithRetry(tempPath, true, 4, out closeError);
            if (closeResult == CloseDocumentResult.Failed && WordComHelper.IsLocked(tempPath))
            {
                MessageBox.Show(this,
                    "无法保存并关闭临时文件：\r\n" + closeError + "\r\n\r\n请先在 Word 中按 Ctrl+S 保存并关闭，然后重试。",
                    "临时文件被占用", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (closeResult == CloseDocumentResult.Closed) Thread.Sleep(350);

            if (WordComHelper.IsLocked(tempPath))
            {
                MessageBox.Show(this,
                    "临时文件仍被 Word 占用。\r\n\r\n请先在 Word 中保存并关闭临时文件，然后重试。",
                    "临时文件被占用", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DateTime nowWrite = File.GetLastWriteTimeUtc(sourcePath);
            long nowLength = new FileInfo(sourcePath).Length;
            if (nowWrite != sourceWriteUtc || nowLength != sourceLength)
            {
                DialogResult overwrite = MessageBox.Show(this,
                    "源文件在扫描后已被修改过。\r\n继续写回会用临时文件的内容覆盖这些修改。\r\n\r\n是否继续？",
                    "源文件已变化", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (overwrite != DialogResult.Yes) return;
            }

            bool makeBackup = chkBackup.Checked;
            if (!makeBackup)
            {
                DialogResult confirm = MessageBox.Show(this,
                    "你没有勾选“写回前备份源文件”。\r\n写回后无法从本工具恢复旧版本，是否继续？",
                    "确认写回", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (confirm != DialogResult.Yes) return;
            }

            MergeRequest request = new MergeRequest();
            request.Temp = tempPath;
            request.Source = sourcePath;
            request.MakeBackup = makeBackup;

            SetBusy(true, "正在把临时文件正文写回源文件…");
            worker = new BackgroundWorker();
            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                MergeRequest req = (MergeRequest)e.Argument;
                e.Result = DocxTool.Merge(req.Temp, req.Source, req.MakeBackup);
            };
            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                SetBusy(false, string.Empty);
                if (e.Error != null)
                {
                    lblStatus.Text = "写回失败：" + e.Error.Message;
                    MessageBox.Show(this, "写回失败：\r\n" + e.Error.Message, "写回失败",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateButtons();
                    return;
                }

                MergeResult result = (MergeResult)e.Result;
                if (result == null || !result.Success)
                {
                    string message = result == null ? "未知错误。" : result.Message;
                    lblStatus.Text = "写回失败：" + message;
                    MessageBox.Show(this, "写回失败：\r\n" + message, "写回失败",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateButtons();
                    return;
                }

                string oldTemp = tempPath;
                tempPath = null;
                tempOpenedUtc = DateTime.MinValue;
                UpdateButtons();
                lblSummary.Text = "已写回源文件，临时文件已删除。";
                lblStatus.Text = result.Message;
                StringBuilder messageText = new StringBuilder();
                messageText.Append("写回完成，临时文件已删除。");
                if (!string.IsNullOrEmpty(result.BackupPath))
                    messageText.Append("\r\n\r\n备份文件：\r\n").Append(result.BackupPath);
                if (!result.TempDeleted)
                    messageText.Append("\r\n\r\n注意：临时文件被占用，未能删除，请手动删除：\r\n").Append(oldTemp);
                MessageBox.Show(this, messageText.ToString(), "写回完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            worker.RunWorkerAsync(request);
        }

        private bool DiscardTemp(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return true;
            if (waitingForWordOpen)
            {
                lblStatus.Text = "Word 正在打开临时文件，请稍等一两秒再操作。";
                return false;
            }
            WaitForTempOpenGrace();

            string error;
            CloseDocumentResult result = WordComHelper.CloseIfOpenWithRetry(path, false, 4, out error);
            if (result == CloseDocumentResult.Closed) Thread.Sleep(250);

            if (WordComHelper.IsLocked(path))
            {
                result = WordComHelper.CloseIfOpenWithRetry(path, false, 3, out error);
                Thread.Sleep(250);
                if (WordComHelper.IsLocked(path))
                {
                    MessageBox.Show(this,
                        "临时文件仍被 Word 占用，暂时无法重新生成：\r\n" + path +
                        "\r\n\r\n请稍等一两秒再点一次，或先在 Word 中关闭该临时文件。",
                        "临时文件被占用", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            if (!WordComHelper.TryDelete(path, 8))
            {
                WordComHelper.CloseIfOpenWithRetry(path, false, 2, out error);
                Thread.Sleep(250);
                if (!WordComHelper.TryDelete(path, 8))
                {
                    MessageBox.Show(this,
                        "无法删除临时文件，可能仍被 Word 占用：\r\n" + path + "\r\n\r\n请关闭该文件后重试。",
                        "无法删除临时文件", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }
            if (tempPath != null && PathsEqual(tempPath, path))
            {
                tempPath = null;
                tempOpenedUtc = DateTime.MinValue;
                UpdateButtons();
            }
            return true;
        }

        private void WaitForTempOpenGrace()
        {
            if (tempOpenedUtc == DateTime.MinValue) return;
            int elapsed = (int)(DateTime.UtcNow - tempOpenedUtc).TotalMilliseconds;
            if (elapsed >= 0 && elapsed < 1200)
                Thread.Sleep(1200 - elapsed);
        }

        private void CheckWordOpenTimer()
        {
            if (!waitingForWordOpen || string.IsNullOrEmpty(tempPath) || !File.Exists(tempPath))
            {
                waitingForWordOpen = false;
                if (wordOpenTimer != null) wordOpenTimer.Stop();
                UpdateButtons();
                return;
            }

            bool open = WordComHelper.IsDocumentOpen(tempPath) || WordComHelper.IsLocked(tempPath);
            bool timeout = (DateTime.UtcNow - wordOpenWaitStart).TotalSeconds >= 12;
            if (!open && !timeout) return;

            waitingForWordOpen = false;
            if (wordOpenTimer != null) wordOpenTimer.Stop();
            UpdateButtons();
            if (open)
                lblStatus.Text = "已打开临时文件。修改后按 Ctrl+S 保存，再回到本窗口点【写回源文件】。";
            else
                lblStatus.Text = "Word 尚未完成打开临时文件；请稍等，或先处理 Word 窗口后再操作。";
        }

        private void OpenApiConfig()
        {
            try
            {
                DeepSeekConfig cfg = DeepSeekConfig.LoadOrCreate();
                deepSeekConfig = cfg;
                if (string.IsNullOrEmpty(cfg.ConfigPath))
                {
                    MessageBox.Show(this, "无法确定配置文件位置。", "API 配置",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                ProcessStartInfo info = new ProcessStartInfo("notepad.exe", "\"" + cfg.ConfigPath + "\"");
                info.UseShellExecute = true;
                Process.Start(info);
                lblStatus.Text = "已打开配置文件。填好 ApiKey 并保存后，勾选【DeepSeek 在线校对】再重新扫描。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法打开配置文件：\r\n" + ex.Message, "API 配置",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ApplySelected()
        {
            if (busy || tempPath == null) return;
            if (waitingForWordOpen)
            {
                lblStatus.Text = "Word 正在打开临时文件，请稍等一两秒再应用修改。";
                return;
            }
            if (!File.Exists(tempPath))
            {
                MessageBox.Show(this, "临时文件不存在，可能已被删除或移动。", "应用修改",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tempPath = null;
                UpdateButtons();
                return;
            }

            List<Issue> selected = new List<Issue>();
            foreach (ListViewItem item in listIssues.CheckedItems)
            {
                Issue issue = item.Tag as Issue;
                if (issue != null) selected.Add(issue);
            }
            if (selected.Count == 0)
            {
                MessageBox.Show(this, "请先勾选要应用的修改项（可点【全选】）。", "应用修改",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            WaitForTempOpenGrace();
            SetBusy(true, "正在应用选中的修改…");
            this.Refresh();
            ApplyResult applied = WordApplyHelper.ApplySelected(tempPath, selected);
            SetBusy(false, string.Empty);

            if (!applied.WordAvailable)
            {
                lblStatus.Text = applied.Error;
                MessageBox.Show(this, applied.Error, "无法自动应用",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!applied.Success)
            {
                lblStatus.Text = "应用修改失败：" + applied.Error;
                MessageBox.Show(this, "应用修改失败：\r\n" + applied.Error, "应用失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateButtons();
                return;
            }

            if (scanResult != null)
            {
                foreach (Issue issue in applied.AppliedIssues) scanResult.Issues.Remove(issue);
            }

            List<ListViewItem> removeItems = new List<ListViewItem>();
            foreach (ListViewItem item in listIssues.Items)
            {
                Issue issue = item.Tag as Issue;
                if (issue != null && applied.AppliedIssues.Contains(issue)) removeItems.Add(item);
            }
            foreach (ListViewItem item in removeItems) listIssues.Items.Remove(item);

            UpdateSelectAllState();
            UpdateApplyButton();
            UpdateButtons();

            if (scanResult != null && scanResult.Issues.Count == 0)
            {
                lblEmpty.Text = EmptyApplied;
                lblEmpty.Visible = true;
                lblSummary.Text = "选中的修改已全部应用 ｜ 全文 " + scanResult.CharacterCount + " 字";
            }
            else
            {
                lblEmpty.Visible = listIssues.Items.Count == 0;
                if (scanResult != null) UpdateSummary(scanResult, scanResult.Issues.Count);
            }

            lblStatus.Text = "已应用 " + applied.AppliedIssues.Count + " 处修改" +
                (applied.Failures.Count > 0 ? "；" + applied.Failures.Count + " 处未能应用。" : "。");

            if (applied.Failures.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("以下修改未能自动应用（原文可能已被改动，或对应批注已被删除）：\r\n\r\n");
                int limit = Math.Min(10, applied.Failures.Count);
                for (int i = 0; i < limit; i++) sb.Append("• ").Append(applied.Failures[i]).Append("\r\n");
                if (applied.Failures.Count > limit)
                    sb.Append("……另有 ").Append(applied.Failures.Count - limit).Append(" 处。\r\n");
                sb.Append("\r\n这些修改仍保留在列表中，可以手动修改。");
                MessageBox.Show(this, sb.ToString(), "部分修改未应用",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void MarkOptionsDirty()
        {
            if (busy) return;
            if (sourcePath != null)
                lblStatus.Text = "扫描选项已更改；点【重新扫描】后生效。";
        }

        private void SetBusy(bool value, string text)
        {
            busy = value;
            progress.Visible = value;
            if (value && !string.IsNullOrEmpty(text)) lblStatus.Text = text;
            UpdateButtons();
        }

        private void UpdateButtons()
        {
            bool hasTemp = tempPath != null && File.Exists(tempPath);
            btnSelect.Enabled = !busy;
            btnRescan.Enabled = !busy && !waitingForWordOpen && !string.IsNullOrEmpty(sourcePath);
            btnOpen.Enabled = !busy && hasTemp;
            btnWrite.Enabled = !busy && !waitingForWordOpen && hasTemp && !string.IsNullOrEmpty(sourcePath);
            chkPunctuation.Enabled = !busy;
            chkAdvanced.Enabled = !busy;
            chkBackup.Enabled = !busy;
            chkOnline.Enabled = !busy;
            linkApi.Enabled = !busy;
            chkSelectAll.Enabled = !busy && listIssues != null && listIssues.Items.Count > 0;
            UpdateApplyButton();
        }

        private void UpdateApplyButton()
        {
            btnApply.Enabled = !busy && !waitingForWordOpen && tempPath != null && File.Exists(tempPath) && CheckedIssueCount() > 0;
        }

        private int CheckedIssueCount()
        {
            int count = 0;
            foreach (ListViewItem item in listIssues.CheckedItems)
            {
                if (item.Tag is Issue) count++;
            }
            return count;
        }

        private void SelectAllChanged()
        {
            if (updatingSelectAll) return;
            updatingSelectAll = true;
            try
            {
                bool value = chkSelectAll.Checked;
                foreach (ListViewItem item in listIssues.Items)
                {
                    if (item.Tag is Issue) item.Checked = value;
                }
            }
            finally
            {
                updatingSelectAll = false;
            }
            UpdateApplyButton();
        }

        private void UpdateSelectAllState()
        {
            if (updatingSelectAll) return;
            updatingSelectAll = true;
            try
            {
                int total = 0;
                int checkedCount = 0;
                foreach (ListViewItem item in listIssues.Items)
                {
                    if (!(item.Tag is Issue)) continue;
                    total++;
                    if (item.Checked) checkedCount++;
                }
                chkSelectAll.Checked = total > 0 && checkedCount == total;
            }
            finally
            {
                updatingSelectAll = false;
            }
        }

        private void ShowHelp()
        {
            string text =
                "使用方法\r\n" +
                "1. 点【选择 Word 文件…】，选中 .docx 后会自动扫描。\r\n" +
                "2. 程序会生成并打开“文件名.校对临时.docx”，错别字和词法问题以 Word 批注和淡黄底纹标出。\r\n" +
                "3. 可在列表中勾选（或点【全选】）修改项，再点【应用选中修改】直接改到临时文件；也可以直接在 Word 中手动修改。\r\n" +
                "4. 按 Ctrl+S 保存后，回到本窗口点【写回源文件并删除临时文件】。程序只把正文写回源文件，批注和底纹不会带入。\r\n\r\n" +
                "DeepSeek 在线校对\r\n" +
                "· 点【API 配置】打开 小说校对助手.ini，填入 DeepSeek API Key；\r\n" +
                "· 勾选【DeepSeek 在线校对】后重新扫描，正文会分片发送到 DeepSeek API；\r\n" +
                "· 不勾选时只使用离线规则库，不联网；\r\n" +
                "· 在线请求失败会自动回退到离线结果，并给出提示。\r\n\r\n" +
                "注意\r\n" +
                "· 仅支持 .docx；.doc 请先在 Word 中另存为 .docx。\r\n" +
                "· 批注只是修改建议，请自行判断；离线规则库和在线模型都可能误报。\r\n" +
                "· 写回前默认保留“文件名.校对备份.docx”，可取消勾选。\r\n" +
                "· 写回时源文件必须已经关闭，否则会提示被占用。\r\n" +
                "· 关闭程序时如果临时文件还在，会询问保留、删除还是取消退出。";
            MessageBox.Show(this, text, "使用说明", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static bool PathsEqual(string a, string b)
        {
            try
            {
                return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.O)
            {
                SelectFile();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.F5)
            {
                Rescan();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.S)
            {
                WriteBack();
                e.Handled = true;
            }
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (busy) return;
            if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length == 1 && files[0].EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    e.Effect = DragDropEffects.Copy;
            }
        }

        private void OnDragDrop(object sender, DragEventArgs e)
        {
            if (busy) return;
            if (e.Data == null || !e.Data.GetDataPresent(DataFormats.FileDrop)) return;
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length != 1) return;
            LoadSource(files[0]);
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (busy)
            {
                e.Cancel = true;
                MessageBox.Show(this, "正在处理文件，请稍候再关闭。", "请稍候",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (tempPath != null && File.Exists(tempPath))
            {
                string path = tempPath;
                DialogResult result = MessageBox.Show(this,
                    "临时文件还没有写回源文件：\r\n" + path +
                    "\r\n\r\n关闭程序前是否保留临时文件？\r\n\r\n" +
                    "是：保留临时文件并退出\r\n" +
                    "否：删除临时文件并退出\r\n" +
                    "取消：不退出",
                    "关闭程序", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);

                if (result == DialogResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
                if (result == DialogResult.No)
                {
                    waitingForWordOpen = false;
                    if (wordOpenTimer != null) wordOpenTimer.Stop();
                    UpdateButtons();
                    if (!DiscardTemp(path))
                    {
                        e.Cancel = true;
                        return;
                    }
                }
                // Yes: keep the temporary file and close normally.
            }
        }

        private class ScanRequest
        {
            public string Source;
            public string Temp;
            public ScanOptions Options;
            public DeepSeekConfig DeepSeek;
        }

        private class ScanWorkResult
        {
            public ScanResult Scan;
            public int Annotated;
            public string Temp;
            public DateTime CreatedUtc;
            public long SourceLength;
            public DateTime SourceWriteUtc;
        }

        private class MergeRequest
        {
            public string Temp;
            public string Source;
            public bool MakeBackup;
        }
    }
}
