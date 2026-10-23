using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

internal static class Launcher
{
    [STAThread]
    private static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new ConverterWindow(args));
    }
}

internal sealed class ConverterWindow : Form
{
    private readonly string folder;
    private readonly string[] startupFiles;
    private readonly string python;
    private readonly string backend;
    private readonly Panel dropPanel;
    private readonly Label dropTitle;
    private readonly Label dropHint;
    private readonly Label dropBadge;
    private readonly Button chooseButton;
    private readonly ComboBox modeBox;
    private readonly ComboBox dpiBox;
    private readonly Label dpiLabel;
    private readonly TextBox outputBox;
    private readonly ProgressBar progress;
    private readonly Label status;
    private readonly Button folderButton;
    private readonly Button resetButton;
    private readonly Button openOutputButton;
    private readonly BackgroundWorker worker;
    private string outputFolder = "";
    private string lastGeneratedFile = "";

    internal ConverterWindow(string[] args)
    {
        folder = AppDomain.CurrentDomain.BaseDirectory;
        startupFiles = args ?? new string[0];
        python = Path.Combine(folder, "runtime", "python.exe");
        backend = Path.Combine(folder, "main.py");
        Text = "PDF 转 Word";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(780, 700);
        MinimumSize = new Size(700, 650);
        BackColor = Color.FromArgb(243, 246, 251);
        Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular);

        string iconPath = Path.Combine(folder, "app.ico");
        if (File.Exists(iconPath))
        {
            try { Icon = new Icon(iconPath); }
            catch { }
        }

        TableLayoutPanel layout = new TableLayoutPanel();
        layout.Dock = DockStyle.Fill;
        layout.Padding = new Padding(34, 24, 34, 22);
        layout.ColumnCount = 1;
        layout.RowCount = 4;
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 188F));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 76F));
        Controls.Add(layout);

        Panel header = new Panel();
        header.Dock = DockStyle.Fill;
        Label title = new Label();
        title.Text = "PDF 转 Word";
        title.Font = new Font("Microsoft YaHei UI", 23F, FontStyle.Bold);
        title.ForeColor = Color.FromArgb(24, 34, 48);
        title.AutoSize = true;
        title.Location = new Point(0, 0);
        header.Controls.Add(title);
        Label subtitle = new Label();
        subtitle.Text = "保留可编辑文字、公式、表格与插图";
        subtitle.Font = new Font("Microsoft YaHei UI", 9.5F);
        subtitle.ForeColor = Color.FromArgb(102, 112, 133);
        subtitle.AutoSize = true;
        subtitle.Location = new Point(2, 43);
        header.Controls.Add(subtitle);
        Label privacy = new Label();
        privacy.Text = "本地转换 · 文件不上传";
        privacy.TextAlign = ContentAlignment.MiddleCenter;
        privacy.Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Bold);
        privacy.ForeColor = Color.FromArgb(29, 118, 83);
        privacy.BackColor = Color.FromArgb(230, 247, 239);
        privacy.Size = new Size(186, 31);
        privacy.Location = new Point(Width - 255, 9);
        privacy.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        header.Controls.Add(privacy);
        layout.Controls.Add(header, 0, 0);

        dropPanel = new Panel();
        dropPanel.Dock = DockStyle.Fill;
        dropPanel.BackColor = Color.White;
        dropPanel.BorderStyle = BorderStyle.FixedSingle;
        dropBadge = new Label();
        dropBadge.Text = "PDF";
        dropBadge.TextAlign = ContentAlignment.MiddleCenter;
        dropBadge.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold);
        dropBadge.ForeColor = Color.FromArgb(37, 99, 235);
        dropBadge.BackColor = Color.FromArgb(235, 243, 255);
        dropBadge.Size = new Size(52, 32);
        dropTitle = new Label();
        dropTitle.Text = "拖放 PDF 文件到这里";
        dropTitle.Font = new Font("Microsoft YaHei UI", 16F, FontStyle.Bold);
        dropTitle.ForeColor = Color.FromArgb(24, 34, 48);
        dropTitle.AutoSize = true;
        dropHint = new Label();
        dropHint.Text = "支持一次添加多个文件 · 也可以点击下方按钮浏览";
        dropHint.Font = new Font("Microsoft YaHei UI", 9F);
        dropHint.ForeColor = Color.FromArgb(102, 112, 133);
        dropHint.AutoSize = true;
        chooseButton = new Button();
        chooseButton.Text = "选择 PDF 文件";
        chooseButton.Size = new Size(150, 42);
        chooseButton.FlatStyle = FlatStyle.Flat;
        chooseButton.FlatAppearance.BorderSize = 0;
        chooseButton.BackColor = Color.FromArgb(37, 99, 235);
        chooseButton.ForeColor = Color.White;
        chooseButton.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold);
        chooseButton.Cursor = Cursors.Hand;
        chooseButton.Padding = new Padding(5, 1, 5, 1);
        chooseButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(29, 78, 216);
        chooseButton.Click += ChooseFiles;
        dropPanel.Controls.Add(dropBadge);
        dropPanel.Controls.Add(dropTitle);
        dropPanel.Controls.Add(dropHint);
        dropPanel.Controls.Add(chooseButton);
        dropPanel.Resize += CenterDropControls;
        layout.Controls.Add(dropPanel, 0, 1);

        Panel settings = new Panel();
        settings.Dock = DockStyle.Fill;
        settings.BackColor = Color.White;
        settings.BorderStyle = BorderStyle.FixedSingle;
        layout.Controls.Add(settings, 0, 2);
        Label settingsTitle = MakeLabel("转换设置", 18, 15, 12F, true, Color.FromArgb(24, 34, 48));
        settings.Controls.Add(settingsTitle);
        settings.Controls.Add(MakeLabel("转换模式", 21, 59, 9.5F, false, Color.FromArgb(52, 64, 84)));
        modeBox = new ComboBox();
        modeBox.DropDownStyle = ComboBoxStyle.DropDownList;
        modeBox.Items.AddRange(new object[] { "可编辑文字", "外观保真（图片）" });
        modeBox.SelectedIndex = 0;
        modeBox.Location = new Point(120, 54);
        modeBox.Size = new Size(300, 30);
        modeBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        modeBox.SelectedIndexChanged += ModeChanged;
        settings.Controls.Add(modeBox);
        dpiLabel = MakeLabel("图片清晰度", 443, 59, 9F, false, Color.FromArgb(176, 184, 196));
        dpiLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        settings.Controls.Add(dpiLabel);
        dpiBox = new ComboBox();
        dpiBox.DropDownStyle = ComboBoxStyle.DropDownList;
        dpiBox.Items.AddRange(new object[] { "300", "600" });
        dpiBox.SelectedIndex = 0;
        dpiBox.Enabled = false;
        dpiBox.Location = new Point(524, 54);
        dpiBox.Size = new Size(64, 30);
        dpiBox.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        settings.Controls.Add(dpiBox);
        settings.Controls.Add(MakeLabel("DPI", 598, 59, 9F, false, Color.FromArgb(102, 112, 133)));

        settings.Controls.Add(MakeLabel("输出位置", 21, 111, 9.5F, false, Color.FromArgb(52, 64, 84)));
        outputBox = new TextBox();
        outputBox.ReadOnly = true;
        outputBox.Text = "与 PDF 同文件夹（默认）";
        outputBox.Location = new Point(120, 105);
        outputBox.Size = new Size(340, 29);
        outputBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        outputBox.BackColor = Color.FromArgb(248, 250, 252);
        settings.Controls.Add(outputBox);
        folderButton = new Button();
        folderButton.Text = "选择文件夹…";
        folderButton.Location = new Point(470, 103);
        folderButton.Size = new Size(116, 33);
        folderButton.FlatStyle = FlatStyle.Flat;
        folderButton.FlatAppearance.BorderSize = 0;
        folderButton.BackColor = Color.FromArgb(235, 243, 255);
        folderButton.ForeColor = Color.FromArgb(37, 99, 235);
        folderButton.Cursor = Cursors.Hand;
        folderButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        folderButton.Click += ChooseOutputFolder;
        settings.Controls.Add(folderButton);
        resetButton = new Button();
        resetButton.Text = "恢复默认";
        resetButton.Location = new Point(594, 103);
        resetButton.Size = new Size(96, 33);
        resetButton.FlatStyle = FlatStyle.Flat;
        resetButton.FlatAppearance.BorderColor = Color.FromArgb(220, 226, 235);
        resetButton.BackColor = Color.White;
        resetButton.Cursor = Cursors.Hand;
        resetButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        resetButton.Click += ResetOutputFolder;
        settings.Controls.Add(resetButton);
        Label outputHint = MakeLabel("留空时，每个 Word 文件会保存到对应 PDF 所在文件夹。", 120, 143,
            8.5F, false, Color.FromArgb(102, 112, 133));
        settings.Controls.Add(outputHint);

        Panel footer = new Panel();
        footer.Dock = DockStyle.Fill;
        progress = new ProgressBar();
        progress.Location = new Point(0, 4);
        progress.Size = new Size(710, 12);
        progress.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        progress.Minimum = 0;
        progress.Maximum = 1;
        footer.Controls.Add(progress);
        status = MakeLabel("就绪 · 选择 PDF 或拖放文件开始转换", 1, 25, 9F, false,
            Color.FromArgb(102, 112, 133));
        status.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        status.AutoEllipsis = true;
        status.Size = new Size(520, status.Height);
        footer.Controls.Add(status);
        Label note = MakeLabel("正文与表格可编辑；复杂公式保留为高清图片。", 1, 49,
            8.5F, false, Color.FromArgb(152, 162, 179));
        footer.Controls.Add(note);
        openOutputButton = new Button();
        openOutputButton.Text = "打开生成文件夹";
        openOutputButton.Location = new Point(548, 24);
        openOutputButton.Size = new Size(150, 31);
        openOutputButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        openOutputButton.Enabled = false;
        openOutputButton.FlatStyle = FlatStyle.Flat;
        openOutputButton.FlatAppearance.BorderColor = Color.FromArgb(211, 222, 238);
        openOutputButton.BackColor = Color.White;
        openOutputButton.ForeColor = Color.FromArgb(37, 99, 235);
        openOutputButton.Cursor = Cursors.Hand;
        openOutputButton.Click += OpenGeneratedFile;
        footer.Controls.Add(openOutputButton);
        layout.Controls.Add(footer, 0, 3);

        worker = new BackgroundWorker();
        worker.WorkerReportsProgress = true;
        worker.DoWork += ConvertFiles;
        worker.ProgressChanged += ConversionProgress;
        worker.RunWorkerCompleted += ConversionCompleted;

        EnableDropRecursively(this);
        Load += delegate { CenterDropControls(null, EventArgs.Empty); };
        Shown += delegate
        {
            if (startupFiles.Length > 0)
                BeginConversion(startupFiles);
        };
    }

    private static Label MakeLabel(string text, int x, int y, float size, bool bold, Color color)
    {
        Label label = new Label();
        label.Text = text;
        label.Location = new Point(x, y);
        label.AutoSize = true;
        label.Font = new Font("Microsoft YaHei UI", size, bold ? FontStyle.Bold : FontStyle.Regular);
        label.ForeColor = color;
        return label;
    }

    private void EnableDropRecursively(Control control)
    {
        control.AllowDrop = true;
        control.DragEnter += OnDragEnterFiles;
        control.DragDrop += OnDropFiles;
        foreach (Control child in control.Controls)
            EnableDropRecursively(child);
    }

    private void CenterDropControls(object sender, EventArgs e)
    {
        if (dropPanel == null || dropTitle == null || dropHint == null || chooseButton == null)
            return;
        int center = dropPanel.ClientSize.Width / 2;
        dropBadge.Location = new Point(center - dropBadge.Width / 2, 29);
        dropTitle.Location = new Point(center - dropTitle.Width / 2, 72);
        dropHint.Location = new Point(center - dropHint.Width / 2, 107);
        chooseButton.Location = new Point(center - chooseButton.Width / 2, 139);
    }

    private void ModeChanged(object sender, EventArgs e)
    {
        bool visual = modeBox.SelectedIndex == 1;
        dpiBox.Enabled = visual;
        dpiLabel.ForeColor = visual ? Color.FromArgb(52, 64, 84) : Color.FromArgb(176, 184, 196);
    }

    private void ChooseFiles(object sender, EventArgs e)
    {
        OpenFileDialog dialog = new OpenFileDialog();
        dialog.Title = "选择 PDF 文件";
        dialog.Filter = "PDF 文件 (*.pdf)|*.pdf|所有文件 (*.*)|*.*";
        dialog.Multiselect = true;
        if (dialog.ShowDialog(this) == DialogResult.OK)
            BeginConversion(dialog.FileNames);
    }

    private void ChooseOutputFolder(object sender, EventArgs e)
    {
        using (FolderBrowserDialog dialog = new FolderBrowserDialog())
        {
            dialog.Description = "选择 Word 文件保存位置";
            if (Directory.Exists(outputFolder))
                dialog.SelectedPath = outputFolder;
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                outputFolder = dialog.SelectedPath;
                outputBox.Text = outputFolder;
                status.Text = "Word 文件将保存到所选文件夹。";
            }
        }
    }

    private void ResetOutputFolder(object sender, EventArgs e)
    {
        outputFolder = "";
        outputBox.Text = "与 PDF 同文件夹（默认）";
        status.Text = "已恢复默认：每个 Word 文件保存到对应 PDF 所在文件夹。";
    }

    private void OnDragEnterFiles(object sender, DragEventArgs e)
    {
        e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
    }

    private void OnDropFiles(object sender, DragEventArgs e)
    {
        string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
        if (files != null)
            BeginConversion(files);
    }

    private void OpenGeneratedFile(object sender, EventArgs e)
    {
        if (String.IsNullOrEmpty(lastGeneratedFile) || !File.Exists(lastGeneratedFile))
            return;
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = "/select,\"" + lastGeneratedFile + "\"",
            UseShellExecute = true
        });
    }

    private void BeginConversion(string[] paths)
    {
        if (worker.IsBusy)
        {
            MessageBox.Show(this, "当前转换完成后再添加其他文件。", "正在转换",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        List<string> pdfs = new List<string>();
        foreach (string path in paths)
            if (String.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
                pdfs.Add(path);
        if (pdfs.Count == 0)
        {
            MessageBox.Show(this, "没有找到 PDF 文件，请重新选择。", "请选择 PDF",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        if (!File.Exists(python) || !File.Exists(backend))
        {
            MessageBox.Show(this, "程序文件不完整，请重新解压整个压缩包。", "无法启动",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        modeBox.Enabled = false;
        dpiBox.Enabled = false;
        folderButton.Enabled = false;
        resetButton.Enabled = false;
        chooseButton.Enabled = false;
        openOutputButton.Enabled = false;
        chooseButton.Text = "正在转换…";
        progress.Value = 0;
        progress.Maximum = pdfs.Count;
        status.Text = "准备转换 " + pdfs.Count + " 个文件…";
        worker.RunWorkerAsync(new Job(pdfs.ToArray(), outputFolder, modeBox.SelectedIndex == 1, dpiBox.Text));
    }

    private void ConvertFiles(object sender, DoWorkEventArgs e)
    {
        Job job = (Job)e.Argument;
        Result result = new Result();
        for (int index = 0; index < job.Files.Length; index++)
        {
            string source = job.Files[index];
            string label = job.Visual ? "保真版" : "可编辑版";
            string destinationFolder = String.IsNullOrEmpty(job.OutputFolder)
                ? Path.GetDirectoryName(source) : job.OutputFolder;
            string destination = UniquePath(Path.Combine(destinationFolder,
                Path.GetFileNameWithoutExtension(source) + "-" + label + ".docx"));
            BackgroundWorker background = (BackgroundWorker)sender;
            background.ReportProgress(index, "正在转换 " + (index + 1) + "/" + job.Files.Length + "：" + Path.GetFileName(source));
            try
            {
                ProcessStartInfo start = new ProcessStartInfo();
                start.FileName = python;
                start.WorkingDirectory = folder;
                start.UseShellExecute = false;
                start.CreateNoWindow = true;
                start.RedirectStandardOutput = true;
                start.RedirectStandardError = true;
                string args = Quote(backend) + " " + Quote(source) + " --output " + Quote(destination);
                if (job.Visual)
                    args += " --visual --dpi " + job.Dpi;
                start.Arguments = args;
                using (Process process = Process.Start(start))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();
                    if (process.ExitCode != 0)
                        throw new InvalidOperationException(String.IsNullOrWhiteSpace(error) ? output.Trim() : error.Trim());
                }
                result.Completed.Add(destination);
            }
            catch (Exception error)
            {
                result.Failures.Add(Path.GetFileName(source) + "：" + error.Message);
            }
            background.ReportProgress(index + 1, null);
        }
        e.Result = result;
    }

    private void ConversionProgress(object sender, ProgressChangedEventArgs e)
    {
        progress.Value = Math.Min(progress.Maximum, Math.Max(progress.Minimum, e.ProgressPercentage));
        if (e.UserState != null)
            status.Text = e.UserState.ToString();
    }

    private void ConversionCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        modeBox.Enabled = true;
        ModeChanged(null, EventArgs.Empty);
        folderButton.Enabled = true;
        resetButton.Enabled = true;
        chooseButton.Enabled = true;
        chooseButton.Text = "选择 PDF 文件";
        if (e.Error != null)
        {
            status.Text = "转换失败，请检查 PDF 文件后重试。";
            MessageBox.Show(this, e.Error.Message, "转换失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        Result result = (Result)e.Result;
        if (result.Completed.Count > 0)
        {
            lastGeneratedFile = result.Completed[result.Completed.Count - 1];
            openOutputButton.Enabled = true;
        }
        if (result.Completed.Count > 0 && result.Failures.Count == 0)
        {
            status.Text = "转换完成 · 共生成 " + result.Completed.Count + " 个 Word 文件";
            MessageBox.Show(this, "已生成文件：\r\n\r\n" + String.Join("\r\n", result.Completed.ToArray()),
                "转换完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else if (result.Completed.Count > 0)
        {
            status.Text = "部分完成 · 成功 " + result.Completed.Count + " 个，失败 " + result.Failures.Count + " 个";
            MessageBox.Show(this, "已生成：\r\n" + String.Join("\r\n", result.Completed.ToArray())
                + "\r\n\r\n未能转换：\r\n" + String.Join("\r\n", result.Failures.ToArray()),
                "部分完成", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        else
        {
            status.Text = "转换失败，请检查 PDF 文件后重试。";
            MessageBox.Show(this, String.Join("\r\n", result.Failures.ToArray()),
                "转换失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static string UniquePath(string path)
    {
        if (!File.Exists(path))
            return path;
        string directory = Path.GetDirectoryName(path);
        string stem = Path.GetFileNameWithoutExtension(path);
        string extension = Path.GetExtension(path);
        int index = 2;
        while (true)
        {
            string candidate = Path.Combine(directory, stem + "(" + index + ")" + extension);
            if (!File.Exists(candidate))
                return candidate;
            index++;
        }
    }

    private static string Quote(string value)
    {
        StringBuilder result = new StringBuilder("\"");
        int slashes = 0;
        foreach (char character in value)
        {
            if (character == '\\')
                slashes++;
            else if (character == '"')
            {
                result.Append('\\', slashes * 2 + 1).Append('"');
                slashes = 0;
            }
            else
            {
                result.Append('\\', slashes).Append(character);
                slashes = 0;
            }
        }
        result.Append('\\', slashes * 2).Append('"');
        return result.ToString();
    }

    private sealed class Job
    {
        internal readonly string[] Files;
        internal readonly string OutputFolder;
        internal readonly bool Visual;
        internal readonly string Dpi;
        internal Job(string[] files, string outputFolder, bool visual, string dpi)
        {
            Files = files; OutputFolder = outputFolder; Visual = visual; Dpi = dpi;
        }
    }

    private sealed class Result
    {
        internal readonly List<string> Completed = new List<string>();
        internal readonly List<string> Failures = new List<string>();
    }
}
