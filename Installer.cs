using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace NovelProofreaderSetup
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool uninstall = HasArg(args, "--uninstall") ||
                string.Equals(Path.GetFileName(Application.ExecutablePath), InstallEngine.UninstallerName, StringComparison.OrdinalIgnoreCase);
            bool silent = HasArg(args, "--silent");

            if (uninstall)
            {
                RunUninstall(silent);
                return;
            }
            if (silent)
            {
                RunSilentInstall(args);
                return;
            }
            Application.Run(new SetupForm());
        }

        private static void RunSilentInstall(string[] args)
        {
            string dir = GetArgValue(args, "--dir");
            if (string.IsNullOrEmpty(dir)) dir = GetArgValue(args, "--install");
            if (string.IsNullOrEmpty(dir)) dir = InstallEngine.DefaultInstallDir;

            string error;
            List<string> warnings;
            bool ok = InstallEngine.Install(dir, HasArg(args, "--desktop"), HasArg(args, "--launch"), out error, out warnings);
            if (!ok)
            {
                InstallEngine.WriteLog("静默安装失败：" + error + Environment.NewLine);
                Environment.ExitCode = 1;
                return;
            }
            if (warnings.Count > 0)
                InstallEngine.WriteLog("安装完成，但有警告：" + Environment.NewLine + string.Join(Environment.NewLine, warnings.ToArray()) + Environment.NewLine);
            Environment.ExitCode = 0;
        }

        private static void RunUninstall(bool silent)
        {
            if (!silent)
            {
                DialogResult confirm = MessageBox.Show(
                    "确定要卸载“小说校对助手”吗？\r\n\r\n将删除程序文件、开始菜单/桌面快捷方式和安装目录中的配置文件。\r\n你的小说原稿和校对临时文件不会被删除。",
                    "卸载小说校对助手", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (confirm != DialogResult.Yes) return;
            }

            string error;
            List<string> warnings;
            bool ok = InstallEngine.Uninstall(out error, out warnings);
            if (!ok)
            {
                if (!silent) MessageBox.Show("卸载失败：\r\n" + error, "卸载小说校对助手", MessageBoxButtons.OK, MessageBoxIcon.Error);
                InstallEngine.WriteLog("卸载失败：" + error + Environment.NewLine);
                Environment.ExitCode = 1;
                return;
            }
            if (!silent)
            {
                string text = "卸载完成。";
                if (warnings.Count > 0) text += "\r\n\r\n部分文件未能立即删除，将在几秒后自动清理。";
                MessageBox.Show(text, "卸载小说校对助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Environment.ExitCode = 0;
        }

        private static bool HasArg(string[] args, string name)
        {
            if (args == null) return false;
            foreach (string arg in args)
            {
                if (string.Equals(arg, name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static string GetArgValue(string[] args, string name)
        {
            if (args == null) return null;
            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];
                if (arg.StartsWith(name + "=", StringComparison.OrdinalIgnoreCase))
                    return arg.Substring(name.Length + 1).Trim('"');
                if (string.Equals(arg, name, StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                    return args[i + 1].Trim('"');
            }
            return null;
        }
    }

    internal class SetupForm : Form
    {
        private TextBox txtPath;
        private CheckBox chkDesktop;
        private CheckBox chkLaunch;
        private Button btnInstall;
        private Button btnCancel;
        private Button btnBrowse;
        private ProgressBar progress;
        private Label lblStatus;
        private bool installing;

        public SetupForm()
        {
            this.Text = "小说校对助手 安装程序";
            this.ClientSize = new Size(620, 430);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(246, 248, 250);
            this.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
            this.AutoScaleMode = AutoScaleMode.Dpi;
            try { this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { }

            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 88;
            header.BackColor = Color.FromArgb(27, 79, 114);
            this.Controls.Add(header);

            PictureBox logo = new PictureBox();
            logo.Location = new Point(20, 16);
            logo.Size = new Size(56, 56);
            logo.SizeMode = PictureBoxSizeMode.StretchImage;
            try
            {
                Icon icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                logo.Image = icon.ToBitmap();
            }
            catch { }
            header.Controls.Add(logo);

            Label title = new Label();
            title.Text = "小说校对助手";
            title.ForeColor = Color.White;
            title.Font = new Font("Microsoft YaHei UI", 16F, FontStyle.Bold);
            title.AutoSize = true;
            title.Location = new Point(90, 14);
            header.Controls.Add(title);

            Label subtitle = new Label();
            subtitle.Text = "Word 错别字 / 词法校对工具  ·  v" + InstallEngine.AppVersion;
            subtitle.ForeColor = Color.FromArgb(214, 234, 248);
            subtitle.AutoSize = true;
            subtitle.Location = new Point(92, 50);
            header.Controls.Add(subtitle);

            Label pathLabel = new Label();
            pathLabel.Text = "安装位置";
            pathLabel.Location = new Point(24, 112);
            pathLabel.Size = new Size(70, 24);
            pathLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.Controls.Add(pathLabel);

            txtPath = new TextBox();
            txtPath.Location = new Point(96, 108);
            txtPath.Size = new Size(420, 26);
            txtPath.Text = InstallEngine.DefaultInstallDir;
            this.Controls.Add(txtPath);

            btnBrowse = new Button();
            btnBrowse.Text = "浏览…";
            btnBrowse.Location = new Point(524, 107);
            btnBrowse.Size = new Size(72, 28);
            btnBrowse.Click += delegate { BrowseFolder(); };
            this.Controls.Add(btnBrowse);

            Label hint = new Label();
            hint.Text = "安装到当前用户目录，不需要管理员权限。";
            hint.ForeColor = Color.FromArgb(127, 140, 141);
            hint.Location = new Point(96, 140);
            hint.Size = new Size(500, 22);
            this.Controls.Add(hint);

            chkDesktop = new CheckBox();
            chkDesktop.Text = "创建桌面快捷方式";
            chkDesktop.Checked = true;
            chkDesktop.AutoSize = true;
            chkDesktop.Location = new Point(96, 178);
            this.Controls.Add(chkDesktop);

            chkLaunch = new CheckBox();
            chkLaunch.Text = "安装完成后启动小说校对助手";
            chkLaunch.Checked = true;
            chkLaunch.AutoSize = true;
            chkLaunch.Location = new Point(96, 206);
            this.Controls.Add(chkLaunch);

            Label version = new Label();
            version.Text = "安装包版本：v" + InstallEngine.AppVersion + "　·　保留已有的配置文件（如果有）";
            version.ForeColor = Color.FromArgb(127, 140, 141);
            version.Location = new Point(24, 252);
            version.Size = new Size(560, 22);
            this.Controls.Add(version);

            progress = new ProgressBar();
            progress.Location = new Point(24, 300);
            progress.Size = new Size(420, 12);
            progress.Style = ProgressBarStyle.Marquee;
            progress.Visible = false;
            this.Controls.Add(progress);

            lblStatus = new Label();
            lblStatus.Text = "准备安装。";
            lblStatus.ForeColor = Color.FromArgb(86, 101, 115);
            lblStatus.Location = new Point(24, 320);
            lblStatus.Size = new Size(420, 50);
            this.Controls.Add(lblStatus);

            btnInstall = MakeButton("安装", true);
            btnInstall.Location = new Point(458, 360);
            btnInstall.Size = new Size(90, 34);
            btnInstall.Click += delegate { StartInstall(); };
            this.Controls.Add(btnInstall);

            btnCancel = MakeButton("取消", false);
            btnCancel.Location = new Point(556, 360);
            btnCancel.Size = new Size(60, 34);
            btnCancel.Click += delegate { Close(); };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnInstall;
            this.CancelButton = btnCancel;
            this.FormClosing += delegate(object sender, FormClosingEventArgs e)
            {
                if (installing) e.Cancel = true;
            };
        }

        private Button MakeButton(string text, bool primary)
        {
            Button button = new Button();
            button.Text = text;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = primary ? Color.FromArgb(31, 97, 141) : Color.FromArgb(189, 195, 199);
            button.BackColor = primary ? Color.FromArgb(41, 128, 185) : Color.FromArgb(250, 251, 252);
            button.ForeColor = primary ? Color.White : Color.FromArgb(44, 62, 80);
            button.UseVisualStyleBackColor = false;
            button.Cursor = Cursors.Hand;
            return button;
        }

        private void BrowseFolder()
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择小说校对助手的安装位置";
                dialog.ShowNewFolderButton = true;
                dialog.SelectedPath = txtPath.Text;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    txtPath.Text = dialog.SelectedPath;
            }
        }

        private void StartInstall()
        {
            if (installing) return;
            string dir = txtPath.Text.Trim();
            if (dir.Length == 0)
            {
                MessageBox.Show(this, "请选择安装位置。", "安装", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            installing = true;
            SetControlsEnabled(false);
            progress.Visible = true;
            progress.Style = ProgressBarStyle.Marquee;
            lblStatus.Text = "正在安装，请稍候…";
            this.Refresh();

            string error;
            List<string> warnings;
            bool ok = InstallEngine.Install(dir, chkDesktop.Checked, false, out error, out warnings);
            if (!ok)
            {
                installing = false;
                SetControlsEnabled(true);
                progress.Visible = false;
                lblStatus.Text = "安装失败。";
                MessageBox.Show(this, "安装失败：\r\n" + error, "安装失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            progress.Style = ProgressBarStyle.Blocks;
            progress.Value = 100;
            lblStatus.Text = "安装完成。";
            if (warnings.Count > 0)
            {
                MessageBox.Show(this,
                    "安装完成，但以下项目未能创建：\r\n\r\n" + string.Join("\r\n", warnings.ToArray()),
                    "安装完成", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show(this,
                    "安装完成！\r\n\r\n安装位置：\r\n" + Path.GetFullPath(dir) +
                    "\r\n\r\n配置文件：\r\n" + Path.Combine(Path.GetFullPath(dir), InstallEngine.ConfigName) +
                    "\r\n\r\n可以在开始菜单或桌面找到“小说校对助手”。",
                    "安装完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (chkLaunch.Checked)
            {
                try
                {
                    string appPath = Path.Combine(Path.GetFullPath(dir), InstallEngine.AppExeName);
                    ProcessStartInfo info = new ProcessStartInfo(appPath);
                    info.UseShellExecute = true;
                    info.WorkingDirectory = Path.GetDirectoryName(appPath);
                    Process.Start(info);
                }
                catch { }
            }

            installing = false;
            Close();
        }

        private void SetControlsEnabled(bool enabled)
        {
            txtPath.Enabled = enabled;
            btnBrowse.Enabled = enabled;
            chkDesktop.Enabled = enabled;
            chkLaunch.Enabled = enabled;
            btnInstall.Enabled = enabled;
            btnCancel.Enabled = enabled;
        }
    }

    internal static class InstallEngine
    {
        public const string AppName = "小说校对助手";
        public const string AppVersion = "1.1.2";
        public const string AppExeName = "小说校对助手.exe";
        public const string ReadmeName = "使用说明.txt";
        public const string ConfigName = "小说校对助手.ini";
        public const string UninstallerName = "卸载.exe";
        public const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\小说校对助手";
        private const string PayloadApp = "Payload.App";
        private const string PayloadReadme = "Payload.Readme";
        private const string PayloadConfig = "Payload.Config";

        public static string DefaultInstallDir
        {
            get
            {
                try
                {
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(UninstallKey))
                    {
                        if (key != null)
                        {
                            string dir = key.GetValue("InstallLocation") as string;
                            if (!string.IsNullOrEmpty(dir)) return dir;
                        }
                    }
                }
                catch { }
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Programs", AppName);
            }
        }

        public static bool Install(string dir, bool createDesktopShortcut, bool launch, out string error, out List<string> warnings)
        {
            error = null;
            warnings = new List<string>();
            try
            {
                dir = NormalizeDirectory(dir);
                if (IsAppRunning())
                {
                    error = "小说校对助手正在运行，请先关闭程序再安装。";
                    return false;
                }

                Directory.CreateDirectory(dir);
                string appPath = Path.Combine(dir, AppExeName);
                string readmePath = Path.Combine(dir, ReadmeName);
                string configPath = Path.Combine(dir, ConfigName);

                WriteResource(PayloadApp, appPath);
                WriteResource(PayloadReadme, readmePath);
                // Never overwrite an existing config: it may contain the user's API key.
                if (!File.Exists(configPath)) WriteResource(PayloadConfig, configPath);

                string uninstallerPath = Path.Combine(dir, UninstallerName);
                if (!PathsEqual(Application.ExecutablePath, uninstallerPath))
                    File.Copy(Application.ExecutablePath, uninstallerPath, true);

                string programs = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
                string startMenuDir = Path.Combine(programs, AppName);
                Directory.CreateDirectory(startMenuDir);
                TryCreateShortcut(Path.Combine(startMenuDir, AppName + ".lnk"), appPath, string.Empty, dir, appPath,
                    AppName, warnings);
                TryCreateShortcut(Path.Combine(startMenuDir, "卸载.lnk"), uninstallerPath, "--uninstall", dir, appPath,
                    "卸载" + AppName, warnings);

                string desktopLink = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), AppName + ".lnk");
                if (createDesktopShortcut)
                    TryCreateShortcut(desktopLink, appPath, string.Empty, dir, appPath, AppName, warnings);
                else
                    TryDeleteFile(desktopLink);

                WriteUninstallRegistry(dir, appPath, uninstallerPath);
                if (launch) Launch(appPath);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public static bool Uninstall(out string error, out List<string> warnings)
        {
            error = null;
            warnings = new List<string>();
            string dir = null;
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(UninstallKey))
                {
                    if (key != null) dir = key.GetValue("InstallLocation") as string;
                }
                if (string.IsNullOrEmpty(dir)) dir = Path.GetDirectoryName(Application.ExecutablePath);
                if (string.IsNullOrEmpty(dir)) dir = Path.GetDirectoryName(Application.ExecutablePath);

                if (IsAppRunning())
                {
                    error = "小说校对助手正在运行，请先关闭程序再卸载。";
                    return false;
                }

                string programs = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
                string startMenuDir = Path.Combine(programs, AppName);
                TryDeleteDirectory(startMenuDir, warnings);
                TryDeleteFile(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), AppName + ".lnk"));

                try
                {
                    Registry.CurrentUser.DeleteSubKeyTree(UninstallKey, false);
                }
                catch (Exception ex) { warnings.Add("删除卸载注册表项失败：" + ex.Message); }

                // Delete the config first so an API key does not remain if folder cleanup is delayed.
                TryDeleteFile(Path.Combine(dir, ConfigName));
                TryDeleteFile(Path.Combine(dir, AppExeName));
                TryDeleteFile(Path.Combine(dir, ReadmeName));

                ScheduleFolderDelete(dir);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public static void Launch(string appPath)
        {
            ProcessStartInfo info = new ProcessStartInfo(appPath);
            info.UseShellExecute = true;
            info.WorkingDirectory = Path.GetDirectoryName(appPath);
            Process.Start(info);
        }

        public static void WriteLog(string text)
        {
            try
            {
                File.AppendAllText(Path.Combine(Path.GetTempPath(), "小说校对助手_安装日志.txt"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + text, new UTF8Encoding(true));
            }
            catch { }
        }

        private static string NormalizeDirectory(string dir)
        {
            if (string.IsNullOrEmpty(dir)) throw new InvalidOperationException("安装位置不能为空。");
            dir = Environment.ExpandEnvironmentVariables(dir.Trim().Trim('"'));
            dir = Path.GetFullPath(dir).TrimEnd('\\');
            if (dir.Length <= 2) throw new InvalidOperationException("不能直接安装到磁盘根目录，请选择一个子目录。");
            if (File.Exists(dir)) throw new InvalidOperationException("安装位置是一个文件，请选择一个文件夹。");
            return dir;
        }

        private static void WriteResource(string name, string path)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (stream == null) throw new InvalidOperationException("安装包缺少资源：" + name);
                using (FileStream output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    stream.CopyTo(output);
                }
            }
        }

        private static void WriteUninstallRegistry(string dir, string appPath, string uninstallerPath)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(UninstallKey))
            {
                if (key == null) throw new InvalidOperationException("无法写入卸载注册表项。");
                key.SetValue("DisplayName", AppName);
                key.SetValue("DisplayVersion", AppVersion);
                key.SetValue("Publisher", AppName);
                key.SetValue("InstallLocation", dir);
                key.SetValue("DisplayIcon", appPath);
                key.SetValue("UninstallString", "\"" + uninstallerPath + "\" --uninstall");
                key.SetValue("QuietUninstallString", "\"" + uninstallerPath + "\" --uninstall --silent");
                key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
                key.SetValue("EstimatedSize", GetDirectorySizeKb(dir), RegistryValueKind.DWord);
            }
        }

        private static int GetDirectorySizeKb(string dir)
        {
            try
            {
                long total = 0;
                foreach (string file in Directory.GetFiles(dir, "*", SearchOption.AllDirectories))
                {
                    try { total += new FileInfo(file).Length; }
                    catch { }
                }
                return (int)Math.Max(1, total / 1024);
            }
            catch { return 1; }
        }

        private static bool IsAppRunning()
        {
            try
            {
                return Process.GetProcessesByName(Path.GetFileNameWithoutExtension(AppExeName)).Length > 0;
            }
            catch { return false; }
        }

        private static void TryCreateShortcut(string linkPath, string targetPath, string arguments, string workingDirectory,
            string iconPath, string description, List<string> warnings)
        {
            object shellObject = null;
            object shortcutObject = null;
            try
            {
                Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                if (shellType == null) throw new InvalidOperationException("系统不支持创建快捷方式。");
                shellObject = Activator.CreateInstance(shellType);
                dynamic shell = shellObject;
                shortcutObject = shell.CreateShortcut(linkPath);
                dynamic shortcut = shortcutObject;
                shortcut.TargetPath = targetPath;
                shortcut.Arguments = arguments ?? string.Empty;
                shortcut.WorkingDirectory = workingDirectory;
                shortcut.IconLocation = iconPath + ",0";
                shortcut.Description = description;
                shortcut.Save();
            }
            catch (Exception ex)
            {
                warnings.Add(Path.GetFileName(linkPath) + "：" + ex.Message);
            }
            finally
            {
                if (shortcutObject != null) { try { Marshal.ReleaseComObject(shortcutObject); } catch { } }
                if (shellObject != null) { try { Marshal.ReleaseComObject(shellObject); } catch { } }
            }
        }

        private static void ScheduleFolderDelete(string dir)
        {
            try
            {
                string batch = Path.Combine(Path.GetTempPath(), "novelproofreader_uninstall_" + Guid.NewGuid().ToString("N") + ".bat");
                string safeDir = dir.Replace("%", "%%");
                string content =
                    "@echo off\r\n" +
                    "ping 127.0.0.1 -n 3 >nul\r\n" +
                    "rmdir /s /q \"" + safeDir + "\"\r\n" +
                    "del \"%~f0\"\r\n";
                File.WriteAllText(batch, content, Encoding.Default);
                ProcessStartInfo info = new ProcessStartInfo("cmd.exe", "/c \"" + batch + "\"");
                info.UseShellExecute = false;
                info.CreateNoWindow = true;
                info.WindowStyle = ProcessWindowStyle.Hidden;
                Process.Start(info);
            }
            catch
            {
                try { Directory.Delete(dir, true); }
                catch { }
            }
        }

        private static void TryDeleteDirectory(string dir, List<string> warnings)
        {
            try { if (Directory.Exists(dir)) Directory.Delete(dir, true); }
            catch (Exception ex) { warnings.Add("删除目录失败：" + ex.Message); }
        }

        private static void TryDeleteFile(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch { }
        }

        private static bool PathsEqual(string a, string b)
        {
            try
            {
                return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);
            }
            catch { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }
        }
    }
}
