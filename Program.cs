using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace NovelProofreader
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ThreadException += delegate(object sender, System.Threading.ThreadExceptionEventArgs e)
            {
                LogException(e.Exception);
                MessageBox.Show("程序遇到错误：\r\n" + e.Exception.Message, "小说校对助手",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            AppDomain.CurrentDomain.UnhandledException += delegate(object sender, UnhandledExceptionEventArgs e)
            {
                Exception ex = e.ExceptionObject as Exception;
                if (ex != null) LogException(ex);
            };
            string initialFile = args != null && args.Length > 0 ? args[0] : null;
            Application.Run(new MainForm(initialFile));
        }

        private static void LogException(Exception ex)
        {
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "novelproofreader_error.log");
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("==== " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " ====");
                sb.AppendLine(ex.ToString());
                sb.AppendLine();
                File.AppendAllText(path, sb.ToString(), new UTF8Encoding(true));
            }
            catch { }
        }
    }
}

