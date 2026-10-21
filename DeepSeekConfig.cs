using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NovelProofreader
{
    public class DeepSeekConfig
    {
        public string ConfigPath = string.Empty;
        public bool Enabled = false;
        public string ApiKey = string.Empty;
        public string ApiBase = "https://api.deepseek.com";
        public string Model = "deepseek-chat";
        public int TimeoutSeconds = 120;
        public int ChunkSize = 6000;

        public bool HasKey
        {
            get { return !string.IsNullOrEmpty(ApiKey) && ApiKey.Trim().Length > 0; }
        }

        public static DeepSeekConfig LoadOrCreate()
        {
            string path = EnsureConfigFile();
            DeepSeekConfig config = new DeepSeekConfig();
            config.ConfigPath = path;
            if (!File.Exists(path)) return config;

            string section = string.Empty;
            string[] lines;
            try { lines = File.ReadAllLines(path, Encoding.UTF8); }
            catch { return config; }

            foreach (string raw in lines)
            {
                string line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";")) continue;
                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                    section = line.Substring(1, line.Length - 2).Trim();
                    continue;
                }
                if (section.Length > 0 && !string.Equals(section, "DeepSeek", StringComparison.OrdinalIgnoreCase)) continue;

                int eq = line.IndexOf('=');
                if (eq <= 0) continue;
                string key = line.Substring(0, eq).Trim();
                string value = line.Substring(eq + 1).Trim();
                if (value.Length >= 2 && ((value[0] == '"' && value[value.Length - 1] == '"') || (value[0] == '\'' && value[value.Length - 1] == '\'')))
                    value = value.Substring(1, value.Length - 2);

                if (string.Equals(key, "Enabled", StringComparison.OrdinalIgnoreCase))
                    config.Enabled = ParseBool(value);
                else if (string.Equals(key, "ApiKey", StringComparison.OrdinalIgnoreCase) || string.Equals(key, "Key", StringComparison.OrdinalIgnoreCase))
                    config.ApiKey = value;
                else if (string.Equals(key, "ApiBase", StringComparison.OrdinalIgnoreCase) || string.Equals(key, "BaseUrl", StringComparison.OrdinalIgnoreCase))
                    config.ApiBase = string.IsNullOrEmpty(value) ? "https://api.deepseek.com" : value;
                else if (string.Equals(key, "Model", StringComparison.OrdinalIgnoreCase))
                    config.Model = string.IsNullOrEmpty(value) ? "deepseek-chat" : value;
                else if (string.Equals(key, "TimeoutSeconds", StringComparison.OrdinalIgnoreCase))
                {
                    int n;
                    if (int.TryParse(value, out n)) config.TimeoutSeconds = Math.Max(10, Math.Min(600, n));
                }
                else if (string.Equals(key, "ChunkSize", StringComparison.OrdinalIgnoreCase))
                {
                    int n;
                    if (int.TryParse(value, out n)) config.ChunkSize = Math.Max(1000, Math.Min(20000, n));
                }
            }
            return config;
        }

        public static string EnsureConfigFile()
        {
            string local = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "小说校对助手.ini");
            string roaming = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "小说校对助手", "config.ini");

            if (File.Exists(local)) return local;
            if (File.Exists(roaming)) return roaming;

            try
            {
                WriteTemplate(local);
                return local;
            }
            catch
            {
                try
                {
                    string dir = Path.GetDirectoryName(roaming);
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    WriteTemplate(roaming);
                    return roaming;
                }
                catch
                {
                    return local;
                }
            }
        }

        private static void WriteTemplate(string path)
        {
            string content =
                "# 小说校对助手 DeepSeek 在线校对配置\r\n" +
                "# 填好 ApiKey 后，把 Enabled 改成 true，重新打开程序即可默认勾选在线校对。\r\n" +
                "# 也可以在程序里手动勾选【DeepSeek 在线校对】。\r\n" +
                "# 注意：ApiKey 以明文保存在本文件中，请勿把本文件发给别人。\r\n" +
                "\r\n" +
                "[DeepSeek]\r\n" +
                "Enabled=false\r\n" +
                "ApiKey=\r\n" +
                "ApiBase=https://api.deepseek.com\r\n" +
                "Model=deepseek-chat\r\n" +
                "TimeoutSeconds=120\r\n" +
                "ChunkSize=6000\r\n";
            File.WriteAllText(path, content, new UTF8Encoding(true));
        }

        private static bool ParseBool(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            string v = value.Trim().ToLowerInvariant();
            return v == "true" || v == "1" || v == "yes" || v == "on" || v == "是";
        }
    }
}
