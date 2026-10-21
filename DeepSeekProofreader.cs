using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace NovelProofreader
{
    public class OnlineProofreadResult
    {
        public bool Success;
        public string Error = string.Empty;
        public List<Issue> Issues = new List<Issue>();
        public int Chunks;
        public int CharactersSent;
        public int RawIssueCount;
    }

    public static class DeepSeekProofreader
    {
        private const string SystemPrompt =
            "你是一名严谨的中文小说校对编辑。请找出文本中的错别字、词语误用、成语误写、标点错误、明显语法或搭配错误，以及人名、专名前后不一致。\r\n" +
            "只输出 JSON，不要 Markdown，不要解释。JSON 格式：\r\n" +
            "{\"issues\":[{\"paragraph\":22,\"original\":\"错词\",\"suggestion\":\"正确词\",\"type\":\"错别字\",\"note\":\"简短说明\"}]}\r\n" +
            "要求：\r\n" +
            "1. paragraph 是文本中【P数字】里的数字，必须原样返回。\r\n" +
            "2. original 必须逐字摘自该段，不能改写、不能跨段；suggestion 是替换后的文字。\r\n" +
            "3. 只报告有把握的错误；不要做风格润色、不要改写句子、不要建议增删整段。\r\n" +
            "4. 小说中的人名、地名、外来语、口语、方言、作者有意为之的用词不要当作错字。\r\n" +
            "5. type 只能是：错别字、词法、标点、名称、冗余、用词建议。\r\n" +
            "6. 没有问题时返回 {\"issues\":[]}。";

        static DeepSeekProofreader()
        {
            try
            {
                // DeepSeek requires TLS 1.2. Numeric value keeps compatibility with old reference assemblies.
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
            }
            catch { }
        }

        public static OnlineProofreadResult Proofread(List<string> paragraphs, DeepSeekConfig config, Action<int, int, string> progress)
        {
            OnlineProofreadResult result = new OnlineProofreadResult();
            if (config == null || !config.HasKey)
            {
                result.Error = "未配置 DeepSeek API Key。";
                return result;
            }
            if (string.IsNullOrEmpty(config.ApiBase))
            {
                result.Error = "未配置 ApiBase。";
                return result;
            }
            if (paragraphs == null) paragraphs = new List<string>();

            List<List<int>> chunks = BuildChunks(paragraphs, config.ChunkSize);
            result.Chunks = chunks.Count;
            if (chunks.Count == 0)
            {
                result.Success = true;
                return result;
            }

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;

            for (int c = 0; c < chunks.Count; c++)
            {
                if (progress != null) progress(c + 1, chunks.Count, "正在请求 DeepSeek（" + (c + 1) + "/" + chunks.Count + "）…");

                string userText = BuildUserText(paragraphs, chunks[c]);
                result.CharactersSent += userText.Length;

                string error;
                string content = CallApi(config, userText, serializer, out error);
                if (content == null)
                {
                    result.Success = false;
                    result.Error = string.IsNullOrEmpty(error) ? "DeepSeek 请求失败。" : error;
                    return result;
                }

                List<Issue> chunkIssues;
                string parseError;
                List<Issue> parsed = ParseContent(content, paragraphs, chunks[c], out parseError, out chunkIssues);
                result.RawIssueCount += chunkIssues == null ? 0 : chunkIssues.Count;
                if (parsed == null)
                {
                    result.Success = false;
                    result.Error = parseError;
                    return result;
                }
                result.Issues.AddRange(parsed);
            }

            result.Success = true;
            return result;
        }

        private static List<List<int>> BuildChunks(List<string> paragraphs, int chunkSize)
        {
            if (chunkSize < 1000) chunkSize = 1000;
            List<List<int>> chunks = new List<List<int>>();
            List<int> current = new List<int>();
            int currentChars = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                string text = paragraphs[i] ?? string.Empty;
                if (text.Length == 0) continue;
                int add = text.Length + 24;
                if (current.Count > 0 && currentChars + add > chunkSize)
                {
                    chunks.Add(current);
                    current = new List<int>();
                    currentChars = 0;
                }
                current.Add(i);
                currentChars += add;
            }
            if (current.Count > 0) chunks.Add(current);
            return chunks;
        }

        private static string BuildUserText(List<string> paragraphs, List<int> chunk)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("请校对下面的小说片段：\r\n\r\n");
            foreach (int index in chunk)
            {
                sb.Append("【P").Append(index + 1).Append("】");
                sb.Append(paragraphs[index] ?? string.Empty);
                sb.Append("\r\n");
            }
            return sb.ToString();
        }

        private static string CallApi(DeepSeekConfig config, string userText, JavaScriptSerializer serializer, out string error)
        {
            error = null;
            try
            {
                string endpoint = config.ApiBase.Trim();
                if (endpoint.Length == 0) throw new InvalidOperationException("ApiBase 为空。");
                if (!endpoint.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                    !endpoint.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("ApiBase 必须以 http:// 或 https:// 开头。");
                endpoint = endpoint.TrimEnd('/');
                if (!endpoint.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
                    endpoint += "/chat/completions";

                List<Dictionary<string, object>> messages = new List<Dictionary<string, object>>();
                Dictionary<string, object> systemMessage = new Dictionary<string, object>();
                systemMessage["role"] = "system";
                systemMessage["content"] = SystemPrompt;
                messages.Add(systemMessage);
                Dictionary<string, object> userMessage = new Dictionary<string, object>();
                userMessage["role"] = "user";
                userMessage["content"] = userText;
                messages.Add(userMessage);

                Dictionary<string, object> payload = new Dictionary<string, object>();
                payload["model"] = string.IsNullOrEmpty(config.Model) ? "deepseek-chat" : config.Model;
                payload["messages"] = messages;
                payload["temperature"] = 0.1;
                payload["stream"] = false;
                payload["max_tokens"] = 8192;

                string json = serializer.Serialize(payload);
                byte[] bytes = Encoding.UTF8.GetBytes(json);

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(endpoint);
                request.Method = "POST";
                request.ContentType = "application/json; charset=utf-8";
                request.Accept = "application/json";
                request.Headers["Authorization"] = "Bearer " + config.ApiKey.Trim();
                request.UserAgent = "NovelProofreader/1.1";
                request.KeepAlive = false;
                request.ContentLength = bytes.Length;
                int timeout = Math.Max(10, config.TimeoutSeconds) * 1000;
                request.Timeout = timeout;
                request.ReadWriteTimeout = timeout;

                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(bytes, 0, bytes.Length);
                }

                string responseText;
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    responseText = reader.ReadToEnd();
                }
                return ExtractContent(responseText, serializer, out error);
            }
            catch (WebException ex)
            {
                string detail = ex.Message;
                try
                {
                    if (ex.Response != null)
                    {
                        using (Stream stream = ex.Response.GetResponseStream())
                        using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                        {
                            string body = reader.ReadToEnd();
                            HttpWebResponse response = ex.Response as HttpWebResponse;
                            if (response != null)
                                detail = "HTTP " + (int)response.StatusCode + " " + response.StatusDescription + " " + body;
                            else
                                detail = body;
                        }
                    }
                }
                catch { }
                error = Shorten(detail, 800);
                return null;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return null;
            }
        }

        private static string ExtractContent(string responseText, JavaScriptSerializer serializer, out string error)
        {
            error = null;
            try
            {
                object parsed = serializer.DeserializeObject(responseText);
                Dictionary<string, object> root = parsed as Dictionary<string, object>;
                if (root == null)
                {
                    error = "DeepSeek 返回的内容不是 JSON 对象。";
                    return null;
                }
                object choicesObject;
                if (!root.TryGetValue("choices", out choicesObject))
                {
                    error = "DeepSeek 返回中没有 choices：" + Shorten(responseText, 500);
                    return null;
                }
                IList choices = choicesObject as IList;
                if (choices == null || choices.Count == 0)
                {
                    error = "DeepSeek 返回的 choices 为空。";
                    return null;
                }
                Dictionary<string, object> choice = choices[0] as Dictionary<string, object>;
                if (choice == null)
                {
                    error = "DeepSeek 返回的 choice 格式异常。";
                    return null;
                }
                Dictionary<string, object> message = GetDictionary(choice, "message");
                if (message == null)
                {
                    error = "DeepSeek 返回的 message 为空。";
                    return null;
                }
                string content = GetString(message, "content");
                if (string.IsNullOrEmpty(content))
                {
                    error = "DeepSeek 返回的 message.content 为空。";
                    return null;
                }
                return content;
            }
            catch (Exception ex)
            {
                error = "解析 DeepSeek 响应失败：" + ex.Message;
                return null;
            }
        }

        private static List<Issue> ParseContent(string raw, List<string> paragraphs, List<int> chunk, out string error, out List<Issue> rawIssues)
        {
            error = null;
            rawIssues = new List<Issue>();
            List<Issue> issues = new List<Issue>();
            try
            {
                string content = raw.Trim();
                if (content.StartsWith("```"))
                {
                    int newline = content.IndexOf('\n');
                    if (newline >= 0) content = content.Substring(newline + 1);
                    if (content.EndsWith("```")) content = content.Substring(0, content.Length - 3);
                    content = content.Trim();
                }
                int firstBrace = content.IndexOf('{');
                int lastBrace = content.LastIndexOf('}');
                if (firstBrace < 0 || lastBrace <= firstBrace)
                {
                    error = "DeepSeek 没有返回 JSON：" + Shorten(content, 300);
                    return null;
                }
                content = content.Substring(firstBrace, lastBrace - firstBrace + 1);

                JavaScriptSerializer serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = int.MaxValue;
                object parsed = serializer.DeserializeObject(content);
                Dictionary<string, object> root = parsed as Dictionary<string, object>;
                if (root == null)
                {
                    error = "DeepSeek 返回的 JSON 不是对象。";
                    return null;
                }
                object issuesObject;
                if (!root.TryGetValue("issues", out issuesObject))
                {
                    error = "DeepSeek 返回的 JSON 缺少 issues 字段。";
                    return null;
                }
                IList list = issuesObject as IList;
                if (list == null)
                {
                    error = "DeepSeek 返回的 issues 不是数组。";
                    return null;
                }

                foreach (object item in list)
                {
                    Dictionary<string, object> dict = item as Dictionary<string, object>;
                    if (dict == null) continue;
                    int paragraphNumber;
                    if (!TryGetInt(dict, "paragraph", out paragraphNumber)) continue;
                    int paragraph = paragraphNumber - 1;
                    if (paragraph < 0 || paragraph >= paragraphs.Count) continue;
                    if (!chunk.Contains(paragraph)) continue;

                    string original = GetString(dict, "original");
                    string suggestion = GetString(dict, "suggestion");
                    if (string.IsNullOrEmpty(original) || original.Length == 0) continue;
                    if (string.IsNullOrEmpty(suggestion)) continue;
                    suggestion = suggestion.Replace("\r", " ").Replace("\n", " ").Trim();
                    if (suggestion.Length == 0) continue;

                    string text = paragraphs[paragraph] ?? string.Empty;
                    int start = FindFreeOccurrence(text, original, issues, paragraph);
                    if (start < 0)
                    {
                        string trimmed = original.Trim();
                        if (trimmed.Length > 0 && trimmed != original)
                        {
                            original = trimmed;
                            start = FindFreeOccurrence(text, original, issues, paragraph);
                        }
                    }
                    if (start < 0) continue;

                    bool duplicate = false;
                    foreach (Issue existing in issues)
                    {
                        if (existing.Paragraph == paragraph && existing.Start == start && existing.Length == original.Length)
                        {
                            duplicate = true;
                            break;
                        }
                    }
                    if (duplicate) continue;

                    string type = NormalizeType(GetString(dict, "type"));
                    string note = GetString(dict, "note");
                    if (string.IsNullOrEmpty(note)) note = "在线校对建议，请结合上下文判断。";

                    Issue issue = new Issue();
                    issue.Paragraph = paragraph;
                    issue.Start = start;
                    issue.Length = original.Length;
                    issue.Type = type;
                    issue.Original = original;
                    issue.Suggestion = suggestion;
                    issue.Note = "DeepSeek：" + note;
                    issue.Priority = 50;
                    issues.Add(issue);
                    rawIssues.Add(issue);
                }
                return issues;
            }
            catch (Exception ex)
            {
                error = "解析 DeepSeek 校对结果失败：" + ex.Message;
                return null;
            }
        }

        private static int FindFreeOccurrence(string text, string original, List<Issue> existing, int paragraph)
        {
            int search = 0;
            while (search <= text.Length - original.Length)
            {
                int start = text.IndexOf(original, search, StringComparison.Ordinal);
                if (start < 0) return -1;
                bool overlap = false;
                for (int i = 0; i < existing.Count; i++)
                {
                    Issue e = existing[i];
                    if (e.Paragraph != paragraph) continue;
                    if (start < e.Start + e.Length && e.Start < start + original.Length)
                    {
                        overlap = true;
                        break;
                    }
                }
                if (!overlap) return start;
                search = start + 1;
            }
            return -1;
        }

        private static string NormalizeType(string type)
        {
            if (string.IsNullOrEmpty(type)) return Scanner.Lex;
            string t = type.Trim();
            if (t.Contains("错别") || t.Contains("错字") || t.Contains("别字")) return Scanner.Typo;
            if (t.Contains("标点") || t.Contains("格式")) return Scanner.Punctuation;
            if (t.Contains("名称") || t.Contains("人名") || t.Contains("专名")) return Scanner.Name;
            if (t.Contains("冗余") || t.Contains("重复") || t.Contains("赘余")) return Scanner.Redundant;
            if (t.Contains("用词") || t.Contains("建议") || t.Contains("风格")) return Scanner.Style;
            return Scanner.Lex;
        }

        private static Dictionary<string, object> GetDictionary(Dictionary<string, object> dict, string key)
        {
            object value;
            if (!dict.TryGetValue(key, out value)) return null;
            return value as Dictionary<string, object>;
        }

        private static string GetString(Dictionary<string, object> dict, string key)
        {
            object value;
            if (!dict.TryGetValue(key, out value) || value == null) return string.Empty;
            return value.ToString();
        }

        private static bool TryGetInt(Dictionary<string, object> dict, string key, out int value)
        {
            value = 0;
            object raw;
            if (!dict.TryGetValue(key, out raw) || raw == null) return false;
            if (raw is int) { value = (int)raw; return true; }
            if (raw is long) { value = (int)(long)raw; return true; }
            if (raw is double) { value = (int)(double)raw; return true; }
            string text = raw.ToString().Trim();
            if (text.StartsWith("P", StringComparison.OrdinalIgnoreCase)) text = text.Substring(1);
            return int.TryParse(text, out value);
        }

        private static string Shorten(string value, int max)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            string s = value.Replace("\r", " ").Replace("\n", " ");
            if (s.Length <= max) return s;
            return s.Substring(0, max) + "…";
        }
    }
}
