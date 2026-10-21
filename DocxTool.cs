using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace NovelProofreader
{
    public class MergeResult
    {
        public bool Success;
        public string Message = string.Empty;
        public string BackupPath = string.Empty;
        public bool TempDeleted;
    }

    public static class DocxTool
    {
        private static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly XNamespace xml = XNamespace.Xml;

        private const string DocumentPart = "word/document.xml";
        private const string CommentsPart = "word/comments.xml";
        private const string DocumentRelsPart = "word/_rels/document.xml.rels";
        private const string ContentTypesPart = "[Content_Types].xml";
        private const string StylesPart = "word/styles.xml";
        private const string CommentsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
        private const string CommentsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

        public static bool IsValidDocx(string path)
        {
            try
            {
                Dictionary<string, byte[]> entries = ReadEntries(path);
                return entries.ContainsKey(DocumentPart);
            }
            catch
            {
                return false;
            }
        }

        public static List<string> ReadParagraphs(string path)
        {
            Dictionary<string, byte[]> entries = ReadEntries(path);
            byte[] docBytes;
            if (!entries.TryGetValue(DocumentPart, out docBytes))
                throw new InvalidDataException("文件中缺少 word/document.xml，可能不是有效的 .docx 文档。");

            XDocument doc = ParseXml(docBytes);
            List<string> result = new List<string>();
            foreach (XElement paragraph in doc.Descendants(w + "p"))
            {
                result.Add(GetParagraphText(paragraph));
            }
            return result;
        }

        public static int Annotate(string sourcePath, string tempPath, List<Issue> issues, out string error)
        {
            error = null;
            try
            {
                Dictionary<string, byte[]> entries = ReadEntries(sourcePath);
                byte[] docBytes;
                if (!entries.TryGetValue(DocumentPart, out docBytes))
                    throw new InvalidDataException("源文件缺少 word/document.xml，可能不是有效的 .docx 文档。");

                XDocument doc = ParseXml(docBytes);
                XDocument comments = null;
                byte[] commentsBytes;
                if (entries.TryGetValue(CommentsPart, out commentsBytes) && commentsBytes.Length > 0)
                    comments = ParseXml(commentsBytes);
                else
                    comments = new XDocument(new XElement(w + "comments", new XAttribute(XNamespace.Xmlns + "w", w.NamespaceName)));
                if (comments.Root == null)
                    comments = new XDocument(new XElement(w + "comments", new XAttribute(XNamespace.Xmlns + "w", w.NamespaceName)));

                bool hasCommentStyle = false;
                byte[] stylesBytes;
                if (entries.TryGetValue(StylesPart, out stylesBytes) && stylesBytes != null)
                {
                    string styles = Encoding.UTF8.GetString(stylesBytes);
                    hasCommentStyle = styles.IndexOf("CommentReference", StringComparison.Ordinal) >= 0;
                }

                List<XElement> paragraphs = doc.Descendants(w + "p").ToList();
                int nextId = MaxCommentId(comments) + 1;
                int annotated = 0;

                for (int i = 0; i < issues.Count; i++)
                {
                    Issue issue = issues[i];
                    if (issue.Paragraph < 0 || issue.Paragraph >= paragraphs.Count) continue;
                    if (issue.Length <= 0 || issue.Start < 0) continue;
                    string text = GetParagraphText(paragraphs[issue.Paragraph]);
                    if (issue.Start + issue.Length > text.Length) continue;
                    if (!AnnotateOne(paragraphs[issue.Paragraph], issue, nextId, hasCommentStyle)) continue;

                    comments.Root.Add(MakeComment(issue, nextId, hasCommentStyle));
                    nextId++;
                    annotated++;
                }

                if (annotated == 0)
                {
                    CopyFileShared(sourcePath, tempPath);
                    return 0;
                }

                byte[] newDoc = SaveXml(doc);
                byte[] newComments = SaveXml(comments);
                byte[] relsBytes;
                if (!entries.TryGetValue(DocumentRelsPart, out relsBytes))
                    relsBytes = Encoding.UTF8.GetBytes("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\" />");
                byte[] newRels = EnsureCommentsRelationship(relsBytes);
                byte[] typesBytes;
                if (!entries.TryGetValue(ContentTypesPart, out typesBytes))
                    typesBytes = Encoding.UTF8.GetBytes("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\" />");
                byte[] newTypes = EnsureCommentsContentType(typesBytes);

                CopyFileShared(sourcePath, tempPath);
                using (ZipArchive zip = ZipFile.Open(tempPath, ZipArchiveMode.Update))
                {
                    ReplaceEntry(zip, DocumentPart, newDoc);
                    ReplaceEntry(zip, CommentsPart, newComments);
                    ReplaceEntry(zip, DocumentRelsPart, newRels);
                    ReplaceEntry(zip, ContentTypesPart, newTypes);
                }
                return annotated;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                try { if (File.Exists(tempPath)) File.Delete(tempPath); }
                catch { }
                return -1;
            }
        }

        public static MergeResult Merge(string tempPath, string sourcePath, bool makeBackup)
        {
            MergeResult result = new MergeResult();
            string cleanPath = tempPath + ".clean";
            try
            {
                if (!File.Exists(tempPath))
                {
                    result.Message = "临时文件不存在，可能已被删除或移动。";
                    return result;
                }

                Dictionary<string, byte[]> entries = ReadEntries(tempPath);
                byte[] docBytes;
                if (!entries.TryGetValue(DocumentPart, out docBytes))
                {
                    result.Message = "临时文件缺少正文内容，无法写回。";
                    return result;
                }

                XDocument doc = ParseXml(docBytes);
                XDocument comments = null;
                bool hadCommentsPart = false;
                byte[] commentsBytes;
                if (entries.TryGetValue(CommentsPart, out commentsBytes) && commentsBytes.Length > 0)
                {
                    comments = ParseXml(commentsBytes);
                    hadCommentsPart = true;
                }

                HashSet<int> ourIds = new HashSet<int>();
                if (comments != null && comments.Root != null)
                {
                    List<XElement> ourComments = new List<XElement>();
                    foreach (XElement comment in comments.Root.Elements(w + "comment"))
                    {
                        string author = (string)comment.Attribute(w + "author");
                        if (string.Equals(author, Scanner.Author, StringComparison.Ordinal))
                        {
                            int id;
                            if (TryGetId(comment, out id)) ourIds.Add(id);
                            ourComments.Add(comment);
                        }
                    }
                    foreach (XElement comment in ourComments) comment.Remove();
                }

                List<XElement> removeNodes = new List<XElement>();
                List<XElement> emptyRuns = new List<XElement>();
                foreach (XElement element in doc.Descendants().ToList())
                {
                    if (element.Name == w + "commentRangeStart" || element.Name == w + "commentRangeEnd" || element.Name == w + "commentReference")
                    {
                        int id;
                        bool remove = false;
                        if (TryGetId(element, out id))
                        {
                            remove = ourIds.Contains(id) || !hadCommentsPart;
                        }
                        else
                        {
                            remove = !hadCommentsPart;
                        }
                        if (remove)
                        {
                            if (element.Name == w + "commentReference" && element.Parent != null && element.Parent.Name == w + "r")
                                emptyRuns.Add(element.Parent);
                            removeNodes.Add(element);
                        }
                    }
                }
                foreach (XElement element in removeNodes) element.Remove();
                foreach (XElement run in emptyRuns)
                {
                    if (run.Parent == null) continue;
                    if (!run.Elements().Any(e => e.Name != w + "rPr")) run.Remove();
                }

                // Remove the pale-yellow marker shading added by this tool, leaving other shading intact.
                foreach (XElement rPr in doc.Descendants(w + "rPr").ToList())
                {
                    List<XElement> shd = rPr.Elements(w + "shd").Where(e =>
                        string.Equals((string)e.Attribute(w + "fill"), Scanner.ShadingColor, StringComparison.OrdinalIgnoreCase)).ToList();
                    foreach (XElement e in shd) e.Remove();
                    if (!rPr.HasElements && !rPr.HasAttributes) rPr.Remove();
                }

                bool keepComments = comments != null && comments.Root != null && comments.Root.Elements(w + "comment").Any();
                byte[] newComments = keepComments ? SaveXml(comments) : null;

                byte[] relsBytes;
                if (!entries.TryGetValue(DocumentRelsPart, out relsBytes))
                    relsBytes = Encoding.UTF8.GetBytes("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\" />");
                byte[] typesBytes;
                if (!entries.TryGetValue(ContentTypesPart, out typesBytes))
                    typesBytes = Encoding.UTF8.GetBytes("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\" />");

                if (!keepComments)
                {
                    relsBytes = RemoveCommentsRelationship(relsBytes);
                    typesBytes = RemoveCommentsContentType(typesBytes);
                }

                Dictionary<string, byte[]> replacements = new Dictionary<string, byte[]>(StringComparer.Ordinal);
                replacements[DocumentPart] = SaveXml(doc);
                replacements[DocumentRelsPart] = relsBytes;
                replacements[ContentTypesPart] = typesBytes;
                if (keepComments) replacements[CommentsPart] = newComments;

                HashSet<string> skip = new HashSet<string>(StringComparer.Ordinal);
                if (!keepComments)
                {
                    skip.Add(CommentsPart);
                    skip.Add("word/commentsExtended.xml");
                    skip.Add("word/commentsIds.xml");
                    skip.Add("word/commentsExtensible.xml");
                    skip.Add("word/_rels/comments.xml.rels");
                }

                WriteCleanZip(entries, cleanPath, replacements, skip);

                // Validate the cleaned package before touching the source.
                Dictionary<string, byte[]> verify = ReadEntries(cleanPath);
                if (!verify.ContainsKey(DocumentPart)) throw new InvalidDataException("清理后的文档缺少正文。");
                ParseXml(verify[DocumentPart]);

                using (FileStream fs = new FileStream(sourcePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // This only verifies that the source is not locked by Word.
                }

                if (makeBackup)
                {
                    string dir = Path.GetDirectoryName(sourcePath);
                    string name = Path.GetFileNameWithoutExtension(sourcePath);
                    string ext = Path.GetExtension(sourcePath);
                    result.BackupPath = Path.Combine(dir, name + ".校对备份" + ext);
                    if (File.Exists(result.BackupPath)) File.Delete(result.BackupPath);
                    try
                    {
                        File.Replace(cleanPath, sourcePath, result.BackupPath, true);
                    }
                    catch (Exception)
                    {
                        // File.Replace is atomic, but some network shares do not support it.
                        // Fall back to copy only after the source lock check above.
                        File.Copy(sourcePath, result.BackupPath, true);
                        File.Copy(cleanPath, sourcePath, true);
                    }
                }
                else
                {
                    try
                    {
                        File.Replace(cleanPath, sourcePath, null, true);
                    }
                    catch (Exception)
                    {
                        File.Copy(cleanPath, sourcePath, true);
                    }
                }

                try
                {
                    File.Delete(tempPath);
                    result.TempDeleted = true;
                }
                catch
                {
                    result.TempDeleted = false;
                }

                try { File.Delete(cleanPath); } catch { }

                result.Success = true;
                result.Message = "已写回源文件。";
                if (makeBackup) result.Message += " 备份：" + result.BackupPath;
                if (!result.TempDeleted) result.Message += " 临时文件仍被占用，未能删除，请手动删除。";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = ex.Message;
                try { if (File.Exists(cleanPath)) File.Delete(cleanPath); } catch { }
                return result;
            }
        }

        private static void WriteCleanZip(Dictionary<string, byte[]> entries, string cleanPath, Dictionary<string, byte[]> replacements, HashSet<string> skip)
        {
            if (File.Exists(cleanPath)) File.Delete(cleanPath);
            using (FileStream fs = new FileStream(cleanPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                foreach (KeyValuePair<string, byte[]> pair in entries)
                {
                    if (skip.Contains(pair.Key)) continue;
                    ZipArchiveEntry entry = zip.CreateEntry(pair.Key, CompressionLevel.Optimal);
                    byte[] data;
                    if (replacements.TryGetValue(pair.Key, out data))
                    {
                        using (Stream s = entry.Open()) s.Write(data, 0, data.Length);
                    }
                    else
                    {
                        using (Stream s = entry.Open()) s.Write(pair.Value, 0, pair.Value.Length);
                    }
                }
                // Add any replacement part that did not exist in the source package.
                foreach (KeyValuePair<string, byte[]> pair in replacements)
                {
                    if (entries.ContainsKey(pair.Key)) continue;
                    ZipArchiveEntry entry = zip.CreateEntry(pair.Key, CompressionLevel.Optimal);
                    using (Stream s = entry.Open()) s.Write(pair.Value, 0, pair.Value.Length);
                }
            }
        }

        private static byte[] EnsureCommentsRelationship(byte[] relsBytes)
        {
            XDocument rels = ParseXml(relsBytes);
            if (rels.Root == null)
                rels = new XDocument(new XElement(rel + "Relationships"));
            XElement existing = rels.Root.Elements(rel + "Relationship").FirstOrDefault(e =>
                ((string)e.Attribute("Type") ?? string.Empty).EndsWith("/comments", StringComparison.Ordinal));
            if (existing == null)
            {
                int max = 0;
                foreach (XElement e in rels.Root.Elements(rel + "Relationship"))
                {
                    string id = (string)e.Attribute("Id");
                    if (id != null && id.StartsWith("rId", StringComparison.Ordinal))
                    {
                        int n;
                        if (int.TryParse(id.Substring(3), out n) && n > max) max = n;
                    }
                }
                rels.Root.Add(new XElement(rel + "Relationship",
                    new XAttribute("Id", "rId" + (max + 1).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", CommentsRelationshipType),
                    new XAttribute("Target", "comments.xml")));
            }
            else
            {
                existing.SetAttributeValue("Target", "comments.xml");
            }
            return SaveXml(rels);
        }

        private static byte[] EnsureCommentsContentType(byte[] typesBytes)
        {
            XDocument types = ParseXml(typesBytes);
            if (types.Root == null)
                types = new XDocument(new XElement(ct + "Types"));
            bool exists = types.Root.Elements(ct + "Override").Any(e =>
                string.Equals((string)e.Attribute("PartName"), "/word/comments.xml", StringComparison.OrdinalIgnoreCase));
            if (!exists)
            {
                types.Root.Add(new XElement(ct + "Override",
                    new XAttribute("PartName", "/word/comments.xml"),
                    new XAttribute("ContentType", CommentsContentType)));
            }
            return SaveXml(types);
        }

        private static byte[] RemoveCommentsRelationship(byte[] relsBytes)
        {
            XDocument rels = ParseXml(relsBytes);
            if (rels.Root != null)
            {
                List<XElement> remove = rels.Root.Elements(rel + "Relationship").Where(e =>
                    ((string)e.Attribute("Type") ?? string.Empty).EndsWith("/comments", StringComparison.Ordinal) ||
                    string.Equals((string)e.Attribute("Target"), "comments.xml", StringComparison.OrdinalIgnoreCase)).ToList();
                foreach (XElement e in remove) e.Remove();
            }
            return SaveXml(rels);
        }

        private static byte[] RemoveCommentsContentType(byte[] typesBytes)
        {
            XDocument types = ParseXml(typesBytes);
            if (types.Root != null)
            {
                string[] parts = new string[]
                {
                    "/word/comments.xml",
                    "/word/commentsExtended.xml",
                    "/word/commentsIds.xml",
                    "/word/commentsExtensible.xml"
                };
                List<XElement> remove = types.Root.Elements(ct + "Override").Where(e =>
                {
                    string part = (string)e.Attribute("PartName");
                    if (part == null) return false;
                    foreach (string p in parts)
                    {
                        if (string.Equals(part, p, StringComparison.OrdinalIgnoreCase)) return true;
                    }
                    return false;
                }).ToList();
                foreach (XElement e in remove) e.Remove();
            }
            return SaveXml(types);
        }

        private static bool EndsWithSentencePunctuation(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            char c = value[value.Length - 1];
            return c == '。' || c == '！' || c == '？' || c == '；' || c == '…' || c == '”';
        }

        private static XElement MakeComment(Issue issue, int id, bool hasStyle)
        {
            string body = "【" + issue.Type + "】「" + issue.Original + "」→「" + issue.Suggestion + "」";
            if (!EndsWithSentencePunctuation(issue.Suggestion)) body += "。";
            if (!string.IsNullOrEmpty(issue.Note)) body += " " + issue.Note;
            if (body.Length > 1200) body = body.Substring(0, 1200);

            XElement firstRun = new XElement(w + "r");
            if (hasStyle)
            {
                firstRun.Add(new XElement(w + "rPr",
                    new XElement(w + "rStyle", new XAttribute(w + "val", "CommentReference"))));
            }
            firstRun.Add(new XElement(w + "annotationRef"));
            XElement textRun = new XElement(w + "r",
                new XElement(w + "t", new XAttribute(xml + "space", "preserve"), body));

            return new XElement(w + "comment",
                new XAttribute(w + "id", id.ToString(CultureInfo.InvariantCulture)),
                new XAttribute(w + "author", Scanner.Author),
                new XAttribute(w + "initials", "校"),
                new XAttribute(w + "date", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)),
                new XElement(w + "p", firstRun, textRun));
        }

        private static int MaxCommentId(XDocument comments)
        {
            int max = -1;
            if (comments.Root == null) return max;
            foreach (XElement comment in comments.Root.Elements(w + "comment"))
            {
                int id;
                if (TryGetId(comment, out id) && id > max) max = id;
            }
            return max;
        }

        private static bool TryGetId(XElement element, out int id)
        {
            id = 0;
            XAttribute attr = element.Attribute(w + "id");
            if (attr == null) return false;
            return int.TryParse(attr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out id);
        }

        private static bool AnnotateOne(XElement paragraph, Issue issue, int id, bool hasStyle)
        {
            List<RunInfo> infos = GetRunInfos(paragraph);
            int total = TotalLength(infos);
            if (issue.Start < 0 || issue.Length <= 0 || issue.Start + issue.Length > total) return false;

            SplitAt(paragraph, infos, issue.Start);
            infos = GetRunInfos(paragraph);
            SplitAt(paragraph, infos, issue.Start + issue.Length);
            infos = GetRunInfos(paragraph);

            RunInfo startInfo = null;
            RunInfo endInfo = null;
            int end = issue.Start + issue.Length;
            foreach (RunInfo info in infos)
            {
                if (info.Start <= issue.Start && issue.Start < info.Start + info.Text.Length) startInfo = info;
                if (info.Start < end && end <= info.Start + info.Text.Length) endInfo = info;
            }
            if (startInfo == null || endInfo == null) return false;

            startInfo.Run.AddBeforeSelf(new XElement(w + "commentRangeStart", new XAttribute(w + "id", id.ToString(CultureInfo.InvariantCulture))));
            endInfo.Run.AddAfterSelf(
                new XElement(w + "commentRangeEnd", new XAttribute(w + "id", id.ToString(CultureInfo.InvariantCulture))),
                MakeCommentReferenceRun(id, hasStyle));

            foreach (RunInfo info in infos)
            {
                if (info.Start < end && info.Start + info.Text.Length > issue.Start)
                    AddShading(info.Run);
            }
            return true;
        }

        private static XElement MakeCommentReferenceRun(int id, bool hasStyle)
        {
            XElement run = new XElement(w + "r");
            if (hasStyle)
            {
                run.Add(new XElement(w + "rPr",
                    new XElement(w + "rStyle", new XAttribute(w + "val", "CommentReference"))));
            }
            run.Add(new XElement(w + "commentReference", new XAttribute(w + "id", id.ToString(CultureInfo.InvariantCulture))));
            return run;
        }

        private static void AddShading(XElement run)
        {
            XElement rPr = run.Element(w + "rPr");
            if (rPr == null)
            {
                rPr = new XElement(w + "rPr");
                run.AddFirst(rPr);
            }
            if (rPr.Elements(w + "shd").Any()) return;

            XElement shd = new XElement(w + "shd",
                new XAttribute(w + "val", "clear"),
                new XAttribute(w + "color", "auto"),
                new XAttribute(w + "fill", Scanner.ShadingColor));

            string[] after = new string[] { "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath" };
            XElement before = rPr.Elements().FirstOrDefault(e =>
            {
                foreach (string name in after) if (e.Name.LocalName == name) return true;
                return false;
            });
            if (before != null) before.AddBeforeSelf(shd);
            else rPr.Add(shd);
        }

        private static void SplitAt(XElement paragraph, List<RunInfo> infos, int offset)
        {
            RunInfo target = null;
            foreach (RunInfo info in infos)
            {
                if (offset > info.Start && offset < info.Start + info.Text.Length)
                {
                    target = info;
                    break;
                }
            }
            if (target == null) return;

            int local = offset - target.Start;
            string leftText = target.Text.Substring(0, local);
            string rightText = target.Text.Substring(local);
            if (leftText.Length == 0 || rightText.Length == 0) return;

            XElement left = target.Run;
            XElement right = new XElement(left);
            SetRunText(left, leftText);
            SetRunText(right, rightText);

            // A cloned page-break hint is unnecessary and could make Word render an extra page boundary.
            foreach (XElement e in right.Elements(w + "lastRenderedPageBreak").ToList()) e.Remove();

            left.AddAfterSelf(right);
        }

        private static void SetRunText(XElement run, string text)
        {
            List<XElement> texts = run.Elements(w + "t").ToList();
            XElement first = texts.Count > 0 ? texts[0] : null;
            for (int i = 1; i < texts.Count; i++) texts[i].Remove();
            if (first == null)
            {
                first = new XElement(w + "t");
                run.Add(first);
            }
            first.Value = text;
            first.SetAttributeValue(xml + "space", "preserve");
        }

        private static string GetParagraphText(XElement paragraph)
        {
            StringBuilder sb = new StringBuilder();
            foreach (RunInfo info in GetRunInfos(paragraph)) sb.Append(info.Text);
            return sb.ToString();
        }

        private static int TotalLength(List<RunInfo> infos)
        {
            int total = 0;
            foreach (RunInfo info in infos) total += info.Text.Length;
            return total;
        }

        private static List<RunInfo> GetRunInfos(XElement paragraph)
        {
            List<RunInfo> infos = new List<RunInfo>();
            int offset = 0;
            foreach (XElement run in paragraph.Elements(w + "r"))
            {
                bool hasText = false;
                bool allowed = true;
                foreach (XElement child in run.Elements())
                {
                    if (child.Name == w + "rPr") continue;
                    if (child.Name == w + "t") { hasText = true; continue; }
                    if (child.Name == w + "lastRenderedPageBreak") continue;
                    allowed = false;
                    break;
                }
                if (!allowed || !hasText) continue;

                string text = string.Concat(run.Elements(w + "t").Select(t => t.Value));
                if (text.Length == 0) continue;
                RunInfo info = new RunInfo();
                info.Run = run;
                info.Start = offset;
                info.Text = text;
                infos.Add(info);
                offset += text.Length;
            }
            return infos;
        }

        private class RunInfo
        {
            public XElement Run;
            public int Start;
            public string Text;
        }

        private static Dictionary<string, byte[]> ReadEntries(string path)
        {
            Dictionary<string, byte[]> entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            byte[] bytes = ReadAllBytesShared(path);
            using (MemoryStream ms = new MemoryStream(bytes))
            using (ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    using (Stream s = entry.Open())
                    using (MemoryStream copy = new MemoryStream())
                    {
                        s.CopyTo(copy);
                        entries[entry.FullName] = copy.ToArray();
                    }
                }
            }
            return entries;
        }

        private static byte[] ReadAllBytesShared(string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (MemoryStream ms = new MemoryStream())
            {
                fs.CopyTo(ms);
                return ms.ToArray();
            }
        }

        private static void CopyFileShared(string source, string dest)
        {
            using (FileStream input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (FileStream output = new FileStream(dest, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                input.CopyTo(output);
            }
        }

        private static XDocument ParseXml(byte[] bytes)
        {
            using (MemoryStream ms = new MemoryStream(bytes))
                return XDocument.Load(ms, LoadOptions.PreserveWhitespace);
        }

        private static byte[] SaveXml(XDocument doc)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                doc.Save(ms, SaveOptions.DisableFormatting);
                return ms.ToArray();
            }
        }

        private static void ReplaceEntry(ZipArchive zip, string name, byte[] data)
        {
            ZipArchiveEntry existing = zip.GetEntry(name);
            if (existing != null) existing.Delete();
            ZipArchiveEntry entry = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (Stream s = entry.Open())
            {
                s.Write(data, 0, data.Length);
            }
        }
    }
}
