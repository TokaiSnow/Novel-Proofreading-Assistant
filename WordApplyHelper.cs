using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace NovelProofreader
{
    public class ApplyResult
    {
        public bool Success;
        public bool WordAvailable = true;
        public string Error = string.Empty;
        public List<Issue> AppliedIssues = new List<Issue>();
        public List<string> Failures = new List<string>();
    }

    public static class WordApplyHelper
    {
        public static ApplyResult ApplySelected(string tempPath, List<Issue> selected)
        {
            ApplyResult result = new ApplyResult();
            if (selected == null || selected.Count == 0)
            {
                result.Success = true;
                return result;
            }
            if (string.IsNullOrEmpty(tempPath) || !File.Exists(tempPath))
            {
                result.Error = "临时文件不存在。";
                return result;
            }

            object appObject = null;
            bool createdApplication = false;
            dynamic doc = null;
            List<CommentSnapshot> snapshots = new List<CommentSnapshot>();

            try
            {
                appObject = GetActiveObject("Word.Application");
                if (appObject == null) appObject = GetActiveObject("KWPS.Application");
                if (appObject == null)
                {
                    Type wordType = Type.GetTypeFromProgID("Word.Application");
                    if (wordType == null) wordType = Type.GetTypeFromProgID("KWPS.Application");
                    if (wordType == null)
                    {
                        result.WordAvailable = false;
                        result.Error = "未检测到 Microsoft Word，无法自动应用修改。请在 Word 中手动修改。";
                        return result;
                    }
                    appObject = Activator.CreateInstance(wordType);
                    createdApplication = true;
                }

                dynamic app = appObject;
                try { app.Visible = true; } catch { }
                try { app.DisplayAlerts = 0; } catch { }

                doc = FindDocument(app, tempPath);
                if (doc == null)
                {
                    doc = app.Documents.Open(tempPath, false, false, false);
                }
                if (doc == null)
                {
                    result.Error = "无法在 Word 中打开临时文件。";
                    return result;
                }

                int commentCount = 0;
                try { commentCount = doc.Comments.Count; }
                catch { commentCount = 0; }
                for (int i = 1; i <= commentCount; i++)
                {
                    dynamic comment = null;
                    try
                    {
                        comment = doc.Comments.Item(i);
                        CommentSnapshot snapshot = new CommentSnapshot();
                        snapshot.Comment = comment;
                        try { snapshot.Author = comment.Author; } catch { snapshot.Author = string.Empty; }
                        try { snapshot.Scope = comment.Scope.Text; } catch { snapshot.Scope = string.Empty; }
                        try { snapshot.Body = comment.Range.Text; } catch { snapshot.Body = string.Empty; }
                        snapshots.Add(snapshot);
                    }
                    catch
                    {
                        if (comment != null)
                        {
                            try { Marshal.ReleaseComObject(comment); } catch { }
                        }
                    }
                }

                foreach (Issue issue in selected)
                {
                    CommentSnapshot match = FindMatch(snapshots, issue, true);
                    if (match == null) match = FindMatch(snapshots, issue, false);
                    if (match == null)
                    {
                        result.Failures.Add(issue.Location + "「" + issue.Original + "」");
                        continue;
                    }

                    try
                    {
                        dynamic comment = match.Comment;
                        try { comment.Scope.Font.Shading.BackgroundPatternColor = -16777216; } catch { }
                        try { comment.Scope.Shading.BackgroundPatternColor = -16777216; } catch { }
                        try { comment.Scope.HighlightColorIndex = 0; } catch { }
                        comment.Scope.Text = issue.Suggestion;
                        try { comment.Delete(); } catch { }
                        match.Used = true;
                        result.AppliedIssues.Add(issue);
                    }
                    catch (Exception ex)
                    {
                        result.Failures.Add(issue.Location + "「" + issue.Original + "」：" + ex.Message);
                    }
                }

                if (result.AppliedIssues.Count > 0)
                {
                    doc.Save();
                }
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = ex.Message;
                return result;
            }
            finally
            {
                for (int i = 0; i < snapshots.Count; i++)
                {
                    if (snapshots[i].Comment != null)
                    {
                        try { Marshal.ReleaseComObject(snapshots[i].Comment); } catch { }
                    }
                }
                if (doc != null)
                {
                    try { Marshal.ReleaseComObject(doc); } catch { }
                }
                if (appObject != null)
                {
                    try { Marshal.ReleaseComObject(appObject); } catch { }
                }
                // Word is intentionally left open so the user can review the applied changes.
                GC.KeepAlive(createdApplication);
            }
        }

        private static object GetActiveObject(string progId)
        {
            try { return Marshal.GetActiveObject(progId); }
            catch { return null; }
        }

        private static dynamic FindDocument(dynamic app, string path)
        {
            int count = 0;
            try { count = app.Documents.Count; }
            catch { return null; }
            for (int i = 1; i <= count; i++)
            {
                dynamic doc = null;
                try
                {
                    doc = app.Documents.Item(i);
                    string fullName = null;
                    try { fullName = doc.FullName; }
                    catch { }
                    if (!string.IsNullOrEmpty(fullName) && PathsEqual(fullName, path)) return doc;
                }
                catch { }
                if (doc != null)
                {
                    try { Marshal.ReleaseComObject(doc); } catch { }
                }
            }
            return null;
        }

        private static CommentSnapshot FindMatch(List<CommentSnapshot> snapshots, Issue issue, bool requireBody)
        {
            for (int i = 0; i < snapshots.Count; i++)
            {
                CommentSnapshot snapshot = snapshots[i];
                if (snapshot.Used) continue;
                if (!string.IsNullOrEmpty(snapshot.Author) &&
                    !string.Equals(snapshot.Author, Scanner.Author, StringComparison.Ordinal)) continue;
                if (!string.Equals(snapshot.Scope, issue.Original, StringComparison.Ordinal)) continue;
                if (requireBody)
                {
                    bool bodyMatch = snapshot.Body.IndexOf("「" + issue.Suggestion + "」", StringComparison.Ordinal) >= 0 ||
                                     snapshot.Body.IndexOf(issue.Suggestion, StringComparison.Ordinal) >= 0;
                    if (!bodyMatch) continue;
                }
                return snapshot;
            }
            return null;
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

        private class CommentSnapshot
        {
            public object Comment;
            public string Author = string.Empty;
            public string Scope = string.Empty;
            public string Body = string.Empty;
            public bool Used;
        }
    }
}
