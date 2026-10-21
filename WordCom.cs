using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace NovelProofreader
{
    public enum CloseDocumentResult
    {
        NotOpen,
        Closed,
        Failed
    }

    public static class WordComHelper
    {
        public static bool OpenDocument(string path, out string error)
        {
            error = null;
            object appObject = null;
            dynamic doc = null;
            try
            {
                appObject = GetActive("Word.Application");
                if (appObject == null) appObject = GetActive("KWPS.Application");
                if (appObject == null)
                {
                    Type wordType = Type.GetTypeFromProgID("Word.Application");
                    if (wordType == null) wordType = Type.GetTypeFromProgID("KWPS.Application");
                    if (wordType == null)
                    {
                        error = "未检测到 Microsoft Word。";
                        return false;
                    }
                    appObject = Activator.CreateInstance(wordType);
                }

                dynamic app = appObject;
                try { app.Visible = true; } catch { }
                try { app.DisplayAlerts = 0; } catch { }

                doc = FindDocument(app, path);
                if (doc == null) doc = app.Documents.Open(path, false, false, false);
                if (doc == null)
                {
                    error = "Word 未能打开临时文件。";
                    return false;
                }
                try { doc.Activate(); } catch { }
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                if (doc != null)
                {
                    try { Marshal.ReleaseComObject(doc); } catch { }
                }
                if (appObject != null)
                {
                    try { Marshal.ReleaseComObject(appObject); } catch { }
                }
            }
        }

        private static object GetActive(string progId)
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
                    doc = app.Documents[i];
                    string fullName = null;
                    try { fullName = doc.FullName; }
                    catch { continue; }
                    if (!string.IsNullOrEmpty(fullName) && PathsEqual(fullName, path)) return doc;
                }
                catch { }
                if (doc != null)
                {
                    try { Marshal.ReleaseComObject(doc); }
                    catch { }
                }
            }
            return null;
        }

        public static bool IsDocumentOpen(string path)
        {
            try
            {
                string target;
                try { target = Path.GetFullPath(path); }
                catch { target = path; }

                string[] progIds = new string[] { "Word.Application", "KWPS.Application", "WPS.Application" };
                foreach (string progId in progIds)
                {
                    object appObject = null;
                    try { appObject = Marshal.GetActiveObject(progId); }
                    catch { continue; }
                    if (appObject == null) continue;

                    try
                    {
                        dynamic app = appObject;
                        int count;
                        try { count = app.Documents.Count; }
                        catch { continue; }

                        for (int i = 1; i <= count; i++)
                        {
                            dynamic doc = null;
                            try
                            {
                                doc = app.Documents[i];
                                string fullName = null;
                                try { fullName = doc.FullName; }
                                catch { continue; }
                                if (!string.IsNullOrEmpty(fullName) && PathsEqual(fullName, target))
                                    return true;
                            }
                            catch { }
                            finally
                            {
                                if (doc != null)
                                {
                                    try { Marshal.ReleaseComObject(doc); }
                                    catch { }
                                }
                            }
                        }
                    }
                    finally
                    {
                        try { Marshal.ReleaseComObject(appObject); }
                        catch { }
                    }
                }
            }
            catch { }
            return false;
        }

        public static CloseDocumentResult CloseIfOpen(string path, bool saveChanges, out string error)
        {
            try
            {
                return CloseIfOpenCore(path, saveChanges, out error);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return CloseDocumentResult.Failed;
            }
        }

        // Word can briefly report RPC_E_* / "object does not exist" while a document
        // is still opening or closing. Retry only while the file is actually locked.
        public static CloseDocumentResult CloseIfOpenWithRetry(string path, bool saveChanges, int attempts, out string error)
        {
            error = null;
            if (attempts < 1) attempts = 1;
            CloseDocumentResult result = CloseDocumentResult.NotOpen;
            for (int i = 0; i < attempts; i++)
            {
                string currentError;
                result = CloseIfOpen(path, saveChanges, out currentError);
                if (!string.IsNullOrEmpty(currentError)) error = currentError;

                if (result == CloseDocumentResult.Closed)
                    return CloseDocumentResult.Closed;

                if (!IsLocked(path))
                    return CloseDocumentResult.NotOpen;

                Thread.Sleep(300);
            }

            if (string.IsNullOrEmpty(error))
                error = "临时文件仍被 Word 占用。";
            return CloseDocumentResult.Failed;
        }

        private static CloseDocumentResult CloseIfOpenCore(string path, bool saveChanges, out string error)
        {
            error = null;
            string target;
            try { target = Path.GetFullPath(path); }
            catch { target = path; }

            string[] progIds = new string[] { "Word.Application", "KWPS.Application", "WPS.Application" };
            foreach (string progId in progIds)
            {
                object appObject = null;
                try
                {
                    appObject = Marshal.GetActiveObject(progId);
                }
                catch
                {
                    continue;
                }
                if (appObject == null) continue;

                try
                {
                    dynamic app = appObject;
                    int count;
                    try { count = app.Documents.Count; }
                    catch { continue; }

                    for (int i = 1; i <= count; i++)
                    {
                        dynamic doc = null;
                        bool matched = false;
                        try
                        {
                            doc = app.Documents[i];
                            string fullName = null;
                            try { fullName = doc.FullName; }
                            catch { continue; }
                            if (string.IsNullOrEmpty(fullName)) continue;
                            if (!PathsEqual(fullName, target)) continue;
                            matched = true;

                            try { app.DisplayAlerts = 0; }
                            catch { }
                            if (saveChanges)
                            {
                                bool saved = false;
                                try { saved = doc.Saved; }
                                catch { saved = false; }
                                if (!saved) doc.Save();
                            }
                            doc.Close(0);
                            return CloseDocumentResult.Closed;
                        }
                        catch (Exception ex)
                        {
                            if (matched)
                            {
                                error = ex.Message;
                                return CloseDocumentResult.Failed;
                            }
                        }
                        finally
                        {
                            if (doc != null)
                            {
                                try { Marshal.ReleaseComObject(doc); }
                                catch { }
                            }
                        }
                    }
                }
                finally
                {
                    try { Marshal.ReleaseComObject(appObject); }
                    catch { }
                }
            }
            return CloseDocumentResult.NotOpen;
        }

        public static bool IsLocked(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch (FileNotFoundException)
            {
                return false;
            }
            catch
            {
                return true;
            }
        }

        public static bool TryDelete(string path, int attempts)
        {
            for (int i = 0; i < attempts; i++)
            {
                try
                {
                    if (!File.Exists(path)) return true;
                    File.Delete(path);
                    return true;
                }
                catch
                {
                }
                Thread.Sleep(120);
            }
            return !File.Exists(path);
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
    }
}
