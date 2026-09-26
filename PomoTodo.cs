// PomoTodo v1.5.0 - Pomodoro timer + Todo list, tray mode, Excel (.xlsx) import/export, cloud folder sync (OneDrive / Google Drive / any folder; one file per PC, auto-merge, no conflict copies)
// Target: .NET Framework 4.x (built into Windows 10/11)
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Media;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;

[assembly: System.Reflection.AssemblyTitle("PomoTodo")]
[assembly: System.Reflection.AssemblyProduct("PomoTodo")]
[assembly: System.Reflection.AssemblyVersion("1.5.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.5.0.0")]
[assembly: System.Reflection.AssemblyInformationalVersion("1.5.0")]

namespace PomoTodo
{
    class TodoItem
    {
        public string Id = Guid.NewGuid().ToString("N");
        public string Text = "";
        public bool Done;
        public DateTime Created = DateTime.Now;
        public DateTime? DoneAt;
        public int Pomos;     // completed pomodoros
        public int Target;    // estimated pomodoros (0 = not set)
        public DateTime Modified = DateTime.MinValue;   // last change (sync: newest copy wins)
    }

    class RecordItem
    {
        public DateTime Start;
        public DateTime End;
        public double Minutes;
        public string Type = "Focus";   // Focus / Short Break / Long Break
        public string Task = "";
        public string Status = "Completed"; // Completed / Interrupted
        public string TaskId = "";
        public string Key { get { return Start.ToString(Store.DF, Store.IC) + "|" + Type + "|" + Task; } }
    }

    class AppSettings
    {
        public int Focus = 25, ShortBreak = 5, LongBreak = 15, LongEvery = 4;
        public bool AutoStartBreak = true, AutoStartFocus = false, Sound = true, TopMost = true;
        public bool CloseToTray = true, StartWithWindows = false, AutoExport = true, TrayShowsUsed = false, PopupWhenDone = false;
        public string LogPath = "";
        public int CountMinutes = 25;
        public string HotkeyStart = "Ctrl+Alt+P", HotkeyShow = "Ctrl+Alt+O";
        public string SoundStart = "", SoundEnd = "";   // "" = built-in chime
        public bool ToastStart = true, ToastEnd = true, ToastEndStays = true;   // "" = <sync folder>\\PomoTodo_Log.xlsx
    }

    static class Store
    {
        public const string DF = "yyyy-MM-dd HH:mm:ss";
        public static readonly CultureInfo IC = CultureInfo.InvariantCulture;

        // Local folder: always holds settings + a backup copy of data
        public static string AppDir
        {
            get
            {
                string d = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PomoTodo");
                Directory.CreateDirectory(d);
                return d;
            }
        }
        // Sync folder (e.g. G:\My Drive\PomoTodo). Empty = local only.
        static string syncDir;
        public static string SyncDir
        {
            get
            {
                if (syncDir == null)
                {
                    string f = Path.Combine(AppDir, "syncfolder.txt");
                    syncDir = File.Exists(f) ? File.ReadAllText(f, Encoding.UTF8).Trim() : "";
                }
                return syncDir;
            }
        }
        public static bool IsSynced { get { return SyncDir != ""; } }
        public static bool SyncAvailable { get { return IsSynced && Directory.Exists(SyncDir); } }
        public static void SetSyncDir(string d)
        {
            d = (d ?? "").Trim();
            if (d != "") Directory.CreateDirectory(d);
            File.WriteAllText(Path.Combine(AppDir, "syncfolder.txt"), d, Encoding.UTF8);
            syncDir = d;
            known.Clear();
        }

        public static string FindOneDrive()
        {
            foreach (var v in new[] { "OneDrive", "OneDriveConsumer", "OneDriveCommercial" })
            {
                string p = Environment.GetEnvironmentVariable(v);
                if (!string.IsNullOrEmpty(p) && Directory.Exists(p)) return p;
            }
            try
            {
                string up = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                foreach (var d in Directory.GetDirectories(up, "OneDrive*")) return d;
            }
            catch { }
            return null;
        }

        public static string FindGoogleDrive()
        {
            var names = new[] { "My Drive", "我的云端硬盘", "我的雲端硬碟", "Meine Ablage", "Mon Drive" };
            try
            {
                foreach (var d in DriveInfo.GetDrives())
                {
                    try
                    {
                        if (!d.IsReady) continue;
                        foreach (var n in names)
                        {
                            string p = Path.Combine(d.RootDirectory.FullName, n);
                            if (Directory.Exists(p)) return p;
                        }
                    }
                    catch { }
                }
            }
            catch { }
            string up = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            foreach (var n in new[] { Path.Combine("Google Drive", "My Drive"), "My Drive", "Google Drive", "我的云端硬盘" })
            {
                string p = Path.Combine(up, n);
                if (Directory.Exists(p)) return p;
            }
            return null;
        }

        static string Clean(string s) { return (s ?? "").Replace("\t", " ").Replace("\r", " ").Replace("\n", " "); }
        static string D(DateTime? d) { return d.HasValue ? d.Value.ToString(DF, IC) : ""; }
        static DateTime? PD(string s)
        {
            DateTime d;
            if (DateTime.TryParseExact(s, DF, IC, DateTimeStyles.None, out d)) return d;
            return null;
        }

        static string[] ReadLines(string f)
        {
            // retry: cloud client may be replacing / downloading the file
            for (int i = 0; i < 3; i++)
            {
                try
                {
                    using (var fs = new FileStream(f, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    using (var r = new StreamReader(fs, Encoding.UTF8))
                        return r.ReadToEnd().Split(new[] { '\n' }).Select(l => l.TrimEnd(new[] { '\r' })).Where(l => l != "").ToArray();
                }
                catch (IOException) { Thread.Sleep(150); }
                catch (UnauthorizedAccessException) { Thread.Sleep(150); }
                catch (Exception) { break; }
            }
            return new string[0];
        }
        static void WriteSafe(string f, string content) { WriteBytes(f, new UTF8Encoding(true).GetBytes(content)); }

        // Direct overwrite (no temp file / rename - those are often refused inside OneDrive), with retries
        public static void WriteBytes(string f, byte[] data)
        {
            Exception last = null;
            for (int i = 0; i < 4; i++)
            {
                try
                {
                    string dir = Path.GetDirectoryName(f);
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    if (File.Exists(f))
                    {
                        var at = File.GetAttributes(f);
                        if ((at & FileAttributes.ReadOnly) != 0) File.SetAttributes(f, at & ~FileAttributes.ReadOnly);
                    }
                    using (var fs = new FileStream(f, FileMode.Create, FileAccess.Write, FileShare.Read))
                        fs.Write(data, 0, data.Length);
                    try { if (File.Exists(f + ".tmp")) File.Delete(f + ".tmp"); } catch { } // leftover from v1.1
                    return;
                }
                catch (IOException ex) { last = ex; }
                catch (UnauthorizedAccessException ex) { last = ex; }
                Thread.Sleep(250 * (i + 1));
            }
            throw last;
        }

        // Try to create + delete a small file in the folder. Returns null if OK, else the error.
        public static Exception TestWrite(string dir)
        {
            try
            {
                Directory.CreateDirectory(dir);
                string f = Path.Combine(dir, "~pomotodo_write_test.txt");
                WriteBytes(f, Encoding.UTF8.GetBytes("test"));
                File.Delete(f);
                return null;
            }
            catch (Exception ex) { return ex; }
        }

        public static bool ControlledFolderAccessOn()
        {
            try
            {
                using (var k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows Defender\Windows Defender Exploit Guard\Controlled Folder Access"))
                {
                    object v = k == null ? null : k.GetValue("EnableControlledFolderAccess");
                    return v != null && Convert.ToInt32(v) == 1;
                }
            }
            catch { return false; }
        }

        // ---------------- Sync model (v1.4) ----------------
        // Each PC writes ONLY its own files in the sync folder:  todos@PC.tsv, records@PC.tsv, deleted@PC.tsv
        // Two PCs never write the same file, so OneDrive / Google Drive never create conflict copies
        // (like records-DESKTOP-XXX-2.tsv). On load, every data file in the folder is merged, including the
        // old shared todos.tsv / records.tsv and old conflict copies:
        //   todos   - per Id, the copy with the newest "Modified" time wins
        //   records - union (same start time + type + task = same record)
        //   deleted - union of tombstones; a deleted todo/record is removed on every PC
        public static string PC
        {
            get
            {
                var s = new string((Environment.MachineName ?? "").Select(c => char.IsLetterOrDigit(c) || c == '-' || c == '_' ? c : '_').ToArray());
                return s == "" ? "PC" : s;
            }
        }
        static string MyFile(string kind) { return kind + "@" + PC + ".tsv"; }

        // todos.tsv, todos-DESKTOP-X.tsv, todos-DESKTOP-X-2.tsv (OneDrive conflict copies), "todos (1).tsv" (Google Drive), todos@PC.tsv
        static string[] DataFiles(string dir, string kind)
        {
            try
            {
                if (dir == "" || !Directory.Exists(dir)) return new string[0];
                return Directory.GetFiles(dir, kind + "*.tsv").Where(f =>
                {
                    string n = Path.GetFileNameWithoutExtension(f);
                    return n.Equals(kind, StringComparison.OrdinalIgnoreCase) || n.StartsWith(kind + "-", StringComparison.OrdinalIgnoreCase) || n.StartsWith(kind + "@", StringComparison.OrdinalIgnoreCase) || n.StartsWith(kind + " (", StringComparison.OrdinalIgnoreCase);
                }).OrderBy(f => Path.GetFileNameWithoutExtension(f).Length == kind.Length ? 0 : 1).ThenBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray();
            }
            catch { return new string[0]; }
        }
        static IEnumerable<string> AllDataFiles(string dir) { return DataFiles(dir, "todos").Concat(DataFiles(dir, "records")).Concat(DataFiles(dir, "deleted")); }
        static bool IsLegacy(string f) { return !Path.GetFileName(f).Contains("@"); }

        static List<TodoItem> ReadTodosFile(string f)
        {
            var list = new List<TodoItem>();
            if (!File.Exists(f)) return list;
            foreach (var line in ReadLines(f))
            {
                var p = line.Split(new[] { '\t' });
                if (p.Length < 6) continue;
                int n; int.TryParse(p[5], out n);
                int tg = 0; if (p.Length > 6) int.TryParse(p[6], out tg);
                var t = new TodoItem { Id = p[0], Text = p[1], Done = p[2] == "1", Created = PD(p[3]) ?? DateTime.Now, DoneAt = PD(p[4]), Pomos = n, Target = tg };
                t.Modified = (p.Length > 7 ? PD(p[7]) : null) ?? t.DoneAt ?? t.Created;   // old files have no Modified column
                list.Add(t);
            }
            return list;
        }
        static string TodoRow(TodoItem t, bool withModified)
        {
            var sb = new StringBuilder();
            sb.Append(t.Id).Append('\t').Append(Clean(t.Text)).Append('\t').Append(t.Done ? "1" : "0").Append('\t')
              .Append(D(t.Created)).Append('\t').Append(D(t.DoneAt)).Append('\t').Append(t.Pomos).Append('\t').Append(t.Target);
            if (withModified) sb.Append('\t').Append(D(t.Modified));
            return sb.ToString();
        }
        static string TodosText(List<TodoItem> list)
        {
            var sb = new StringBuilder();
            foreach (var t in list) sb.Append(TodoRow(t, true)).Append('\n');
            return sb.ToString();
        }
        static List<RecordItem> ReadRecordsFile(string f)
        {
            var list = new List<RecordItem>();
            if (!File.Exists(f)) return list;
            foreach (var line in ReadLines(f))
            {
                var p = line.Split(new[] { '\t' });
                if (p.Length < 6) continue;
                var s = PD(p[0]); var e = PD(p[1]);
                if (s == null || e == null) continue;
                double m; double.TryParse(p[2], NumberStyles.Any, IC, out m);
                list.Add(new RecordItem { Start = s.Value, End = e.Value, Minutes = m, Type = p[3], Task = p[4], Status = p[5], TaskId = p.Length > 6 ? p[6] : "" });
            }
            return list;
        }
        static string RecordsText(List<RecordItem> list)
        {
            var sb = new StringBuilder();
            foreach (var r in list)
                sb.Append(D(r.Start)).Append('\t').Append(D(r.End)).Append('\t').Append(r.Minutes.ToString("0.##", IC)).Append('\t')
                  .Append(Clean(r.Type)).Append('\t').Append(Clean(r.Task)).Append('\t').Append(Clean(r.Status)).Append('\t').Append(r.TaskId ?? "").Append('\n');
            return sb.ToString();
        }
        public static string RKey(RecordItem r) { return Clean(r.Key); }
        public static bool IsDeleted(RecordItem r) { return tombs.ContainsKey("R" + RKey(r)); }

        // Move files that are already merged into <their folder>\old_files (keeps the cloud folder clean, deletes nothing)
        public static void MoveToOld(string f)
        {
            try
            {
                string bak = Path.Combine(Path.GetDirectoryName(f), "old_files");
                Directory.CreateDirectory(bak);
                string to = Path.Combine(bak, Path.GetFileName(f));
                if (File.Exists(to)) File.Delete(to);
                File.Move(f, to);
            }
            catch { }
        }

        // Tombstones: "T" + todo id / "R" + record key -> deletion time
        static readonly Dictionary<string, DateTime> tombs = new Dictionary<string, DateTime>();
        static void ReadTombs(string f)
        {
            if (!File.Exists(f)) return;
            foreach (var line in ReadLines(f))
            {
                int i = line.LastIndexOf('\t');
                if (i <= 0) continue;
                string key = line.Substring(0, i);
                DateTime at = PD(line.Substring(i + 1)) ?? DateTime.Now;
                DateTime old;
                if (!tombs.TryGetValue(key, out old) || at > old) tombs[key] = at;
            }
        }
        static string TombsText()
        {
            var sb = new StringBuilder();
            foreach (var kv in tombs.OrderBy(x => x.Value)) sb.Append(Clean(kv.Key)).Append('\t').Append(D(kv.Value)).Append('\n');
            return sb.ToString();
        }

        // what we last loaded/saved - used to find which todos changed and what was deleted on this PC
        static readonly Dictionary<string, string> snapTodos = new Dictionary<string, string>();
        static readonly HashSet<string> snapRecs = new HashSet<string>();
        static void Snapshot(List<TodoItem> todos, List<RecordItem> records)
        {
            snapTodos.Clear(); snapRecs.Clear();
            foreach (var t in todos) snapTodos[t.Id] = TodoRow(t, false);
            foreach (var r in records) snapRecs.Add(RKey(r));
        }
        static void Track(List<TodoItem> todos, List<RecordItem> records)
        {
            var now = DateTime.Now;
            var ids = new HashSet<string>();
            foreach (var t in todos)
            {
                ids.Add(t.Id);
                string old;
                if (!snapTodos.TryGetValue(t.Id, out old) || old != TodoRow(t, false) || t.Modified == DateTime.MinValue) t.Modified = now;
            }
            foreach (var id in snapTodos.Keys) if (!ids.Contains(id)) tombs["T" + id] = now;
            var keys = new HashSet<string>(records.Select(RKey));
            foreach (var k in snapRecs) if (!keys.Contains(k)) tombs["R" + k] = now;
            Snapshot(todos, records);
        }

        // mtime of every sync-folder file at the moment we read or wrote it
        static readonly Dictionary<string, DateTime> known = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
        static DateTime MTimeOf(string f) { try { return File.GetLastWriteTimeUtc(f); } catch { return DateTime.MinValue; } }

        static bool Newer(TodoItem a, TodoItem b)
        {
            if (a.Modified != b.Modified) return a.Modified > b.Modified;
            if (a.Pomos != b.Pomos) return a.Pomos > b.Pomos;
            return a.Done && !b.Done;
        }
        static void Collect(string dir, bool remember, List<string> order, Dictionary<string, TodoItem> tm, Dictionary<string, RecordItem> rm)
        {
            foreach (var f in DataFiles(dir, "deleted")) { if (remember) known[f] = MTimeOf(f); ReadTombs(f); }
            foreach (var f in DataFiles(dir, "todos"))
            {
                if (remember) known[f] = MTimeOf(f);
                foreach (var t in ReadTodosFile(f))
                {
                    TodoItem o;
                    if (!tm.TryGetValue(t.Id, out o)) { order.Add(t.Id); tm[t.Id] = t; }
                    else if (Newer(t, o)) tm[t.Id] = t;
                }
            }
            foreach (var f in DataFiles(dir, "records"))
            {
                if (remember) known[f] = MTimeOf(f);
                foreach (var r in ReadRecordsFile(f)) { string k = RKey(r); if (!rm.ContainsKey(k)) rm[k] = r; }
            }
        }
        static void Finish(List<string> order, Dictionary<string, TodoItem> tm, Dictionary<string, RecordItem> rm, out List<TodoItem> todos, out List<RecordItem> records)
        {
            todos = order.Select(id => tm[id]).Where(t => !tombs.ContainsKey("T" + t.Id)).ToList();
            records = rm.Where(kv => !tombs.ContainsKey("R" + kv.Key)).Select(kv => kv.Value).OrderBy(r => r.Start).ToList();
        }

        // Merged data of one folder (used when choosing a sync folder)
        public static void LoadFolder(string dir, out List<TodoItem> todos, out List<RecordItem> records)
        {
            var order = new List<string>(); var tm = new Dictionary<string, TodoItem>(); var rm = new Dictionary<string, RecordItem>();
            Collect(dir, false, order, tm, rm);
            Finish(order, tm, rm, out todos, out records);
        }

        public static bool HasData(string dir) { return DataFiles(dir, "todos").Length > 0 || DataFiles(dir, "records").Length > 0; }

        // Load = merge of the local backup + every PC's files in the sync folder
        public static void LoadData(out List<TodoItem> todos, out List<RecordItem> records)
        {
            var order = new List<string>(); var tm = new Dictionary<string, TodoItem>(); var rm = new Dictionary<string, RecordItem>();
            Collect(AppDir, false, order, tm, rm);
            if (SyncAvailable) Collect(SyncDir, true, order, tm, rm);
            Finish(order, tm, rm, out todos, out records);
            Snapshot(todos, records);
        }

        public static string LastSyncError = null;
        public static Exception LastSyncException = null;
        public static DateTime? LastSyncOk = null;
        static readonly Dictionary<string, string> lastWritten = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        static void WriteSync(string f, string content)
        {
            string old;
            if (File.Exists(f) && lastWritten.TryGetValue(f, out old) && old == content && known.ContainsKey(f) && Math.Abs((MTimeOf(f) - known[f]).TotalSeconds) <= 1) return;
            WriteSafe(f, content);
            lastWritten[f] = content; known[f] = MTimeOf(f);
        }
        public static void SaveData(List<TodoItem> todos, List<RecordItem> records)
        {
            Track(todos, records);
            string tt = TodosText(todos), rt = RecordsText(records), dt = TombsText();
            WriteSafe(Path.Combine(AppDir, "todos.tsv"), tt);
            WriteSafe(Path.Combine(AppDir, "records.tsv"), rt);
            WriteSafe(Path.Combine(AppDir, "deleted.tsv"), dt);
            if (!IsSynced) return;
            try
            {
                if (!Directory.Exists(SyncDir)) Directory.CreateDirectory(SyncDir);
                WriteSync(Path.Combine(SyncDir, MyFile("todos")), tt);
                WriteSync(Path.Combine(SyncDir, MyFile("records")), rt);
                WriteSync(Path.Combine(SyncDir, MyFile("deleted")), dt);
                ArchiveLegacy();
                LastSyncError = null; LastSyncException = null; LastSyncOk = DateTime.Now;
            }
            catch (Exception ex) { LastSyncError = ex.Message; LastSyncException = ex; }
        }

        // Old shared files (todos.tsv) and conflict copies (records-DESKTOP-X-2.tsv) are already merged into
        // our own files -> move them to "old_files" so the folder stays clean. Only files unchanged since we read them.
        static void ArchiveLegacy()
        {
            foreach (var f in AllDataFiles(SyncDir).Where(IsLegacy).ToList())
            {
                try
                {
                    DateTime k;
                    if (!known.TryGetValue(f, out k) || Math.Abs((MTimeOf(f) - k).TotalSeconds) > 1) continue;
                    MoveToOld(f);
                    known.Remove(f);
                }
                catch { }
            }
        }

        // True if another PC changed / added a file in the sync folder since we last read it
        public static bool SyncChangedExternally()
        {
            if (!SyncAvailable) return false;
            foreach (var f in AllDataFiles(SyncDir))
            {
                DateTime k;
                if (!known.TryGetValue(f, out k) || Math.Abs((MTimeOf(f) - k).TotalSeconds) > 1) return true;
            }
            return false;
        }

        public static void Merge(List<TodoItem> todos, List<RecordItem> records, List<TodoItem> otherTodos, List<RecordItem> otherRecords)
        {
            var ids = new HashSet<string>(todos.Select(t => t.Id));
            foreach (var t in otherTodos) if (ids.Add(t.Id)) todos.Add(t);
            var keys = new HashSet<string>(records.Select(r => r.Key));
            foreach (var r in otherRecords) if (keys.Add(r.Key)) records.Add(r);
        }

        public static AppSettings LoadSettings()
        {
            var s = new AppSettings();
            string f = Path.Combine(AppDir, "settings.txt");
            if (!File.Exists(f)) return s;
            foreach (var line in File.ReadAllLines(f, Encoding.UTF8))
            {
                int i = line.IndexOf('='); if (i < 0) continue;
                string k = line.Substring(0, i).Trim(), v = line.Substring(i + 1).Trim();
                int n; int.TryParse(v, out n);
                bool b = v == "1";
                switch (k)
                {
                    case "Focus": if (n > 0) s.Focus = n; break;
                    case "ShortBreak": if (n > 0) s.ShortBreak = n; break;
                    case "LongBreak": if (n > 0) s.LongBreak = n; break;
                    case "LongEvery": if (n > 0) s.LongEvery = n; break;
                    case "AutoStartBreak": s.AutoStartBreak = b; break;
                    case "AutoStartFocus": s.AutoStartFocus = b; break;
                    case "Sound": s.Sound = b; break;
                    case "TopMost": s.TopMost = b; break;
                    case "CloseToTray": s.CloseToTray = b; break;
                    case "StartWithWindows": s.StartWithWindows = b; break;
                    case "AutoExport": s.AutoExport = b; break;
                    case "TrayShowsUsed": s.TrayShowsUsed = b; break;
                    case "ShowWindowWhenDone": s.PopupWhenDone = b; break;
                    case "CountMinutes": if (n > 0) s.CountMinutes = n; break;
                    case "HotkeyStart": s.HotkeyStart = v; break;
                    case "HotkeyShow": s.HotkeyShow = v; break;
                    case "SoundStart": s.SoundStart = v; break;
                    case "SoundEnd": s.SoundEnd = v; break;
                    case "ToastStart": s.ToastStart = b; break;
                    case "ToastEnd": s.ToastEnd = b; break;
                    case "ToastEndStays": s.ToastEndStays = b; break;
                    case "LogPath": s.LogPath = v; break;
                }
            }
            return s;
        }
        public static void SaveSettings(AppSettings s)
        {
            Func<bool, string> B = x => x ? "1" : "0";
            File.WriteAllText(Path.Combine(AppDir, "settings.txt"),
                "Focus=" + s.Focus + "\nShortBreak=" + s.ShortBreak + "\nLongBreak=" + s.LongBreak + "\nLongEvery=" + s.LongEvery +
                "\nAutoStartBreak=" + B(s.AutoStartBreak) + "\nAutoStartFocus=" + B(s.AutoStartFocus) +
                "\nSound=" + B(s.Sound) + "\nTopMost=" + B(s.TopMost) + "\nCloseToTray=" + B(s.CloseToTray) +
                "\nStartWithWindows=" + B(s.StartWithWindows) + "\nAutoExport=" + B(s.AutoExport) +
                "\nTrayShowsUsed=" + B(s.TrayShowsUsed) + "\nShowWindowWhenDone=" + B(s.PopupWhenDone) + "\nLogPath=" + (s.LogPath ?? "") +
                "\nCountMinutes=" + s.CountMinutes + "\nHotkeyStart=" + s.HotkeyStart + "\nHotkeyShow=" + s.HotkeyShow +
                "\nSoundStart=" + s.SoundStart + "\nSoundEnd=" + s.SoundEnd + "\nToastStart=" + B(s.ToastStart) +
                "\nToastEnd=" + B(s.ToastEnd) + "\nToastEndStays=" + B(s.ToastEndStays) + "\n", Encoding.UTF8);
        }
    }

    // ---------------- Minimal XLSX writer / reader (no external libraries) ----------------
    static class Xlsx
    {
        static string Esc(string s)
        {
            if (s == null) return "";
            var sb = new StringBuilder();
            foreach (char c in s) if (c == '\t' || c == '\n' || c == '\r' || c >= 0x20) sb.Append(c);
            return SecurityElement.Escape(sb.ToString());
        }
        static string ColName(int i)
        {
            string s = ""; i++;
            while (i > 0) { int m = (i - 1) % 26; s = (char)('A' + m) + s; i = (i - 1) / 26; }
            return s;
        }
        static void Put(ZipArchive z, string name, string content)
        {
            var e = z.CreateEntry(name, CompressionLevel.Optimal);
            using (var w = new StreamWriter(e.Open(), new UTF8Encoding(false))) w.Write(content);
        }

        public static void Write(string path, List<KeyValuePair<string, List<object[]>>> sheets)
        {
            var ms = new MemoryStream();
            using (var z = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var ct = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
                var wb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
                var rels = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for (int i = 0; i < sheets.Count; i++)
                {
                    int n = i + 1;
                    ct.Append("<Override PartName=\"/xl/worksheets/sheet" + n + ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                    wb.Append("<sheet name=\"" + Esc(sheets[i].Key) + "\" sheetId=\"" + n + "\" r:id=\"rId" + n + "\"/>");
                    rels.Append("<Relationship Id=\"rId" + n + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet" + n + ".xml\"/>");
                    Put(z, "xl/worksheets/sheet" + n + ".xml", SheetXml(sheets[i].Value));
                }
                int sid = sheets.Count + 1;
                rels.Append("<Relationship Id=\"rId" + sid + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");
                ct.Append("</Types>");
                wb.Append("</sheets></workbook>");
                Put(z, "[Content_Types].xml", ct.ToString());
                Put(z, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                Put(z, "xl/workbook.xml", wb.ToString());
                Put(z, "xl/_rels/workbook.xml.rels", rels.ToString());
                Put(z, "xl/styles.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"11\"/><color rgb=\"FFFFFFFF\"/><name val=\"Calibri\"/></font></fonts><fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFE5533D\"/><bgColor indexed=\"64\"/></patternFill></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>");
            }
            Store.WriteBytes(path, ms.ToArray());
        }

        static string SheetXml(List<object[]> rows)
        {
            int cols = rows.Count == 0 ? 1 : rows.Max(r => r.Length);
            var widths = new double[cols];
            foreach (var r in rows)
                for (int c = 0; c < r.Length; c++)
                {
                    string s = r[c] == null ? "" : Convert.ToString(r[c], Store.IC);
                    double w = 0; foreach (char ch in s) w += ch > 0x2E80 ? 2.1 : 1.1;
                    widths[c] = Math.Max(widths[c], w);
                }
            var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sb.Append("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews><cols>");
            for (int c = 0; c < cols; c++)
                sb.Append("<col min=\"" + (c + 1) + "\" max=\"" + (c + 1) + "\" width=\"" + Math.Min(60, Math.Max(9, widths[c] + 2)).ToString("0.#", Store.IC) + "\" customWidth=\"1\"/>");
            sb.Append("</cols><sheetData>");
            for (int r = 0; r < rows.Count; r++)
            {
                sb.Append("<row r=\"" + (r + 1) + "\">");
                for (int c = 0; c < rows[r].Length; c++)
                {
                    object v = rows[r][c];
                    if (v == null) continue;
                    string refc = ColName(c) + (r + 1);
                    string st = r == 0 ? " s=\"1\"" : "";
                    if (v is int || v is double || v is long || v is float || v is decimal)
                        sb.Append("<c r=\"" + refc + "\"" + st + "><v>" + Convert.ToString(v, Store.IC) + "</v></c>");
                    else
                        sb.Append("<c r=\"" + refc + "\"" + st + " t=\"inlineStr\"><is><t xml:space=\"preserve\">" + Esc(v.ToString()) + "</t></is></c>");
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData></worksheet>");
            return sb.ToString();
        }

        static XDocument Load(ZipArchive z, string name)
        {
            var e = z.GetEntry(name);
            if (e == null) return null;
            using (var s = e.Open()) return XDocument.Load(s);
        }
        static IEnumerable<XElement> All(XContainer x, string local) { return x.Descendants().Where(e => e.Name.LocalName == local); }
        static int ColIndex(string cellRef)
        {
            int n = 0;
            foreach (char ch in cellRef) { if (ch >= 'A' && ch <= 'Z') n = n * 26 + (ch - 'A' + 1); else if (ch >= 'a' && ch <= 'z') n = n * 26 + (ch - 'a' + 1); else break; }
            return n - 1;
        }

        // Returns ordered list of (sheet name, rows)
        public static List<KeyValuePair<string, List<string[]>>> Read(string path)
        {
            var result = new List<KeyValuePair<string, List<string[]>>>();
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var z = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var shared = new List<string>();
                var sst = Load(z, "xl/sharedStrings.xml");
                if (sst != null)
                    foreach (var si in All(sst, "si"))
                        shared.Add(string.Concat(si.Descendants().Where(e => e.Name.LocalName == "t" && e.Parent.Name.LocalName != "rPh").Select(e => e.Value)));
                var relMap = new Dictionary<string, string>();
                var rels = Load(z, "xl/_rels/workbook.xml.rels");
                if (rels != null)
                    foreach (var r in All(rels, "Relationship"))
                    {
                        string t = (string)r.Attribute("Target") ?? "";
                        t = t.StartsWith("/") ? t.TrimStart(new[] { '/' }) : "xl/" + t;
                        relMap[(string)r.Attribute("Id") ?? ""] = t;
                    }
                var wb = Load(z, "xl/workbook.xml");
                if (wb == null) throw new Exception("Not a valid .xlsx file");
                foreach (var sh in All(wb, "sheet"))
                {
                    string name = (string)sh.Attribute("name") ?? "Sheet";
                    string rid = sh.Attributes().Where(a => a.Name.LocalName == "id").Select(a => a.Value).FirstOrDefault() ?? "";
                    string target;
                    if (!relMap.TryGetValue(rid, out target)) continue;
                    var sx = Load(z, target);
                    if (sx == null) continue;
                    var rows = new List<string[]>();
                    foreach (var row in All(sx, "row"))
                    {
                        var cells = new SortedDictionary<int, string>();
                        int auto = 0;
                        foreach (var c in row.Elements().Where(e => e.Name.LocalName == "c"))
                        {
                            string r = (string)c.Attribute("r");
                            int ci = r != null ? ColIndex(r) : auto;
                            auto = ci + 1;
                            string t = (string)c.Attribute("t") ?? "";
                            var vEl = c.Elements().FirstOrDefault(e => e.Name.LocalName == "v");
                            string val;
                            if (t == "s") { int idx; val = vEl != null && int.TryParse(vEl.Value, out idx) && idx < shared.Count ? shared[idx] : ""; }
                            else if (t == "inlineStr") val = string.Concat(All(c, "t").Select(e => e.Value));
                            else val = vEl != null ? vEl.Value : "";
                            cells[ci] = val;
                        }
                        if (cells.Count == 0) { rows.Add(new string[0]); continue; }
                        var arr = new string[cells.Keys.Max() + 1];
                        foreach (var kv in cells) arr[kv.Key] = kv.Value;
                        rows.Add(arr);
                    }
                    result.Add(new KeyValuePair<string, List<string[]>>(name, rows));
                }
            }
            return result;
        }

        public static List<string[]> ReadCsv(string path)
        {
            var rows = new List<string[]>();
            string text = File.ReadAllText(path, Encoding.UTF8);
            var cur = new List<string>(); var sb = new StringBuilder(); bool q = false;
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (q) { if (ch == '"') { if (i + 1 < text.Length && text[i + 1] == '"') { sb.Append('"'); i++; } else q = false; } else sb.Append(ch); }
                else if (ch == '"') q = true;
                else if (ch == ',') { cur.Add(sb.ToString()); sb.Clear(); }
                else if (ch == '\n' || ch == '\r')
                {
                    if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    cur.Add(sb.ToString()); sb.Clear(); rows.Add(cur.ToArray()); cur.Clear();
                }
                else sb.Append(ch);
            }
            if (sb.Length > 0 || cur.Count > 0) { cur.Add(sb.ToString()); rows.Add(cur.ToArray()); }
            return rows;
        }
    }

    // ---------------- Export / import logic ----------------
    static class DataIO
    {
        public static bool Counted(RecordItem r, int threshold) { return r.Type == "Focus" && r.Minutes >= threshold - 0.01; }

        public static List<KeyValuePair<string, List<object[]>>> BuildSheets(List<TodoItem> todos, List<RecordItem> records, int threshold)
        {
            var sheets = new List<KeyValuePair<string, List<object[]>>>();
            var rec = new List<object[]> { new object[] { "Date", "Start", "End", "Minutes", "Type", "Task", "Status", "Counted as Pomodoro" } };
            foreach (var r in records.OrderBy(x => x.Start))
                rec.Add(new object[] { r.Start.ToString("yyyy-MM-dd"), r.Start.ToString("HH:mm:ss"), r.End.ToString("HH:mm:ss"), Math.Round(r.Minutes, 1), r.Type, r.Task, r.Status, r.Type == "Focus" ? (Counted(r, threshold) ? "Yes" : "No") : "" });
            sheets.Add(new KeyValuePair<string, List<object[]>>("Records", rec));

            var td = new List<object[]> { new object[] { "Task", "Status", "Done Pomodoros", "Target", "Progress %", "Created", "Completed At" } };
            foreach (var t in todos)
                td.Add(new object[] { t.Text, t.Done ? "Done" : "Open", t.Pomos, t.Target, t.Target > 0 ? (object)Math.Round(100.0 * t.Pomos / t.Target, 1) : "", t.Created.ToString(Store.DF), t.DoneAt.HasValue ? t.DoneAt.Value.ToString(Store.DF) : "" });
            sheets.Add(new KeyValuePair<string, List<object[]>>("Todos", td));

            var focus = records.Where(r => r.Type == "Focus").ToList();
            var daily = new List<object[]> { new object[] { "Date", "Pomodoros (>= " + threshold + " min)", "Focus Minutes", "Short sessions (not counted)" } };
            foreach (var g in focus.GroupBy(r => r.Start.Date).OrderBy(g => g.Key))
                daily.Add(new object[] { g.Key.ToString("yyyy-MM-dd"), g.Count(r => Counted(r, threshold)), Math.Round(g.Sum(r => r.Minutes), 1), g.Count(r => !Counted(r, threshold)) });
            sheets.Add(new KeyValuePair<string, List<object[]>>("Daily Summary", daily));

            var byTask = new List<object[]> { new object[] { "Task", "Pomodoros (>= " + threshold + " min)", "Focus Minutes", "First", "Last" } };
            foreach (var g in focus.GroupBy(r => string.IsNullOrEmpty(r.Task) ? "(no task)" : r.Task).OrderByDescending(g => g.Sum(r => r.Minutes)))
                byTask.Add(new object[] { g.Key, g.Count(r => Counted(r, threshold)), Math.Round(g.Sum(r => r.Minutes), 1), g.Min(r => r.Start).ToString("yyyy-MM-dd"), g.Max(r => r.Start).ToString("yyyy-MM-dd") });
            sheets.Add(new KeyValuePair<string, List<object[]>>("Task Summary", byTask));

            var dayTask = new List<object[]> { new object[] { "Date", "Task", "Pomodoros (>= " + threshold + " min)", "Focus Minutes" } };
            foreach (var g in focus.GroupBy(r => new { D = r.Start.Date, T = string.IsNullOrEmpty(r.Task) ? "(no task)" : r.Task }).OrderBy(g => g.Key.D).ThenByDescending(g => g.Sum(r => r.Minutes)))
                dayTask.Add(new object[] { g.Key.D.ToString("yyyy-MM-dd"), g.Key.T, g.Count(r => Counted(r, threshold)), Math.Round(g.Sum(r => r.Minutes), 1) });
            sheets.Add(new KeyValuePair<string, List<object[]>>("Daily by Task", dayTask));
            return sheets;
        }

        static string Cell(string[] row, int i) { return row != null && i >= 0 && i < row.Length ? (row[i] ?? "").Trim() : ""; }
        static int Find(string[] header, params string[] names)
        {
            if (header == null) return -1;
            for (int i = 0; i < header.Length; i++)
            {
                string h = (header[i] ?? "").Trim().ToLowerInvariant();
                foreach (var n in names) if (h == n) return i;
            }
            return -1;
        }
        static DateTime? ExcelDate(string date, string time)
        {
            DateTime d; double od;
            string s = (date + " " + time).Trim();
            if (DateTime.TryParse(s, Store.IC, DateTimeStyles.None, out d)) return d;
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out d)) return d;
            if (double.TryParse(date, NumberStyles.Any, Store.IC, out od))
            {
                double ot; double.TryParse(time, NumberStyles.Any, Store.IC, out ot);
                try { return DateTime.FromOADate(Math.Floor(od) + (ot > 0 && ot < 1 ? ot : od - Math.Floor(od))); } catch { }
            }
            return null;
        }

        static List<KeyValuePair<string, List<string[]>>> ReadSheets(string path)
        {
            if (path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                return new List<KeyValuePair<string, List<string[]>>> { new KeyValuePair<string, List<string[]>>("csv", Xlsx.ReadCsv(path)) };
            return Xlsx.Read(path);
        }
        static List<string[]> Sheet(List<KeyValuePair<string, List<string[]>>> sheets, string name)
        {
            return sheets.FirstOrDefault(s => s.Key.Equals(name, StringComparison.OrdinalIgnoreCase)).Value;
        }

        // The "Records" sheet of an exported workbook -> records
        static List<RecordItem> ParseRecords(List<string[]> rows)
        {
            var list = new List<RecordItem>();
            if (rows == null || rows.Count < 2) return list;
            var h = rows[0];
            int cD = Find(h, "date", "日期"), cS = Find(h, "start", "开始"), cE = Find(h, "end", "结束"), cM = Find(h, "minutes", "分钟"),
                cT = Find(h, "type", "类型"), cK = Find(h, "task", "任务"), cSt = Find(h, "status", "状态");
            foreach (var row in rows.Skip(1))
            {
                var st = ExcelDate(Cell(row, cD), Cell(row, cS));
                if (st == null) continue;
                double m; double.TryParse(Cell(row, cM), NumberStyles.Any, Store.IC, out m);
                DateTime en = ExcelDate(Cell(row, cD), Cell(row, cE)) ?? st.Value.AddMinutes(m);
                if (en < st.Value) en = en.AddDays(1);
                list.Add(new RecordItem { Start = st.Value, End = en, Minutes = m > 0 ? m : (en - st.Value).TotalMinutes, Type = cT >= 0 && Cell(row, cT) != "" ? Cell(row, cT) : "Focus", Task = Cell(row, cK), Status = cSt >= 0 && Cell(row, cSt) != "" ? Cell(row, cSt) : "Completed" });
            }
            return list;
        }
        public static List<RecordItem> ReadRecords(string path) { return ParseRecords(Sheet(ReadSheets(path), "Records")); }

        // Adds records that are not in the list yet (and were not deleted); returns how many were added
        public static int AddRecords(List<RecordItem> records, IEnumerable<RecordItem> more)
        {
            var existing = new HashSet<string>(records.Select(Store.RKey));
            int n = 0;
            foreach (var r in more)
                if (!Store.IsDeleted(r) && existing.Add(Store.RKey(r))) { records.Add(r); n++; }
            return n;
        }

        // Text form of the sheets, to see if a workbook already holds exactly this data (then it is not rewritten)
        static string Norm(string[] row)
        {
            var cells = row.Select(c => c ?? "").ToList();
            while (cells.Count > 0 && cells[cells.Count - 1] == "") cells.RemoveAt(cells.Count - 1);
            return string.Join("\u0001", cells);
        }
        public static string Signature(List<KeyValuePair<string, List<object[]>>> sheets)
        {
            return string.Join("\u0003", sheets.Select(s => s.Key + "\u0002" + string.Join("\u0002", s.Value.Select(r =>
                Norm(r.Select(v => v == null ? "" : v is string ? (string)v : Convert.ToString(v, Store.IC)).ToArray())))));
        }
        public static string Signature(string path)
        {
            try { return string.Join("\u0003", Xlsx.Read(path).Select(s => s.Key + "\u0002" + string.Join("\u0002", s.Value.Select(Norm)))); }
            catch { return null; }
        }

        // Returns message describing what was imported
        public static string Import(string path, List<TodoItem> todos, List<RecordItem> records)
        {
            var sheets = ReadSheets(path);
            if (sheets.Count == 0) return "No sheets found.";

            int addedTodos = 0, addedRecords = 0;
            var recSheet = sheets.FirstOrDefault(s => s.Key.Equals("Records", StringComparison.OrdinalIgnoreCase));
            var todoSheet = sheets.FirstOrDefault(s => s.Key.Equals("Todos", StringComparison.OrdinalIgnoreCase));

            addedRecords = AddRecords(records, ParseRecords(recSheet.Value));

            List<string[]> taskRows = todoSheet.Value;
            if (taskRows == null && recSheet.Value == null) taskRows = sheets[0].Value;   // generic sheet: task list
            if (taskRows != null && taskRows.Count > 0)
            {
                var h = taskRows[0];
                int cTask = Find(h, "task", "todo", "tasks", "todos", "name", "text", "title", "任务", "待办", "事项", "名称");
                int cStat = Find(h, "status", "状态");
                int cPomo = Find(h, "done pomodoros", "pomodoros", "pomos", "done", "番茄", "番茄数", "已完成");
                int cTarget = Find(h, "target", "estimate", "预计", "预计番茄", "目标");
                bool hasHeader = cTask >= 0;
                if (cTask < 0) cTask = 0;
                var open = new HashSet<string>(todos.Where(t => !t.Done).Select(t => t.Text.Trim().ToLowerInvariant()));
                foreach (var row in taskRows.Skip(hasHeader ? 1 : 0))
                {
                    string text = Cell(row, cTask);
                    if (text == "") continue;
                    string stat = Cell(row, cStat).ToLowerInvariant();
                    bool done = stat == "done" || stat == "yes" || stat == "true" || stat == "1" || stat == "完成" || stat == "已完成";
                    if (!done && !open.Add(text.ToLowerInvariant())) continue;
                    if (done && todos.Any(t => t.Done && t.Text == text)) continue;
                    int p, tg; int.TryParse(Cell(row, cPomo), out p); int.TryParse(Cell(row, cTarget), out tg);
                    var item = new TodoItem { Text = text, Done = done, Pomos = p, Target = tg };
                    if (done) item.DoneAt = item.Created; // imported done tasks: completion time = creation time (same day)
                    todos.Add(item);
                    addedTodos++;
                }
            }
            return "已导入 " + addedTodos + " 个任务，" + addedRecords + " 条番茄记录。";
        }
    }

    // ---------------- Sounds: built-in chimes + custom wav/mp3 ----------------
    static class Sound
    {
        [DllImport("winmm.dll", CharSet = CharSet.Unicode)] static extern int mciSendString(string cmd, StringBuilder ret, int len, IntPtr cb);
        static SoundPlayer player; // keep reference so it is not collected while playing

        static byte[] Chime(double[] freqs, double gap, double len)
        {
            int rate = 44100;
            int n = (int)(rate * (gap * (freqs.Length - 1) + len));
            var buf = new double[n];
            for (int i = 0; i < freqs.Length; i++)
            {
                int start = (int)(i * gap * rate), L = (int)(len * rate);
                double f = freqs[i];
                for (int j = 0; j < L && start + j < n; j++)
                {
                    double t = (double)j / rate;
                    double env = Math.Exp(-t * 3.2) * Math.Min(1.0, j / 300.0);
                    double v = Math.Sin(2 * Math.PI * f * t) * 0.62 + Math.Sin(2 * Math.PI * 2 * f * t) * 0.22 + Math.Sin(2 * Math.PI * 3.01 * f * t) * 0.08;
                    buf[start + j] += v * env * 0.45;
                }
            }
            var ms = new MemoryStream();
            var w = new BinaryWriter(ms);
            int dataLen = n * 2;
            w.Write(Encoding.ASCII.GetBytes("RIFF")); w.Write(36 + dataLen); w.Write(Encoding.ASCII.GetBytes("WAVE"));
            w.Write(Encoding.ASCII.GetBytes("fmt ")); w.Write(16); w.Write((short)1); w.Write((short)1); w.Write(rate); w.Write(rate * 2); w.Write((short)2); w.Write((short)16);
            w.Write(Encoding.ASCII.GetBytes("data")); w.Write(dataLen);
            foreach (var d in buf) w.Write((short)Math.Max(short.MinValue, Math.Min(short.MaxValue, d * 32767)));
            w.Flush();
            return ms.ToArray();
        }
        static byte[] startWav, endFocusWav, endBreakWav;
        public static byte[] TestWav() { return Chime(new[] { 659.25, 987.77 }, 0.16, 0.9); }

        public static void Play(string kind, string customFile)
        {
            try
            {
                if (!string.IsNullOrEmpty(customFile) && File.Exists(customFile))
                {
                    mciSendString("close pomosnd", null, 0, IntPtr.Zero);
                    if (mciSendString("open \"" + customFile + "\" type mpegvideo alias pomosnd", null, 0, IntPtr.Zero) == 0)
                    { mciSendString("play pomosnd", null, 0, IntPtr.Zero); return; }
                }
            }
            catch { }
            try
            {
                byte[] data;
                if (kind == "start") data = startWav ?? (startWav = Chime(new[] { 659.25, 987.77 }, 0.16, 0.9));
                else if (kind == "endBreak") data = endBreakWav ?? (endBreakWav = Chime(new[] { 987.77, 783.99, 659.25 }, 0.2, 1.1));
                else data = endFocusWav ?? (endFocusWav = Chime(new[] { 783.99, 987.77, 1318.51, 1567.98 }, 0.18, 1.4));
                if (player != null) player.Stop();
                player = new SoundPlayer(new MemoryStream(data));
                player.Play();
            }
            catch { try { SystemSounds.Exclamation.Play(); } catch { } }
        }
    }

    // ---------------- Global hotkeys ----------------
    static class Hotkey
    {
        public static bool Parse(string s, out uint mods, out Keys key)
        {
            mods = 0; key = Keys.None;
            if (string.IsNullOrWhiteSpace(s)) return false;
            foreach (var raw in s.Split(new[] { '+' }))
            {
                string p = raw.Trim();
                switch (p.ToLowerInvariant())
                {
                    case "ctrl": case "control": mods |= 0x2; break;
                    case "alt": mods |= 0x1; break;
                    case "shift": mods |= 0x4; break;
                    case "win": mods |= 0x8; break;
                    default:
                        Keys k;
                        if (p.Length == 1 && char.IsDigit(p[0])) p = "D" + p;
                        if (!Enum.TryParse(p, true, out k)) return false;
                        key = k; break;
                }
            }
            return key != Keys.None;
        }
        public static string KeyName(Keys k)
        {
            string n = k.ToString();
            if (n.Length == 2 && n[0] == 'D' && char.IsDigit(n[1])) return n.Substring(1);
            return n;
        }
    }

    class HotkeyBox : TextBox
    {
        public HotkeyBox() { ReadOnly = true; BackColor = Color.White; ShortcutsEnabled = false; Cursor = Cursors.Hand; }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            Keys key = keyData & Keys.KeyCode, mods = keyData & Keys.Modifiers;
            if (key == Keys.ControlKey || key == Keys.Menu || key == Keys.ShiftKey || key == Keys.LWin || key == Keys.RWin) return true;
            if (mods == Keys.None && (key == Keys.Back || key == Keys.Delete || key == Keys.Escape)) { Text = ""; return true; }
            bool fkey = key >= Keys.F1 && key <= Keys.F24;
            if (mods == Keys.None && !fkey) return true; // need Ctrl/Alt/Shift
            var parts = new List<string>();
            if ((mods & Keys.Control) != 0) parts.Add("Ctrl");
            if ((mods & Keys.Alt) != 0) parts.Add("Alt");
            if ((mods & Keys.Shift) != 0) parts.Add("Shift");
            parts.Add(Hotkey.KeyName(key));
            Text = string.Join("+", parts);
            return true;
        }
    }

    // ---------------- Popup notification window ----------------
    class Toast : Form
    {
        static readonly List<Toast> open = new List<Toast>();
        System.Windows.Forms.Timer life;
        Color accent;
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override CreateParams CreateParams
        {
            get { var cp = base.CreateParams; cp.ExStyle |= 0x08000000 | 0x00000080 | 0x00000008; return cp; } // NOACTIVATE | TOOLWINDOW | TOPMOST
        }

        public static void ShowToast(string title, string msg, Color accent, Font baseFont, float k, int autoCloseMs, string[] buttons, Action<int> onButton)
        {
            foreach (var t in open.ToArray()) { try { t.Close(); } catch { } }
            var f = new Toast(title, msg, accent, baseFont, k, autoCloseMs, buttons, onButton);
            open.Add(f);
            f.FormClosed += (a, b) => open.Remove(f);
            f.Show();
        }

        Toast(string title, string msg, Color accent, Font baseFont, float k, int autoCloseMs, string[] buttons, Action<int> onButton)
        {
            this.accent = accent;
            Func<int, int> S = v => (int)Math.Round(v * k);
            FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; TopMost = true; StartPosition = FormStartPosition.Manual;
            BackColor = Color.White; Width = S(340);
            var fTitle = new Font(baseFont.FontFamily, 12f, FontStyle.Bold);
            var fMsg = new Font(baseFont.FontFamily, 10f);
            var lt = new Label { Text = title, Font = fTitle, ForeColor = accent, AutoSize = false, Location = new Point(S(22), S(14)), Size = new Size(S(290), S(26)) };
            var lm = new Label { Text = msg, Font = fMsg, ForeColor = Color.FromArgb(60, 60, 60), AutoSize = false, Location = new Point(S(22), S(42)) };
            var sz = TextRenderer.MeasureText(msg, fMsg, new Size(S(300), 1000), TextFormatFlags.WordBreak);
            lm.Size = new Size(S(300), Math.Max(S(22), sz.Height + S(4)));
            var close = new Label { Text = "×", Font = new Font("Arial", 13f), ForeColor = Color.Gray, AutoSize = false, Size = new Size(S(24), S(24)), Location = new Point(S(310), S(6)), Cursor = Cursors.Hand, TextAlign = ContentAlignment.MiddleCenter };
            close.Click += (a, b) => Close();
            Controls.AddRange(new Control[] { lt, lm, close });
            int y = lm.Bottom + S(10);
            if (buttons != null && buttons.Length > 0)
            {
                int x = Width - S(16);
                for (int i = buttons.Length - 1; i >= 0; i--)
                {
                    int idx = i;
                    var b = new Button { Text = buttons[i], FlatStyle = FlatStyle.Flat, Font = fMsg, Height = S(30), Cursor = Cursors.Hand, UseVisualStyleBackColor = false };
                    b.Width = TextRenderer.MeasureText(b.Text, fMsg).Width + S(24);
                    b.FlatAppearance.BorderColor = i == 0 ? accent : Color.FromArgb(210, 210, 210);
                    b.BackColor = i == 0 ? accent : Color.White; b.ForeColor = i == 0 ? Color.White : Color.Black;
                    x -= b.Width; b.Location = new Point(x, y); x -= S(8);
                    b.Click += (a, e) => { Close(); if (onButton != null) onButton(idx); };
                    Controls.Add(b);
                }
                y += S(30) + S(12);
            }
            Height = y + S(4);
            var wa = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(wa.Right - Width - S(16), wa.Bottom - Height - S(16));
            foreach (Control c in Controls) if (!(c is Button) && c != close) c.Click += (a, b) => { Close(); if (onButton != null) onButton(-1); };
            if (autoCloseMs > 0)
            {
                life = new System.Windows.Forms.Timer { Interval = autoCloseMs };
                life.Tick += (a, b) => { life.Stop(); Close(); };
                life.Start();
            }
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using (var p = new Pen(Color.FromArgb(200, 200, 200))) e.Graphics.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
            using (var b = new SolidBrush(accent)) e.Graphics.FillRectangle(b, 0, 0, (int)Math.Max(5, Width / 60f), Height);
        }
        protected override void OnFormClosed(FormClosedEventArgs e) { if (life != null) life.Dispose(); base.OnFormClosed(e); }
    }

    // ---------------- Settings dialog ----------------
    class SettingsForm : Form
    {
        NumericUpDown nFocus, nShort, nLong, nEvery, nCount;
        CheckBox cAutoBreak, cAutoFocus, cSound, cTray, cStartup, cExport, cUsed, cPopup, cToastStart, cToastEnd, cToastStays;
        HotkeyBox hkStart, hkShow;
        TextBox tSync;
        public AppSettings Result;
        public string NewSyncDir; // null = unchanged, "" = local only
        Func<string> logPathGetter;

        public SettingsForm(AppSettings s, Font f, float k)
        {
            Text = "设置 Settings"; Font = f; FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen; AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink; ShowInTaskbar = false; TopMost = true;
            var tabs = new TabControl { Dock = DockStyle.Fill };
            Func<string, TableLayoutPanel> page = name =>
            {
                var tp = new TabPage(name) { BackColor = Color.White, AutoScroll = true };
                var t = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Padding = new Padding(10), Dock = DockStyle.Top };
                t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 320 * k)); t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110 * k));
                tp.Controls.Add(t); tabs.TabPages.Add(tp); return t;
            };
            TableLayoutPanel tl = null;
            Func<string, Label> head = t => { var l = new Label { Text = t, AutoSize = true, Font = new Font(f, FontStyle.Bold), Margin = new Padding(3, 10, 3, 3) }; tl.Controls.Add(l); tl.SetColumnSpan(l, 2); return l; };
            Func<string, Label> note = t => { var l = new Label { Text = t, AutoSize = true, MaximumSize = new Size((int)(430 * k), 0), ForeColor = Color.DimGray }; tl.Controls.Add(l); tl.SetColumnSpan(l, 2); return l; };
            Func<string, int, int, NumericUpDown> addNum = (label, val, max) =>
            {
                tl.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 8, 12, 3) });
                var n = new NumericUpDown { Minimum = 1, Maximum = max, Value = Math.Min(max, Math.Max(1, val)), Width = (int)(70 * k) };
                tl.Controls.Add(n); return n;
            };
            Func<string, bool, CheckBox> addChk = (t, v) => { var c = new CheckBox { Text = t, Checked = v, AutoSize = true }; tl.Controls.Add(c); tl.SetColumnSpan(c, 2); return c; };
            Func<Control[], FlowLayoutPanel> addRow = cs => { var r = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill }; r.Controls.AddRange(cs); tl.Controls.Add(r); tl.SetColumnSpan(r, 2); return r; };

            // ---- Page 1: timer
            tl = page("番茄钟");
            head("时长");
            nFocus = addNum("专注时长（分钟）", s.Focus, 180);
            nShort = addNum("短休息（分钟）", s.ShortBreak, 60);
            nLong = addNum("长休息（分钟）", s.LongBreak, 90);
            nEvery = addNum("每几个番茄后长休息", s.LongEvery, 20);
            nCount = addNum("专注满多少分钟才计为 1 个番茄", s.CountMinutes, 180);
            note("不到这个分钟数的专注（例如中途停止）会记录下来，但不计入番茄数。");
            head("自动");
            cAutoBreak = addChk("专注结束后自动开始休息", s.AutoStartBreak);
            cAutoFocus = addChk("休息结束后自动开始下一个专注", s.AutoStartFocus);
            head("托盘（通知区域）");
            cTray = addChk("点 X 关闭窗口时，程序继续在托盘运行", s.CloseToTray);
            cUsed = addChk("托盘图标数字显示“已用分钟”（不勾 = 剩余分钟）", s.TrayShowsUsed);
            cStartup = addChk("开机自动启动（隐藏在托盘）", s.StartWithWindows);

            // ---- Page 2: alerts + hotkeys
            tl = page("提醒与快捷键");
            head("铃声");
            cSound = addChk("开始和结束番茄时播放铃声", s.Sound);
            string soundStart = s.SoundStart ?? "", soundEnd = s.SoundEnd ?? "";
            Func<string, string> sname = p => p == "" ? "内置铃声" : Path.GetFileName(p);
            var lStart = new Label { AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
            var lEnd = new Label { AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
            Action upd = () => { lStart.Text = "开始铃声：" + sname(soundStart); lEnd.Text = "结束铃声：" + sname(soundEnd); };
            upd();
            Func<string> pick = () =>
            {
                using (var d = new OpenFileDialog { Filter = "音频 (*.wav;*.mp3)|*.wav;*.mp3" })
                    return d.ShowDialog(this) == DialogResult.OK ? d.FileName : null;
            };
            var b1 = new Button { Text = "选择...", AutoSize = true }; var b1t = new Button { Text = "试听", AutoSize = true }; var b1d = new Button { Text = "默认", AutoSize = true };
            var b2 = new Button { Text = "选择...", AutoSize = true }; var b2t = new Button { Text = "试听", AutoSize = true }; var b2d = new Button { Text = "默认", AutoSize = true };
            b1.Click += (a, e) => { var p = pick(); if (p != null) { soundStart = p; upd(); } };
            b2.Click += (a, e) => { var p = pick(); if (p != null) { soundEnd = p; upd(); } };
            b1t.Click += (a, e) => Sound.Play("start", soundStart);
            b2t.Click += (a, e) => Sound.Play("endFocus", soundEnd);
            b1d.Click += (a, e) => { soundStart = ""; upd(); };
            b2d.Click += (a, e) => { soundEnd = ""; upd(); };
            addRow(new Control[] { lStart, b1, b1t, b1d });
            addRow(new Control[] { lEnd, b2, b2t, b2d });
            note("可以用自己的 .wav 或 .mp3 文件。");
            head("屏幕弹窗提示");
            cToastStart = addChk("开始番茄时，屏幕右下角弹出提示", s.ToastStart);
            cToastEnd = addChk("番茄/休息结束时，屏幕右下角弹出提示", s.ToastEnd);
            cToastStays = addChk("结束提示一直显示，直到我点击", s.ToastEndStays);
            cPopup = addChk("结束时同时把主窗口弹到最前面", s.PopupWhenDone);
            head("全局快捷键（在任何程序里都能用）");
            tl.Controls.Add(new Label { Text = "开始番茄 / 暂停 / 继续", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 8, 12, 3) });
            hkStart = new HotkeyBox { Text = s.HotkeyStart ?? "", Width = (int)(110 * k) }; tl.Controls.Add(hkStart);
            tl.Controls.Add(new Label { Text = "显示 / 隐藏窗口", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 8, 12, 3) });
            hkShow = new HotkeyBox { Text = s.HotkeyShow ?? "", Width = (int)(110 * k) }; tl.Controls.Add(hkShow);
            note("点一下输入框，然后直接按下组合键（例如 Ctrl+Alt+P）。按 Backspace 清除 = 不使用。休息中按“开始番茄”快捷键会直接跳过休息开始专注。");

            // ---- Page 3: sync + excel
            tl = page("同步与 Excel");
            head("同步文件夹（OneDrive / Google Drive / 任意文件夹）");
            note("任务和番茄记录会保存到这个文件夹。每台电脑只写自己的文件（todos@电脑名.tsv），打开时自动合并所有电脑的数据，不会再产生冲突副本。");
            tSync = new TextBox { Text = Store.SyncDir == "" ? "（关闭 - 只存本机）" : Store.SyncDir, ReadOnly = true, Width = (int)(430 * k) };
            tl.Controls.Add(tSync); tl.SetColumnSpan(tSync, 2);
            var bOne = new Button { Text = "使用 OneDrive", AutoSize = true };
            var bBrowse = new Button { Text = "选择文件夹...", AutoSize = true };
            var bFind = new Button { Text = "Google Drive", AutoSize = true };
            var bOff = new Button { Text = "关闭同步", AutoSize = true };
            var bOpen = new Button { Text = "打开", AutoSize = true };
            addRow(new Control[] { bOne, bBrowse, bFind, bOff, bOpen });

            head("Excel 记录文件（自动保存）");
            string logPath = s.LogPath ?? "";
            var tLog = new TextBox { ReadOnly = true, Width = (int)(430 * k) };
            Action showLog = () =>
            {
                string sd = NewSyncDir ?? Store.SyncDir;
                tLog.Text = logPath != "" ? logPath : sd != "" ? Path.Combine(sd, "PomoTodo_Log.xlsx") + "  （在同步文件夹）" : "（请选择位置，或先开启同步文件夹）";
            };
            showLog();
            tl.Controls.Add(tLog); tl.SetColumnSpan(tLog, 2);
            var bLogChoose = new Button { Text = "选择位置...", AutoSize = true };
            var bLogSync = new Button { Text = "放在同步文件夹", AutoSize = true };
            var bLogOpen = new Button { Text = "打开所在文件夹", AutoSize = true };
            addRow(new Control[] { bLogChoose, bLogSync, bLogOpen });
            cExport = addChk("每次有变化时自动更新这个 Excel", s.AutoExport);

            Action<string> setSync = d => { NewSyncDir = d; tSync.Text = d == "" ? "（关闭 - 只存本机）" : d; showLog(); };
            bOne.Click += (a, b) =>
            {
                string o = Store.FindOneDrive();
                if (o == null) { MessageBox.Show(this, "没有找到 OneDrive 文件夹。\n请确认 OneDrive 已登录，或点“选择文件夹...”。", "OneDrive"); return; }
                setSync(Path.Combine(o, "PomoTodo"));
            };
            bFind.Click += (a, b) =>
            {
                string g = Store.FindGoogleDrive();
                if (g == null) { MessageBox.Show(this, "没有找到 Google Drive for desktop。\n请点“选择文件夹...”。", "Google Drive"); return; }
                setSync(Path.Combine(g, "PomoTodo"));
            };
            bBrowse.Click += (a, b) =>
            {
                using (var d = new FolderBrowserDialog { Description = "选择同步文件夹（例如 OneDrive 里的某个文件夹）", ShowNewFolderButton = true })
                {
                    string cur = NewSyncDir ?? Store.SyncDir;
                    string o = cur != "" && Directory.Exists(cur) ? cur : Store.FindOneDrive();
                    if (o != null) d.SelectedPath = o;
                    if (d.ShowDialog(this) == DialogResult.OK) setSync(d.SelectedPath);
                }
            };
            bOff.Click += (a, b) => setSync("");
            bOpen.Click += (a, b) =>
            {
                string d = NewSyncDir ?? Store.SyncDir; if (string.IsNullOrEmpty(d)) d = Store.AppDir;
                try { System.Diagnostics.Process.Start(Directory.Exists(d) ? d : Store.AppDir); } catch { }
            };
            bLogChoose.Click += (a, b) =>
            {
                using (var d = new SaveFileDialog { Filter = "Excel Workbook (*.xlsx)|*.xlsx", FileName = "PomoTodo_Log.xlsx", OverwritePrompt = false, Title = "Excel 记录保存在哪里？" })
                {
                    string init = logPath != "" ? Path.GetDirectoryName(logPath) : (NewSyncDir ?? Store.SyncDir);
                    if (string.IsNullOrEmpty(init) || !Directory.Exists(init)) init = Store.FindOneDrive();
                    if (init != null) d.InitialDirectory = init;
                    if (d.ShowDialog(this) == DialogResult.OK) { logPath = d.FileName; showLog(); cExport.Checked = true; }
                }
            };
            bLogSync.Click += (a, b) => { logPath = ""; showLog(); };
            bLogOpen.Click += (a, b) =>
            {
                string lf = logPath != "" ? logPath : Path.Combine(NewSyncDir ?? Store.SyncDir, "PomoTodo_Log.xlsx");
                string d = Path.GetDirectoryName(lf);
                try { if (File.Exists(lf)) System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + lf + "\""); else if (Directory.Exists(d)) System.Diagnostics.Process.Start(d); } catch { }
            };
            logPathGetter = () => logPath;

            var ok = new Button { Text = "保存", DialogResult = DialogResult.OK, AutoSize = true };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };
            var row = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Bottom, Padding = new Padding(6) };
            row.Controls.Add(cancel); row.Controls.Add(ok);
            var outer = new Panel { Width = (int)(480 * k), Height = (int)(580 * k) };
            outer.Controls.Add(tabs); outer.Controls.Add(row);
            Controls.Add(outer);
            CancelButton = cancel;
            cancel.Click += (a, b) => NewSyncDir = null;
            ok.Click += (a, b) =>
            {
                Result = new AppSettings
                {
                    Focus = (int)nFocus.Value, ShortBreak = (int)nShort.Value, LongBreak = (int)nLong.Value, LongEvery = (int)nEvery.Value,
                    CountMinutes = (int)nCount.Value,
                    AutoStartBreak = cAutoBreak.Checked, AutoStartFocus = cAutoFocus.Checked, Sound = cSound.Checked, TopMost = s.TopMost,
                    CloseToTray = cTray.Checked, StartWithWindows = cStartup.Checked, AutoExport = cExport.Checked, TrayShowsUsed = cUsed.Checked,
                    PopupWhenDone = cPopup.Checked, LogPath = logPathGetter(),
                    SoundStart = soundStart, SoundEnd = soundEnd, ToastStart = cToastStart.Checked, ToastEnd = cToastEnd.Checked, ToastEndStays = cToastStays.Checked,
                    HotkeyStart = hkStart.Text.Trim(), HotkeyShow = hkShow.Text.Trim()
                };
            };
        }
    }

    // ---------------- Main window ----------------
    class MainForm : Form
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, string l);
        [DllImport("user32.dll")] static extern bool FlashWindow(IntPtr h, bool invert);
        [DllImport("user32.dll")] static extern bool DestroyIcon(IntPtr h);
        [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        List<TodoItem> todos; List<RecordItem> records; AppSettings st;
        float k = 1f; int S(int v) { return (int)Math.Round(v * k); }

        static readonly Color Accent = Color.FromArgb(24, 144, 255), RedC = Color.FromArgb(229, 83, 61), GreenC = Color.FromArgb(46, 160, 90),
            BorderC = Color.FromArgb(217, 217, 217), Gray = Color.FromArgb(150, 150, 150), Bg = Color.White, TagC = Color.FromArgb(40, 120, 210);
        static readonly string[] WeekCn = { "周日", "周一", "周二", "周三", "周四", "周五", "周六" };
        Font fUI, fBold, fBig, fClock, fStrike, fSmall, fHead;

        Panel top, bottom, todoPanel, pomoPanel, btnRow, taskRow, progress, clockBox, viewRow;
        Button btnTodo, btnPomo, btnPin, btnStart, btnReset, btnSkip, btnImport, btnExport, btnSettings, btnClearDone, btnByTime, btnByTask;
        Label lblMode, lblBig, lblToday, lblTodoInfo, lblUsed, lblSync;
        TextBox txtAdd, txtDoneSearch; ListBox lbTodos, lbDone, lbHist, lastList; ComboBox cboTask;
        Panel donePanel; Button btnDoneHeader; bool doneExpanded;
        System.Windows.Forms.Timer timer; NotifyIcon tray; ContextMenuStrip trayMenu, todoMenu, histMenu; ToolStripMenuItem miStart;
        Icon appIcon; IntPtr lastHIcon = IntPtr.Zero; string lastIconKey = "";

        string mode = "Focus";         // Focus / Short Break / Long Break
        bool running; DateTime endAt; TimeSpan remaining; DateTime? segStart; double activeSecs; DateTime lastTick;
        int focusDone; TodoItem current; bool loadingList; bool histByTask;
        bool allowShow, reallyExit, trayTipShown, exportPending; DateTime exportAt, nextSyncCheck;

        public MainForm(bool startHidden)
        {
            allowShow = !startHidden;
            st = Store.LoadSettings();
            Store.LoadData(out todos, out records);
            using (var g = CreateGraphics()) k = g.DpiX / 96f;
            string fam = "Microsoft YaHei UI";
            try { using (var t = new Font(fam, 9f)) if (t.Name != fam) fam = "Segoe UI"; } catch { fam = "Segoe UI"; }
            fUI = new Font(fam, 10f); fBold = new Font(fam, 10f, FontStyle.Bold); fSmall = new Font(fam, 8.5f);
            fStrike = new Font(fam, 10f, FontStyle.Strikeout); fHead = new Font(fam, 11f, FontStyle.Bold);
            fBig = new Font("Consolas", 36f, FontStyle.Bold); fClock = new Font("Consolas", 11f, FontStyle.Bold);

            Text = "PomoTodo v" + Program.Version; Font = fUI; BackColor = Bg;
            AutoScaleMode = AutoScaleMode.None;
            ClientSize = new Size(S(380), S(600)); MinimumSize = new Size(S(330), S(480));
            try { appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); Icon = appIcon; } catch { }
            if (appIcon == null) appIcon = SystemIcons.Application;
            TopMost = st.TopMost;
            StartPosition = FormStartPosition.Manual;
            var wa = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(wa.Right - Width - S(12), wa.Bottom - Height - S(12));

            BuildTop(); BuildBottom(); BuildTodo(); BuildPomo(); BuildTray();
            Controls.Add(pomoPanel); Controls.Add(todoPanel); Controls.Add(bottom); Controls.Add(top);

            remaining = Duration(mode);
            timer = new System.Windows.Forms.Timer { Interval = 250 }; timer.Tick += Tick; timer.Start();
            ShowTab(true);
            RefreshTodos(); RefreshRecords(); UpdateTimerUI(); UpdateSyncLabel();
            FormClosing += OnClosing;
            nextSyncCheck = DateTime.Now.AddSeconds(15);
            if (!IsHandleCreated) CreateHandle();
        }

        int Threshold { get { return Math.Max(1, st.CountMinutes); } }
        bool Counted(RecordItem r) { return DataIO.Counted(r, Threshold); }
        static string ModeName(string m) { return m == "Focus" ? "专注" : m == "Long Break" ? "长休息" : "短休息"; }

        // ---------- hotkeys ----------
        protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); RegisterHotkeys(false); }
        void UnregisterHotkeys() { try { UnregisterHotKey(Handle, 1); UnregisterHotKey(Handle, 2); } catch { } }
        void RegisterHotkeys(bool report)
        {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT || st == null) return;
            UnregisterHotkeys();
            var failed = new List<string>();
            uint m; Keys key;
            if (Hotkey.Parse(st.HotkeyStart, out m, out key) && !RegisterHotKey(Handle, 1, m | 0x4000, (uint)key)) failed.Add(st.HotkeyStart);
            if (Hotkey.Parse(st.HotkeyShow, out m, out key) && !RegisterHotKey(Handle, 2, m | 0x4000, (uint)key)) failed.Add(st.HotkeyShow);
            if (report && failed.Count > 0)
                MessageBox.Show(this, "这些快捷键已被其他程序占用，无法使用：\n" + string.Join("\n", failed) + "\n\n请在设置里换一个组合。", "快捷键");
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312)
            {
                int id = m.WParam.ToInt32();
                if (id == 1) HotkeyStartPressed();
                else if (id == 2) { if (Visible && WindowState != FormWindowState.Minimized && ContainsFocus) Hide(); else ShowMain(); }
                return;
            }
            base.WndProc(ref m);
        }
        void HotkeyStartPressed()
        {
            if (running) { Pause(); ShowToastMsg("已暂停", ModeName(mode) + " 剩余 " + MMSS(remaining), Color.Gray, 2500); return; }
            if (segStart != null) { StartTimer(); ShowToastMsg("继续" + ModeName(mode), "剩余 " + MMSS(remaining), mode == "Focus" ? RedC : GreenC, 2500); return; }
            if (mode != "Focus") { mode = "Focus"; remaining = Duration(mode); }
            StartTimer();
        }

        protected override void SetVisibleCore(bool value)
        {
            if (!allowShow) { value = false; if (!IsHandleCreated) CreateHandle(); }
            base.SetVisibleCore(value);
        }
        public void ShowMain()
        {
            allowShow = true;
            Show();
            if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
            bool tm = TopMost; TopMost = true; Activate(); TopMost = tm;
        }

        Button FlatBtn(string text, int w)
        {
            var b = new Button { Text = text, FlatStyle = FlatStyle.Flat, BackColor = Bg, Width = S(w), Height = S(28), Cursor = Cursors.Hand, UseVisualStyleBackColor = false, TabStop = false };
            b.FlatAppearance.BorderColor = BorderC; b.FlatAppearance.MouseOverBackColor = Color.FromArgb(240, 247, 255);
            return b;
        }

        void BuildTray()
        {
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add(new ToolStripMenuItem("显示 PomoTodo", null, (a, b) => ShowMain()) { Font = fBold });
            trayMenu.Items.Add(new ToolStripSeparator());
            miStart = new ToolStripMenuItem("开始番茄", null, (a, b) => HotkeyStartPressed());
            trayMenu.Items.Add(miStart);
            trayMenu.Items.Add("跳过", null, (a, b) => SkipTimer());
            trayMenu.Items.Add("重置", null, (a, b) => ResetTimer());
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("导出 Excel...", null, (a, b) => { ShowMain(); DoExport(); });
            trayMenu.Items.Add("立即同步", null, (a, b) => SyncNow(true));
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("退出", null, (a, b) => { reallyExit = true; Close(); });
            tray = new NotifyIcon { Icon = appIcon, Text = "PomoTodo", Visible = true, ContextMenuStrip = trayMenu };
            tray.MouseClick += (a, e) =>
            {
                if (e.Button != MouseButtons.Left) return;
                if (Visible && WindowState != FormWindowState.Minimized) Hide(); else { ShowMain(); ShowTab(false); }
            };
        }

        void BuildTop()
        {
            top = new Panel { Dock = DockStyle.Top, Height = S(46), BackColor = Bg };
            top.Paint += (a, e) => e.Graphics.DrawLine(new Pen(BorderC), 0, top.Height - 1, top.Width, top.Height - 1);
            btnTodo = FlatBtn("任务", 60); btnPomo = FlatBtn("番茄", 60);
            btnTodo.Click += (a, b) => ShowTab(true); btnPomo.Click += (a, b) => ShowTab(false);
            clockBox = new Panel { Size = new Size(S(92), S(30)), Cursor = Cursors.Hand, BackColor = Bg };
            clockBox.Paint += (a, e) =>
            {
                var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
                Color c = mode == "Focus" ? RedC : GreenC;
                bool idle = !running && segStart == null;
                Color fc = idle ? Color.FromArgb(60, 60, 60) : c;
                using (var p = new Pen(BorderC)) g.DrawRectangle(p, 0, 0, clockBox.Width - 1, clockBox.Height - 1);
                TextRenderer.DrawText(g, MMSS(remaining), fClock, new Rectangle(S(6), 0, S(58), clockBox.Height), fc, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
                int cx = clockBox.Width - S(18), cy = clockBox.Height / 2, h = S(6);
                using (var br = new SolidBrush(fc))
                {
                    if (running) { g.FillRectangle(br, cx - S(5), cy - h, S(3), h * 2); g.FillRectangle(br, cx + S(1), cy - h, S(3), h * 2); }
                    else g.FillPolygon(br, new[] { new Point(cx - S(4), cy - h), new Point(cx - S(4), cy + h), new Point(cx + S(6), cy) });
                }
            };
            clockBox.Click += (a, b) => HotkeyStartPressed();
            btnPin = FlatBtn("置顶", 50);
            btnPin.Click += (a, b) => { st.TopMost = !st.TopMost; TopMost = st.TopMost; UpdatePin(); Store.SaveSettings(st); };
            top.Controls.AddRange(new Control[] { btnTodo, btnPomo, clockBox, btnPin });
            top.Resize += (a, b) => LayoutTop(); LayoutTop(); UpdatePin();
        }
        void LayoutTop()
        {
            int y = (top.Height - S(28)) / 2;
            btnTodo.Location = new Point(S(10), y); btnPomo.Location = new Point(btnTodo.Right - 1, y);
            btnPin.Location = new Point(top.Width - btnPin.Width - S(8), y);
            clockBox.Location = new Point(btnPin.Left - clockBox.Width - S(6), (top.Height - clockBox.Height) / 2);
        }
        void UpdatePin() { btnPin.ForeColor = st.TopMost ? Accent : Color.Black; btnPin.FlatAppearance.BorderColor = st.TopMost ? Accent : BorderC; btnPin.Width = TextRenderer.MeasureText(btnPin.Text, btnPin.Font).Width + S(18); LayoutTop(); }

        void BuildBottom()
        {
            bottom = new Panel { Dock = DockStyle.Bottom, Height = S(40), BackColor = Bg };
            bottom.Paint += (a, e) => e.Graphics.DrawLine(new Pen(BorderC), 0, 0, bottom.Width, 0);
            btnImport = FlatBtn("导入", 56); btnExport = FlatBtn("导出 Excel", 92); btnSettings = FlatBtn("设置", 56);
            lblSync = new Label { AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, Font = fSmall, Cursor = Cursors.Hand };
            lblSync.Click += (a, b) => { if (Store.IsSynced) SyncNow(true); else DoSettings(); };
            btnImport.Click += (a, b) => DoImport(); btnExport.Click += (a, b) => DoExport(); btnSettings.Click += (a, b) => DoSettings();
            bottom.Controls.AddRange(new Control[] { btnImport, btnExport, lblSync, btnSettings });
            bottom.Resize += (a, b) =>
            {
                int y = (bottom.Height - S(28)) / 2;
                btnImport.Location = new Point(S(8), y); btnExport.Location = new Point(btnImport.Right + S(6), y);
                btnSettings.Location = new Point(bottom.Width - btnSettings.Width - S(8), y);
                lblSync.Bounds = new Rectangle(btnExport.Right + S(2), y, Math.Max(0, btnSettings.Left - btnExport.Right - S(4)), S(28));
            };
        }
        void UpdateSyncLabel()
        {
            if (!Store.IsSynced) { lblSync.Text = "未同步"; lblSync.ForeColor = Gray; }
            else if (!Store.SyncAvailable || Store.LastSyncError != null) { lblSync.Text = "同步出错"; lblSync.ForeColor = RedC; }
            else { lblSync.Text = "已同步"; lblSync.ForeColor = GreenC; }
        }

        // ---------- Todo tab ----------
        static readonly System.Text.RegularExpressions.Regex TargetRx = new System.Text.RegularExpressions.Regex(@"^(.*?)\s*(?:/\s*(\d+)|[\(（]\s*(\d+)\s*[\)）])\s*$");
        static void ParseTask(string input, out string text, out int target)
        {
            text = input.Trim(); target = 0;
            var m = TargetRx.Match(text);
            if (m.Success && m.Groups[1].Value.Trim() != "")
            {
                text = m.Groups[1].Value.Trim();
                int.TryParse(m.Groups[2].Success && m.Groups[2].Value != "" ? m.Groups[2].Value : m.Groups[3].Value, out target);
            }
        }

        void BuildTodo()
        {
            todoPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(S(10)), BackColor = Bg };
            txtAdd = new TextBox { Dock = DockStyle.Top, Font = new Font(fUI.FontFamily, 11f) };
            txtAdd.HandleCreated += (a, b) => { try { SendMessage(txtAdd.Handle, 0x1501, (IntPtr)1, "添加新任务（末尾加 /10 = 预计 10 个番茄）"); } catch { } };
            txtAdd.KeyDown += (a, e) =>
            {
                if (e.KeyCode != Keys.Enter) return;
                e.SuppressKeyPress = true;
                string text; int target; ParseTask(txtAdd.Text, out text, out target);
                if (text != "") { todos.Add(new TodoItem { Text = text, Target = target }); txtAdd.Clear(); SaveAll(); RefreshTodos(); }
            };
            var spacer = new Panel { Dock = DockStyle.Top, Height = S(8) };
            lbTodos = MakeTodoList();
            lastList = lbTodos;
            todoMenu = new ContextMenuStrip();
            todoMenu.Items.Add("开始番茄（用这个任务）", null, (a, b) => { var t = SelTodo(); if (t != null) StartWith(t); });
            todoMenu.Items.Add("设为当前任务", null, (a, b) => { var t = SelTodo(); if (t != null) { current = t; RefreshTodos(); } });
            todoMenu.Items.Add("设置番茄数（已完成 / 预计）...", null, (a, b) => EditCounts());
            todoMenu.Items.Add("重命名 (F2)", null, (a, b) => RenameSel());
            todoMenu.Items.Add("删除 (Del)", null, (a, b) => DeleteSel());

            var info = new Panel { Dock = DockStyle.Bottom, Height = S(30) };
            lblTodoInfo = new Label { Dock = DockStyle.Fill, ForeColor = Gray, Font = fSmall, TextAlign = ContentAlignment.MiddleLeft, Text = "" };
            btnClearDone = FlatBtn("清除已完成", 90); btnClearDone.Dock = DockStyle.Right; btnClearDone.Font = fSmall;
            btnClearDone.Click += (a, b) =>
            {
                if (!todos.Any(t => t.Done)) return;
                if (MessageBox.Show(this, "从列表中移除所有已完成的任务？\n（番茄记录会保留）", "PomoTodo", MessageBoxButtons.OKCancel) == DialogResult.OK)
                { todos.RemoveAll(t => t.Done); SaveAll(); RefreshTodos(); }
            };
            info.Controls.Add(lblTodoInfo); info.Controls.Add(btnClearDone);
            var hint = new Label { Dock = DockStyle.Bottom, Height = S(20), ForeColor = Gray, Font = fSmall, Text = "双击 = 开始番茄  ·  点圆圈 = 完成  ·  右键 = 更多" };

            // ---- "已完成" section: folded by default; ticked tasks move here. Has its own search box.
            donePanel = new Panel { Dock = DockStyle.Bottom, Padding = new Padding(0, S(6), 0, 0) };
            btnDoneHeader = new Button { Dock = DockStyle.Top, Height = S(30), FlatStyle = FlatStyle.Flat, TextAlign = ContentAlignment.MiddleLeft, Font = fBold, BackColor = Color.FromArgb(245, 245, 245), Cursor = Cursors.Hand };
            btnDoneHeader.FlatAppearance.BorderColor = BorderC;
            var doneMenu = new ContextMenuStrip();
            doneMenu.Items.Add("按完成时间排序", null, (a, b) => { doneSort = DoneSort.ByTime; RefreshDone(); });
            doneMenu.Items.Add("按完成数量排序", null, (a, b) => { doneSort = DoneSort.ByPomos; RefreshDone(); });
            btnDoneHeader.Click += (a, b) =>
            {
                if (ModifierKeys == Keys.Alt) doneMenu.Show(btnDoneHeader, new Point(btnDoneHeader.Width - S(20), btnDoneHeader.Height));
                else { doneExpanded = !doneExpanded; LayoutDone(); if (doneExpanded) txtDoneSearch.Focus(); }
            };
            txtDoneSearch = new TextBox { Dock = DockStyle.Top, Font = fUI };
            txtDoneSearch.HandleCreated += (a, b) => { try { SendMessage(txtDoneSearch.Handle, 0x1501, (IntPtr)1, "搜索已完成的任务（任务名 或 日期，如 2026-09）"); } catch { } };
            txtDoneSearch.TextChanged += (a, b) => RefreshDone();
            txtDoneSearch.KeyDown += (a, e) => { if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; txtDoneSearch.Clear(); } };
            var searchGap = new Panel { Dock = DockStyle.Top, Height = S(4) };
            lbDone = MakeTodoList();
            donePanel.Controls.Add(lbDone); donePanel.Controls.Add(searchGap); donePanel.Controls.Add(txtDoneSearch); donePanel.Controls.Add(btnDoneHeader);
            todoPanel.Resize += (a, b) => LayoutDone();

            todoPanel.Controls.Add(lbTodos); todoPanel.Controls.Add(donePanel); todoPanel.Controls.Add(hint); todoPanel.Controls.Add(info); todoPanel.Controls.Add(spacer); todoPanel.Controls.Add(txtAdd);
            LayoutDone();
        }

        // Owner-drawn task list (used for the open list and for the "已完成" list)
        ListBox MakeTodoList()
        {
            var lb = new ListBox { Dock = DockStyle.Fill, DrawMode = DrawMode.OwnerDrawVariable, IntegralHeight = false, BorderStyle = BorderStyle.FixedSingle, Font = fUI };
            lb.MeasureItem += (a, e) => { var t = lb.Items[e.Index] as TodoItem; e.ItemHeight = t != null && (t.Pomos > 0 || t.Target > 0) ? S(70) : S(52); };
            lb.DrawItem += DrawTodo;
            lb.Enter += (a, b) => lastList = lb;
            lb.MouseDown += (a, e) =>
            {
                lastList = lb;
                int i = lb.IndexFromPoint(e.Location);
                if (i < 0 || i >= lb.Items.Count) return;
                lb.SelectedIndex = i;
                var t = lb.Items[i] as TodoItem; if (t == null) return;
                if (e.Button == MouseButtons.Left && e.X < S(40) && e.Clicks == 1)
                {
                    t.Done = !t.Done; t.DoneAt = t.Done ? (DateTime?)DateTime.Now : null;
                    SaveAll(); BeginInvoke(new Action(RefreshTodos));
                }
                else if (e.Button == MouseButtons.Right) todoMenu.Show(lb, e.Location);
            };
            lb.MouseDoubleClick += (a, e) => { if (e.X >= S(40)) { var t = SelTodo(); if (t != null && !t.Done) StartWith(t); } };
            lb.KeyDown += (a, e) => { if (e.KeyCode == Keys.Delete) DeleteSel(); else if (e.KeyCode == Keys.F2) RenameSel(); else if (e.KeyCode == Keys.Enter) { var t = SelTodo(); if (t != null && !t.Done) StartWith(t); } };
            lb.Resize += (a, b) => lb.Invalidate();
            return lb;
        }

        // Folded: only the header bar. Open: about half of the task area.
        void LayoutDone()
        {
            if (donePanel == null) return;
            int n = todos.Count(t => t.Done);
            string sortInfo = doneSort == DoneSort.ByPomos ? " [按数量]" : " [按时间]";
            btnDoneHeader.Text = (doneExpanded ? "▾  " : "▸  ") + "已完成 (" + n + ")" + sortInfo + (doneExpanded ? "  (Alt+Click改排序)" : "   点击展开 / Alt+Click排序");
            txtDoneSearch.Visible = doneExpanded; lbDone.Visible = doneExpanded;
            int head = btnDoneHeader.Height + donePanel.Padding.Top;
            int h = doneExpanded ? Math.Max(head + S(140), (todoPanel.ClientSize.Height - txtAdd.Height - S(60)) / 2) : head;
            if (donePanel.Height != h) donePanel.Height = h;
            if (!doneExpanded && lastList == lbDone) lastList = lbTodos;
        }

        enum DoneSort { ByTime, ByPomos }
        DoneSort doneSort = DoneSort.ByTime;

        static bool DoneMatches(TodoItem t, string q)
        {
            if (q == "") return true;
            var ci = StringComparison.OrdinalIgnoreCase;
            return (t.Text ?? "").IndexOf(q, ci) >= 0
                || t.Created.ToString("yyyy-MM-dd HH:mm", Store.IC).IndexOf(q, ci) >= 0
                || (t.DoneAt.HasValue && t.DoneAt.Value.ToString("yyyy-MM-dd HH:mm", Store.IC).IndexOf(q, ci) >= 0);
        }
        void RefreshDone()
        {
            if (lbDone == null) return;
            var sel = lbDone.SelectedItem as TodoItem;
            string q = (txtDoneSearch.Text ?? "").Trim();
            lbDone.BeginUpdate(); lbDone.Items.Clear();
            var sorted = todos.Where(x => x.Done && DoneMatches(x, q));
            if (doneSort == DoneSort.ByPomos)
                sorted = sorted.OrderByDescending(x => x.Pomos).ThenByDescending(x => x.DoneAt ?? x.Created);
            else
                sorted = sorted.OrderByDescending(x => x.DoneAt ?? x.Created);
            foreach (var t in sorted)
                lbDone.Items.Add(t);
            if (sel != null && lbDone.Items.Contains(sel)) lbDone.SelectedItem = sel;
            if (q != "" && lbDone.Items.Count == 0) lbDone.Items.Add("没有找到匹配的已完成任务");
            lbDone.EndUpdate();
            LayoutDone();
        }
        static string When(DateTime d) { return d.ToString(d.Year == DateTime.Now.Year ? "MM-dd HH:mm" : "yyyy-MM-dd HH:mm", Store.IC); }

        // draw text with #tags in blue
        int DrawTagged(Graphics g, string text, Font f, int x, int y, int maxX, Color normal)
        {
            var flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix;
            var tokens = System.Text.RegularExpressions.Regex.Split(text, @"(\s+)");
            foreach (var tok in tokens)
            {
                if (tok == "") continue;
                bool tag = tok.StartsWith("#") && normal != Gray;
                var sz = TextRenderer.MeasureText(g, tok, f, new Size(10000, 100), flags);
                if (x + sz.Width > maxX)
                {
                    TextRenderer.DrawText(g, tok, f, new Rectangle(x, y, Math.Max(0, maxX - x), sz.Height), tag ? TagC : normal, flags | TextFormatFlags.EndEllipsis);
                    return maxX;
                }
                TextRenderer.DrawText(g, tok, f, new Point(x, y), tag ? TagC : normal, flags);
                x += sz.Width;
            }
            return x;
        }

        void DrawTodo(object sender, DrawItemEventArgs e)
        {
            var lb = (ListBox)sender;
            if (e.Index < 0 || e.Index >= lb.Items.Count) return;
            var g = e.Graphics; var r = e.Bounds;
            var t = lb.Items[e.Index] as TodoItem;
            if (t == null)
            {
                using (var bg0 = new SolidBrush(Color.White)) g.FillRectangle(bg0, r);
                TextRenderer.DrawText(g, Convert.ToString(lb.Items[e.Index]), fSmall, new Rectangle(r.Left + S(12), r.Top, r.Width - S(24), r.Height), Gray, TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
                return;
            }
            bool sel = (e.State & DrawItemState.Selected) != 0;
            using (var bg = new SolidBrush(sel ? Color.FromArgb(232, 243, 255) : Color.White)) g.FillRectangle(bg, r);
            using (var p = new Pen(Color.FromArgb(238, 238, 238))) g.DrawLine(p, r.Left, r.Bottom - 1, r.Right, r.Bottom - 1);
            if (t == current && !t.Done) using (var b = new SolidBrush(RedC)) g.FillRectangle(b, r.Left, r.Top + S(6), S(3), r.Height - S(12));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int cs = S(18), cx = r.Left + S(12), cy = r.Top + S(9);
            if (t.Done)
            {
                using (var b = new SolidBrush(Color.FromArgb(190, 190, 190))) g.FillEllipse(b, cx, cy, cs, cs);
                using (var p = new Pen(Color.White, Math.Max(1.5f, 2 * k)))
                    g.DrawLines(p, new[] { new PointF(cx + cs * 0.27f, cy + cs * 0.52f), new PointF(cx + cs * 0.44f, cy + cs * 0.7f), new PointF(cx + cs * 0.75f, cy + cs * 0.32f) });
            }
            else using (var p = new Pen(Color.FromArgb(185, 185, 185), Math.Max(1.2f, 1.5f * k))) g.DrawEllipse(p, cx, cy, cs, cs);
            g.SmoothingMode = SmoothingMode.Default;
            int tx = r.Left + S(40), maxX = r.Right - S(8);
            DrawTagged(g, t.Text, t.Done ? fStrike : (t == current ? fBold : fUI), tx, r.Top + S(8), maxX, t.Done ? Gray : Color.FromArgb(30, 30, 30));
            if (t.Pomos > 0 || t.Target > 0)
            {
                string sub = t.Target > 0 ? "(" + t.Pomos + "/" + t.Target + ")" : "(" + t.Pomos + ")";
                Color sc = Gray;
                if (t.Target > 0 && t.Pomos >= t.Target && !t.Done) { sub += "  已达到预计"; sc = GreenC; }
                else if (t.Target > 0 && !t.Done) sub += "  还差 " + (t.Target - t.Pomos) + " 个";
                if (t == current && running && mode == "Focus") { sub += "  · 进行中"; sc = RedC; }
                TextRenderer.DrawText(g, sub, fSmall, new Point(tx, r.Top + S(31)), sc, TextFormatFlags.NoPadding);
                if (t.Target > 0)
                {
                    int bw = S(60), bx = maxX - bw, by = r.Top + S(36);
                    using (var b1 = new SolidBrush(Color.FromArgb(236, 236, 236))) g.FillRectangle(b1, bx, by, bw, S(4));
                    using (var b2 = new SolidBrush(t.Pomos >= t.Target ? GreenC : RedC)) g.FillRectangle(b2, bx, by, (int)(bw * Math.Min(1.0, (double)t.Pomos / t.Target)), S(4));
                }
            }
            // creation / completion time
            string times = "创建 " + When(t.Created) + (t.Done && t.DoneAt.HasValue ? "   ·   完成 " + When(t.DoneAt.Value) : "");
            int ty = r.Top + (t.Pomos > 0 || t.Target > 0 ? S(49) : S(31));
            TextRenderer.DrawText(g, times, fSmall, new Rectangle(tx, ty, Math.Max(0, maxX - tx), S(18)), Color.FromArgb(150, 150, 150), TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }

        // ---------- Pomo tab ----------
        class TaskChoice { public TodoItem Item; public override string ToString() { return Item == null ? "（不选任务）" : Item.Text; } }
        class HistRow
        {
            public int Kind; // 0 header, 1 session, 2 task summary
            public DateTime Day; public string Left = "", Text = "", Right = ""; public RecordItem Rec; public bool Counted = true;
        }

        void BuildPomo()
        {
            pomoPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(S(10), S(4), S(10), S(6)), BackColor = Bg };
            lblMode = new Label { Dock = DockStyle.Top, Height = S(24), TextAlign = ContentAlignment.MiddleCenter, Font = fBold };
            lblBig = new Label { Dock = DockStyle.Top, Height = S(60), TextAlign = ContentAlignment.MiddleCenter, Font = fBig };
            progress = new Panel { Dock = DockStyle.Top, Height = S(6) };
            progress.Paint += (a, e) =>
            {
                double total = Duration(mode).TotalSeconds, used = Math.Max(0, total - Math.Max(0, remaining.TotalSeconds));
                int w = progress.Width - S(40), x = S(20);
                using (var b1 = new SolidBrush(Color.FromArgb(236, 236, 236))) e.Graphics.FillRectangle(b1, x, 0, w, progress.Height);
                using (var b2 = new SolidBrush(mode == "Focus" ? RedC : GreenC)) e.Graphics.FillRectangle(b2, x, 0, (int)(w * (total > 0 ? used / total : 0)), progress.Height);
            };
            lblUsed = new Label { Dock = DockStyle.Top, Height = S(24), TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.FromArgb(90, 90, 90), Font = fSmall };
            taskRow = new Panel { Dock = DockStyle.Top, Height = S(32) };
            var lt = new Label { Text = "任务：", AutoSize = false, Width = S(48), Dock = DockStyle.Left, TextAlign = ContentAlignment.MiddleLeft };
            cboTask = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
            cboTask.SelectedIndexChanged += (a, b) => { if (!loadingList) { current = cboTask.SelectedItem is TaskChoice ? ((TaskChoice)cboTask.SelectedItem).Item : null; UpdateTimerUI(); lbTodos.Invalidate(); } };
            var tWrap = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, S(3), 0, 0) }; tWrap.Controls.Add(cboTask);
            taskRow.Controls.Add(tWrap); taskRow.Controls.Add(lt);

            btnRow = new Panel { Dock = DockStyle.Top, Height = S(44) };
            btnStart = FlatBtn("开始", 90); btnReset = FlatBtn("重置", 70); btnSkip = FlatBtn("跳过", 70);
            foreach (var b in new[] { btnStart, btnReset, btnSkip }) b.Height = S(32);
            btnStart.Font = fBold;
            btnStart.Click += (a, b) => { if (running) Pause(); else StartTimer(); };
            btnReset.Click += (a, b) => ResetTimer();
            btnSkip.Click += (a, b) => SkipTimer();
            btnRow.Controls.AddRange(new Control[] { btnStart, btnReset, btnSkip });
            btnRow.Resize += (a, b) =>
            {
                int total = btnStart.Width + btnReset.Width + btnSkip.Width + S(16);
                int x = (btnRow.Width - total) / 2, y = S(6);
                btnStart.Location = new Point(x, y); btnReset.Location = new Point(btnStart.Right + S(8), y); btnSkip.Location = new Point(btnReset.Right + S(8), y);
            };

            viewRow = new Panel { Dock = DockStyle.Top, Height = S(34) };
            lblToday = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.FromArgb(50, 50, 50), Font = fBold };
            btnByTime = FlatBtn("按时间", 56); btnByTask = FlatBtn("按任务", 56);
            btnByTime.Font = fSmall; btnByTask.Font = fSmall; btnByTime.Height = btnByTask.Height = S(26);
            var vbtns = new Panel { Dock = DockStyle.Right, Width = btnByTime.Width + btnByTask.Width };
            btnByTime.Location = new Point(0, S(4)); btnByTask.Location = new Point(btnByTime.Width - 1, S(4));
            vbtns.Controls.Add(btnByTime); vbtns.Controls.Add(btnByTask);
            btnByTime.Click += (a, b) => { histByTask = false; RefreshRecords(); };
            btnByTask.Click += (a, b) => { histByTask = true; RefreshRecords(); };
            viewRow.Controls.Add(lblToday); viewRow.Controls.Add(vbtns);

            lbHist = new ListBox { Dock = DockStyle.Fill, DrawMode = DrawMode.OwnerDrawVariable, IntegralHeight = false, BorderStyle = BorderStyle.FixedSingle };
            lbHist.MeasureItem += (a, e) => { var h = lbHist.Items[e.Index] as HistRow; e.ItemHeight = h != null && h.Kind == 0 ? S(36) : S(28); };
            lbHist.DrawItem += DrawHist;
            lbHist.Resize += (a, b) => lbHist.Invalidate();
            histMenu = new ContextMenuStrip();
            histMenu.Items.Add("删除这条记录", null, (a, b) =>
            {
                var h = lbHist.SelectedItem as HistRow;
                if (h == null || h.Rec == null) return;
                if (MessageBox.Show(this, "删除这条番茄记录？", "PomoTodo", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    var t = FindTodoFor(h.Rec);
                    if (t != null && Counted(h.Rec) && t.Pomos > 0) t.Pomos--;
                    records.Remove(h.Rec); SaveAll(); RefreshRecords(); RefreshTodos();
                }
            });
            lbHist.MouseDown += (a, e) =>
            {
                int i = lbHist.IndexFromPoint(e.Location);
                if (i < 0) return;
                lbHist.SelectedIndex = i;
                var h = lbHist.Items[i] as HistRow;
                if (e.Button == MouseButtons.Right && h != null && h.Rec != null) histMenu.Show(lbHist, e.Location);
            };

            pomoPanel.Controls.Add(lbHist); pomoPanel.Controls.Add(viewRow); pomoPanel.Controls.Add(btnRow);
            pomoPanel.Controls.Add(taskRow); pomoPanel.Controls.Add(lblUsed); pomoPanel.Controls.Add(progress);
            pomoPanel.Controls.Add(lblBig); pomoPanel.Controls.Add(lblMode);
        }

        TodoItem FindTodoFor(RecordItem r)
        {
            if (!string.IsNullOrEmpty(r.TaskId)) { var t = todos.FirstOrDefault(x => x.Id == r.TaskId); if (t != null) return t; }
            return todos.FirstOrDefault(x => x.Text == r.Task);
        }

        void DrawHist(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= lbHist.Items.Count) return;
            var h = lbHist.Items[e.Index] as HistRow; if (h == null) return;
            var g = e.Graphics; var r = e.Bounds;
            bool sel = (e.State & DrawItemState.Selected) != 0 && h.Kind != 0;
            using (var bg = new SolidBrush(sel ? Color.FromArgb(232, 243, 255) : Color.White)) g.FillRectangle(bg, r);
            var flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix;
            int pad = S(10);
            if (h.Kind == 0)
            {
                using (var p = new Pen(Color.FromArgb(225, 225, 225))) g.DrawLine(p, r.Left, r.Top, r.Right, r.Top);
                string day = h.Day.ToString("MM-dd") + "  " + (h.Day == DateTime.Today ? "今天" : WeekCn[(int)h.Day.DayOfWeek]);
                TextRenderer.DrawText(g, day, fHead, new Rectangle(r.Left + pad, r.Top, r.Width, r.Height), Color.FromArgb(30, 30, 30), flags);
                var rs = TextRenderer.MeasureText(g, h.Right, fSmall, new Size(1000, 100), flags);
                TextRenderer.DrawText(g, h.Right, fSmall, new Rectangle(r.Right - pad - rs.Width, r.Top, rs.Width + 2, r.Height), Gray, flags);
                return;
            }
            var rsz = TextRenderer.MeasureText(g, h.Right, fSmall, new Size(1000, 100), flags);
            int rightX = r.Right - pad - rsz.Width;
            TextRenderer.DrawText(g, h.Right, fSmall, new Rectangle(rightX, r.Top, rsz.Width + 2, r.Height), h.Counted ? Gray : Color.FromArgb(200, 140, 60), flags);
            int x = r.Left + pad;
            if (h.Kind == 1)
            {
                TextRenderer.DrawText(g, h.Left, fSmall, new Rectangle(x, r.Top, S(90), r.Height), Gray, flags);
                x += S(92);
            }
            var tsz = TextRenderer.MeasureText(g, "Ag", fUI, new Size(1000, 100), TextFormatFlags.NoPadding);
            if (h.Kind == 2 && h.Rec == null && h.Text == "" && h.Right == "还没有番茄记录") return;
            DrawTagged(g, h.Text == "" ? "（未选任务）" : h.Text, fUI, x, r.Top + (r.Height - tsz.Height) / 2, rightX - S(8), h.Text == "" ? Gray : (h.Counted ? Color.FromArgb(30, 30, 30) : Gray));
        }

        void ShowTab(bool todo)
        {
            todoPanel.Visible = todo; pomoPanel.Visible = !todo;
            btnTodo.ForeColor = todo ? Accent : Color.Black; btnPomo.ForeColor = todo ? Color.Black : Accent;
            btnTodo.FlatAppearance.BorderColor = todo ? Accent : BorderC; btnPomo.FlatAppearance.BorderColor = todo ? BorderC : Accent;
            if (todo) txtAdd.Focus(); else RefreshTaskCombo();
        }

        TodoItem SelTodo() { return ((lastList != null && lastList.Visible ? lastList : lbTodos).SelectedItem) as TodoItem; }
        void DeleteSel()
        {
            var t = SelTodo(); if (t == null) return;
            if (MessageBox.Show(this, "删除任务“" + t.Text + "”？\n（番茄记录会保留）", "PomoTodo", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
            todos.Remove(t); if (current == t) current = null; SaveAll(); RefreshTodos();
        }
        void RenameSel()
        {
            var t = SelTodo(); if (t == null) return;
            string s = Prompt("重命名任务", t.Text);
            if (string.IsNullOrWhiteSpace(s)) return;
            string text; int target; ParseTask(s, out text, out target);
            string old = t.Text; t.Text = text; if (target > 0) t.Target = target;
            foreach (var r in records.Where(r => r.TaskId == t.Id || (r.TaskId == "" && r.Task == old))) { r.Task = t.Text; r.TaskId = t.Id; }
            SaveAll(); RefreshTodos(); RefreshRecords();
        }
        void EditCounts()
        {
            var t = SelTodo(); if (t == null) return;
            using (var f = new Form { Text = "番茄数 - " + t.Text, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, ClientSize = new Size(S(300), S(130)), MaximizeBox = false, MinimizeBox = false, Font = fUI, ShowInTaskbar = false })
            {
                var l1 = new Label { Text = "已完成番茄数", Location = new Point(S(14), S(18)), AutoSize = true };
                var n1 = new NumericUpDown { Minimum = 0, Maximum = 100000, Value = t.Pomos, Location = new Point(S(170), S(14)), Width = S(110) };
                var l2 = new Label { Text = "预计需要（0 = 不设置）", Location = new Point(S(14), S(52)), AutoSize = true };
                var n2 = new NumericUpDown { Minimum = 0, Maximum = 100000, Value = t.Target, Location = new Point(S(170), S(48)), Width = S(110) };
                var ok = new Button { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(S(200), S(90)), Size = new Size(S(80), S(30)) };
                f.Controls.AddRange(new Control[] { l1, n1, l2, n2, ok }); f.AcceptButton = ok;
                if (f.ShowDialog(this) != DialogResult.OK) return;
                t.Pomos = (int)n1.Value; t.Target = (int)n2.Value;
                SaveAll(); RefreshTodos();
            }
        }
        string Prompt(string title, string val)
        {
            using (var f = new Form { Text = title, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, ClientSize = new Size(S(320), S(84)), MaximizeBox = false, MinimizeBox = false, Font = fUI, ShowInTaskbar = false })
            {
                var tb = new TextBox { Text = val, Location = new Point(S(10), S(10)), Width = S(300) };
                var ok = new Button { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(S(230), S(46)), Size = new Size(S(80), S(30)) };
                f.Controls.Add(tb); f.Controls.Add(ok); f.AcceptButton = ok;
                return f.ShowDialog(this) == DialogResult.OK ? tb.Text : null;
            }
        }

        void RefreshTodos()
        {
            loadingList = true;
            var sel = SelTodo();
            lbTodos.BeginUpdate(); lbTodos.Items.Clear();
            foreach (var t in todos.Where(x => !x.Done))
                lbTodos.Items.Add(t);
            if (sel != null && lbTodos.Items.Contains(sel)) lbTodos.SelectedItem = sel;
            lbTodos.EndUpdate();
            loadingList = false;
            lblTodoInfo.Text = todos.Count(t => !t.Done) + " 个进行中  ·  " + todos.Count(t => t.Done) + " 个已完成";
            RefreshDone();
            RefreshTaskCombo();
        }
        void RefreshTaskCombo()
        {
            loadingList = true;
            cboTask.Items.Clear();
            cboTask.Items.Add(new TaskChoice());
            if (current != null && (current.Done || !todos.Contains(current))) current = null;
            int sel = 0;
            foreach (var t in todos.Where(x => !x.Done))
            {
                cboTask.Items.Add(new TaskChoice { Item = t });
                if (t == current) sel = cboTask.Items.Count - 1;
            }
            cboTask.SelectedIndex = sel;
            loadingList = false;
        }
        void RefreshRecords()
        {
            var focus = records.Where(r => r.Type == "Focus").ToList();
            lbHist.BeginUpdate(); lbHist.Items.Clear();
            foreach (var day in focus.GroupBy(r => r.Start.Date).OrderByDescending(g => g.Key).Take(120))
            {
                int cnt = day.Count(Counted);
                double mins = day.Sum(r => r.Minutes);
                lbHist.Items.Add(new HistRow { Kind = 0, Day = day.Key, Right = "完成了 " + cnt + " 个番茄 · " + Math.Round(mins) + " 分钟" });
                if (histByTask)
                {
                    foreach (var tg in day.GroupBy(r => r.Task ?? "").OrderByDescending(g => g.Count(Counted)).ThenByDescending(g => g.Sum(r => r.Minutes)))
                    {
                        int c = tg.Count(Counted);
                        lbHist.Items.Add(new HistRow { Kind = 2, Day = day.Key, Text = tg.Key, Right = c + " 个番茄 · " + Math.Round(tg.Sum(r => r.Minutes)) + " 分钟", Counted = c > 0 });
                    }
                }
                else
                {
                    foreach (var r in day.OrderByDescending(x => x.Start))
                    {
                        bool c = Counted(r);
                        lbHist.Items.Add(new HistRow { Kind = 1, Day = day.Key, Rec = r, Left = r.Start.ToString("HH:mm") + "-" + r.End.ToString("HH:mm"), Text = r.Task ?? "", Counted = c, Right = c ? Math.Round(r.Minutes) + " 分钟" : Math.Round(r.Minutes, 1) + " 分钟 · 未计入" });
                    }
                }
            }
            if (lbHist.Items.Count == 0) lbHist.Items.Add(new HistRow { Kind = 2, Text = "", Right = "还没有番茄记录", Counted = true });
            lbHist.EndUpdate();
            var today = focus.Where(r => r.Start.Date == DateTime.Today).ToList();
            lblToday.Text = "今天：" + today.Count(Counted) + " 个番茄 · " + Math.Round(today.Sum(r => r.Minutes)) + " 分钟";
            btnByTime.ForeColor = histByTask ? Color.Black : Accent; btnByTask.ForeColor = histByTask ? Accent : Color.Black;
            btnByTime.FlatAppearance.BorderColor = histByTask ? BorderC : Accent; btnByTask.FlatAppearance.BorderColor = histByTask ? Accent : BorderC;
        }

        // ---------------- Timer logic ----------------
        TimeSpan Duration(string m)
        {
            return TimeSpan.FromMinutes(m == "Focus" ? st.Focus : m == "Long Break" ? st.LongBreak : st.ShortBreak);
        }
        void StartWith(TodoItem t)
        {
            if (running && mode == "Focus" && current == t) { ShowTab(false); return; }
            if (segStart != null) { LogInterrupted(); segStart = null; activeSecs = 0; running = false; }
            current = t; ShowTab(false);
            mode = "Focus"; remaining = Duration(mode);
            StartTimer();
            RefreshTodos();
        }
        void StartTimer()
        {
            bool fresh = segStart == null;
            if (fresh) { segStart = DateTime.Now; activeSecs = 0; }
            endAt = DateTime.Now + remaining; lastTick = DateTime.Now; running = true; UpdateTimerUI();
            if (fresh && mode == "Focus")
            {
                if (st.Sound) Sound.Play("start", st.SoundStart);
                if (st.ToastStart)
                    ShowToastMsg("开始专注 " + st.Focus + " 分钟", current != null ? "任务：" + current.Text + TaskProgressText(current) : "未选择任务", RedC, 4000);
            }
            lbTodos.Invalidate();
        }
        string TaskProgressText(TodoItem t)
        {
            if (t == null) return "";
            return t.Target > 0 ? "\n已完成 " + t.Pomos + " / 预计 " + t.Target + " 个番茄" : (t.Pomos > 0 ? "\n已完成 " + t.Pomos + " 个番茄" : "");
        }
        void Pause() { remaining = endAt - DateTime.Now; if (remaining < TimeSpan.Zero) remaining = TimeSpan.Zero; running = false; UpdateTimerUI(); lbTodos.Invalidate(); }
        void LogInterrupted()
        {
            if (mode == "Focus" && segStart != null && activeSecs >= 60)
            {
                var r = new RecordItem { Start = segStart.Value, End = DateTime.Now, Minutes = activeSecs / 60.0, Type = "Focus", Task = current != null ? current.Text : "", TaskId = current != null ? current.Id : "", Status = "Interrupted" };
                if (Counted(r) && current != null) current.Pomos++;
                AddRecord(r);
            }
        }
        void ResetTimer()
        {
            LogInterrupted();
            running = false; segStart = null; activeSecs = 0; remaining = Duration(mode); UpdateTimerUI(); RefreshTodos();
        }
        void SkipTimer()
        {
            LogInterrupted();
            running = false; segStart = null; activeSecs = 0;
            mode = mode == "Focus" ? NextBreak(false) : "Focus";
            remaining = Duration(mode); UpdateTimerUI(); RefreshTodos();
        }
        string NextBreak(bool completed) { return completed && st.LongEvery > 0 && focusDone > 0 && focusDone % st.LongEvery == 0 ? "Long Break" : "Short Break"; }
        void AddRecord(RecordItem r) { records.Add(r); SaveAll(); RefreshRecords(); }

        void Tick(object sender, EventArgs e)
        {
            var now = DateTime.Now;
            if (running)
            {
                double dt = (now - lastTick).TotalSeconds; lastTick = now;
                if (dt > 0 && dt < 5) activeSecs += dt; else if (dt >= 5) activeSecs += Math.Max(0, Math.Min(dt, (endAt - now).TotalSeconds + dt));
                remaining = endAt - now;
                if (remaining <= TimeSpan.Zero) Complete();
                UpdateTimerUI();
            }
            if (exportPending && now >= exportAt) AutoExportNow();
            if (now >= nextSyncCheck) { nextSyncCheck = now.AddSeconds(20); SyncNow(false); }
        }
        void Complete()
        {
            running = false;
            var end = endAt;
            string finished = mode;
            var rec = new RecordItem { Start = segStart ?? end - Duration(mode), End = end, Minutes = Duration(mode).TotalMinutes, Type = mode, Task = current != null ? current.Text : "", TaskId = current != null ? current.Id : "", Status = "Completed" };
            if (mode == "Focus" && Counted(rec) && current != null) current.Pomos++;
            AddRecord(rec);
            string title, msg;
            if (finished == "Focus")
            {
                focusDone++;
                mode = NextBreak(true);
                int todayCnt = records.Count(r => r.Start.Date == DateTime.Today && Counted(r));
                title = Counted(rec) ? "番茄完成！" : "专注结束";
                msg = (current != null ? "任务：" + current.Text + TaskProgressText(current) + "\n" : "") + "今天已完成 " + todayCnt + " 个番茄。休息 " + (mode == "Long Break" ? st.LongBreak : st.ShortBreak) + " 分钟吧。";
            }
            else { mode = "Focus"; title = "休息结束"; msg = "准备开始下一个 " + st.Focus + " 分钟专注？"; }
            segStart = null; activeSecs = 0; remaining = Duration(mode);
            RefreshTodos();
            if (st.Sound) Sound.Play(finished == "Focus" ? "endFocus" : "endBreak", st.SoundEnd);
            bool autoNext = (mode == "Focus" && st.AutoStartFocus) || (mode != "Focus" && st.AutoStartBreak);
            if (st.ToastEnd)
            {
                string[] btns = autoNext ? new[] { "好的" } : mode == "Focus" ? new[] { "开始专注", "稍后" } : new[] { "开始休息", "跳过休息" };
                ShowToastMsg(title, msg, finished == "Focus" ? RedC : GreenC, st.ToastEndStays ? 0 : 10000, btns, i =>
                {
                    if (autoNext || i < 0) return;
                    if (i == 0 && !running) StartTimer();
                    else if (i == 1 && mode != "Focus") { mode = "Focus"; remaining = Duration(mode); UpdateTimerUI(); }
                });
            }
            else { try { tray.ShowBalloonTip(3000, title, msg, ToolTipIcon.Info); } catch { } }
            if (st.PopupWhenDone) { try { ShowMain(); ShowTab(false); FlashWindow(Handle, true); } catch { } }
            if (autoNext) StartTimer();
        }
        void ShowToastMsg(string title, string msg, Color c, int ms, string[] buttons = null, Action<int> onBtn = null)
        {
            try { Toast.ShowToast(title, msg, c, fUI, k, ms, buttons, onBtn); } catch { }
        }

        static string MMSS(TimeSpan t)
        {
            int secs = (int)Math.Ceiling(Math.Max(0, t.TotalSeconds));
            return (secs / 60).ToString("00") + ":" + (secs % 60).ToString("00");
        }
        void UpdateTimerUI()
        {
            var r = remaining < TimeSpan.Zero ? TimeSpan.Zero : remaining;
            var total = Duration(mode);
            var used = total - r; if (used < TimeSpan.Zero) used = TimeSpan.Zero;
            string txt = MMSS(r);
            Color c = mode == "Focus" ? RedC : GreenC;
            clockBox.Invalidate();
            lblBig.Text = txt; lblBig.ForeColor = c;
            bool idle = !running && segStart == null;
            lblMode.Text = ModeName(mode) + (running ? "中" : idle ? "" : "（已暂停）");
            lblMode.ForeColor = c;
            lblUsed.Text = "已用 " + MMSS(used) + " / " + (int)total.TotalMinutes + " 分钟  ·  剩余 " + txt;
            progress.Invalidate();
            btnStart.Text = running ? "暂停" : idle ? "开始" : "继续";
            btnStart.ForeColor = c;
            miStart.Text = running ? "暂停" : idle ? "开始番茄" : "继续";
            string title = txt + " " + ModeName(mode) + " - PomoTodo";
            if (Text != title) Text = title;

            string tip = idle ? "PomoTodo - " + ModeName(mode) + " " + (int)total.TotalMinutes + " 分钟（未开始）"
                              : ModeName(mode) + (running ? "" : "（暂停）") + "：已用 " + (int)Math.Floor(used.TotalMinutes) + " 分钟，剩余 " + txt;
            if (!idle && current != null && mode == "Focus") tip += "\n" + current.Text;
            if (tip.Length > 63) tip = tip.Substring(0, 62) + "…";
            if (tray.Text != tip) tray.Text = tip;
            UpdateTrayIcon(idle, used, r);
        }

        void UpdateTrayIcon(bool idle, TimeSpan used, TimeSpan left)
        {
            int n = st.TrayShowsUsed ? (int)Math.Floor(used.TotalMinutes) : (int)Math.Ceiling(left.TotalMinutes);
            if (n > 99) n = 99;
            string key = idle ? "idle" : mode + "|" + n + "|" + running;
            if (key == lastIconKey) return;
            lastIconKey = key;
            if (idle) { SetTrayIcon(appIcon, IntPtr.Zero); return; }
            try
            {
                int w = Math.Max(16, SystemInformation.SmallIconSize.Width);
                using (var bmp = new Bitmap(w, w))
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias; g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                    g.Clear(Color.Transparent);
                    Color c = mode == "Focus" ? RedC : GreenC;
                    if (!running) c = Color.FromArgb(120, 120, 120);
                    using (var path = new GraphicsPath())
                    {
                        float rr = w * 0.28f;
                        path.AddArc(0, 0, rr, rr, 180, 90); path.AddArc(w - 1 - rr, 0, rr, rr, 270, 90);
                        path.AddArc(w - 1 - rr, w - 1 - rr, rr, rr, 0, 90); path.AddArc(0, w - 1 - rr, rr, rr, 90, 90);
                        path.CloseFigure();
                        using (var br = new SolidBrush(c)) g.FillPath(br, path);
                    }
                    string t = n.ToString();
                    float fs = w * (t.Length > 1 ? 0.66f : 0.8f);
                    using (var f = new Font("Arial", fs, FontStyle.Bold, GraphicsUnit.Pixel))
                    using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        g.DrawString(t, f, Brushes.White, new RectangleF(-w * 0.1f, 0, w * 1.2f, w + 1), sf);
                    IntPtr h = bmp.GetHicon();
                    SetTrayIcon(Icon.FromHandle(h), h);
                }
            }
            catch { }
        }
        void SetTrayIcon(Icon ic, IntPtr h)
        {
            var old = lastHIcon;
            tray.Icon = ic;
            lastHIcon = h;
            if (old != IntPtr.Zero) { try { DestroyIcon(old); } catch { } }
        }

        // ---------------- Sync ----------------
        void SyncNow(bool manual)
        {
            if (!Store.IsSynced) { if (manual) MessageBox.Show(this, "同步未开启。请到 设置 > 同步与 Excel 开启。", "PomoTodo"); return; }
            if (!Store.SyncAvailable)
            {
                UpdateSyncLabel();
                if (manual) MessageBox.Show(this, "无法访问同步文件夹：\n" + Store.SyncDir + "\n\nOneDrive 是否在运行并已登录？\n数据仍然保存在本机，文件夹恢复后会自动写入。", "PomoTodo");
                return;
            }
            try
            {
                if (Store.SyncChangedExternally())
                {
                    string curId = current != null ? current.Id : null;
                    List<TodoItem> t; List<RecordItem> r;
                    Store.LoadData(out t, out r);
                    todos = t; records = r;
                    current = curId != null ? todos.FirstOrDefault(x => x.Id == curId) : null;
                    RefreshTodos(); RefreshRecords();
                    Store.SaveData(todos, records);
                }
                else if (manual) Store.SaveData(todos, records);
            }
            catch (Exception ex) { Store.LastSyncError = ex.Message; Store.LastSyncException = ex; }
            UpdateSyncLabel();
            if (manual)
            {
                if (Store.LastSyncException != null) ShowWriteHelp(this, Store.SyncDir, Store.LastSyncException);
                else ShowToastMsg("已同步", Store.SyncDir, GreenC, 3000);
            }
        }

        bool syncHelpShown;
        public static void ShowWriteHelp(IWin32Window owner, string path, Exception ex)
        {
            bool access = ex is UnauthorizedAccessException;
            bool cfa = Store.ControlledFolderAccessOn();
            string msg = "PomoTodo 无法写入：\n" + path + "\n\n错误：" + (ex != null ? ex.Message : "unknown") + "\n\n";
            if (access || cfa)
            {
                msg += (cfa ? "这台电脑开启了 Windows“受控制文件夹的访问”（勒索软件防护），它会阻止新程序在 文档 / 桌面 / OneDrive 里建立文件。\n\n"
                            : "最可能的原因：Windows 安全中心的“受控制文件夹的访问”（勒索软件防护）。\n\n") +
                       "解决：Windows 安全中心 > 病毒和威胁防护 > 勒索软件防护 > 通过受控制文件夹的访问允许应用 > 添加允许的应用 > 选择 PomoTodo.exe\n\n" +
                       "也可能是文件夹只读，或者可以在设置里换一个文件夹。\n\n现在打开勒索软件防护设置？";
                if (MessageBox.Show(owner, msg, "PomoTodo - 无法写入文件夹", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    try { System.Diagnostics.Process.Start("windowsdefender://ransomwareprotection"); } catch { }
            }
            else MessageBox.Show(owner, msg + "文件可能正被其他程序（例如 Excel）打开，或 OneDrive 正忙。PomoTodo 会自动重试。", "PomoTodo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        void MaybeShowSyncHelp()
        {
            if (syncHelpShown || Store.LastSyncException == null) return;
            syncHelpShown = true;
            var ex = Store.LastSyncException;
            BeginInvoke(new Action(() => ShowWriteHelp(Visible ? this : null, Store.SyncDir, ex)));
        }
        string lastLogError;
        string LogTarget()
        {
            if (!string.IsNullOrEmpty(st.LogPath)) return st.LogPath;
            return Store.SyncAvailable ? Path.Combine(Store.SyncDir, "PomoTodo_Log.xlsx") : null;
        }
        // Conflict copies of the log: "PomoTodo_Log-DESKTOP-XXX.xlsx", "PomoTodo_Log-DESKTOP-XXX-2.xlsx" (OneDrive), "PomoTodo_Log (1).xlsx" (Google Drive)
        static List<string> ExcelCopies(string target)
        {
            try
            {
                string dir = Path.GetDirectoryName(target), name = Path.GetFileNameWithoutExtension(target);
                return Directory.GetFiles(dir, name + "*.xlsx").Where(f =>
                {
                    string n = Path.GetFileNameWithoutExtension(f);
                    return !Path.GetFileName(f).StartsWith("~$") && (n.StartsWith(name + "-", StringComparison.OrdinalIgnoreCase) || n.StartsWith(name + " (", StringComparison.OrdinalIgnoreCase));
                }).ToList();
            }
            catch { return new List<string>(); }
        }
        void AutoExportNow()
        {
            exportPending = false;
            string target = LogTarget();
            if (target == null || !st.AutoExport) return;
            try
            {
                string dir = Path.GetDirectoryName(target);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                // 1) Records that exist only in the Excel file (or in its conflict copies from another PC) are merged
                //    back into the data first, so rewriting the workbook never loses history.
                var copies = ExcelCopies(target);
                int added = 0;
                foreach (var f in new[] { target }.Concat(copies))
                    if (File.Exists(f)) added += DataIO.AddRecords(records, DataIO.ReadRecords(f));
                if (added > 0) { records.Sort((a, b) => a.Start.CompareTo(b.Start)); Store.SaveData(todos, records); RefreshRecords(); RefreshTodos(); }
                // 2) Only write when the content really changed - the other PC usually already wrote the same merged data,
                //    and fewer writes = no more "PomoTodo_Log-DESKTOP-XXX.xlsx" conflict copies.
                var sheets = DataIO.BuildSheets(todos, records, Threshold);
                if (!File.Exists(target) || DataIO.Signature(target) != DataIO.Signature(sheets)) Xlsx.Write(target, sheets);
                foreach (var f in copies) Store.MoveToOld(f);
                lastLogError = null;
            }
            catch (UnauthorizedAccessException ex)
            {
                exportPending = true; exportAt = DateTime.Now.AddMinutes(1);
                if (lastLogError == null) { lastLogError = ex.Message; BeginInvoke(new Action(() => ShowWriteHelp(this, target, ex))); }
            }
            catch { exportPending = true; exportAt = DateTime.Now.AddMinutes(1); } // probably open in Excel; retry later
        }

        // ---------------- Import / Export / Settings ----------------
        void DoExport()
        {
            using (var d = new SaveFileDialog { Filter = "Excel Workbook (*.xlsx)|*.xlsx", FileName = "PomoTodo_Log_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx" })
            {
                string init = !string.IsNullOrEmpty(st.LogPath) ? Path.GetDirectoryName(st.LogPath) : Store.SyncDir;
                if (!string.IsNullOrEmpty(init) && Directory.Exists(init)) d.InitialDirectory = init;
                if (d.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    Xlsx.Write(d.FileName, DataIO.BuildSheets(todos, records, Threshold));
                    if (MessageBox.Show(this, "已导出到：\n" + d.FileName + "\n\n现在打开？", "PomoTodo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        System.Diagnostics.Process.Start(d.FileName);
                }
                catch (UnauthorizedAccessException ex) { ShowWriteHelp(this, d.FileName, ex); }
                catch (IOException ex) { MessageBox.Show(this, "无法写入文件，是否正在 Excel 中打开？\n\n" + ex.Message, "导出失败"); }
                catch (Exception ex) { MessageBox.Show(this, ex.Message, "导出失败"); }
            }
        }
        void DoImport()
        {
            using (var d = new OpenFileDialog { Filter = "Excel / CSV (*.xlsx;*.csv)|*.xlsx;*.csv|All files|*.*" })
            {
                if (d.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    string msg = DataIO.Import(d.FileName, todos, records);
                    SaveAll(); RefreshTodos(); RefreshRecords();
                    MessageBox.Show(this, msg + "\n\n提示：有 “Task/任务” 列（或 A 列）的表会导入为任务；“Target/预计” 列 = 预计番茄数；“Done Pomodoros/已完成” 列 = 已完成番茄数。\nPomoTodo 导出的文件还能恢复全部番茄记录。", "导入");
                }
                catch (IOException ex) { MessageBox.Show(this, "无法读取文件，请先在 Excel 里关闭它。\n\n" + ex.Message, "导入失败"); }
                catch (Exception ex) { MessageBox.Show(this, "只支持 .xlsx 和 .csv（不支持旧的 .xls）。\n\n" + ex.Message, "导入失败"); }
            }
        }
        void DoSettings()
        {
            ShowMain();
            UnregisterHotkeys(); // so the hotkey boxes can capture the current combos
            bool hkChanged = false;
            try
            {
                using (var f = new SettingsForm(st, fUI, k))
                {
                    if (f.ShowDialog(this) != DialogResult.OK || f.Result == null) return;
                    hkChanged = f.Result.HotkeyStart != st.HotkeyStart || f.Result.HotkeyShow != st.HotkeyShow;
                    st = f.Result; Store.SaveSettings(st);
                    ApplyStartup();
                    exportPending = true; exportAt = DateTime.Now.AddSeconds(1);
                    if (!running && segStart == null) remaining = Duration(mode);
                    lastIconKey = "";
                    syncHelpShown = false;
                    if (f.NewSyncDir != null && !string.Equals(f.NewSyncDir, Store.SyncDir, StringComparison.OrdinalIgnoreCase)) ChangeSyncDir(f.NewSyncDir);
                    else if (Store.IsSynced) SaveAll();
                    if (!string.IsNullOrEmpty(st.LogPath))
                    {
                        string ld = Path.GetDirectoryName(st.LogPath);
                        var le = Store.TestWrite(ld);
                        if (le != null) ShowWriteHelp(this, ld, le);
                    }
                    RefreshTodos(); RefreshRecords(); UpdateTimerUI(); UpdateSyncLabel();
                }
            }
            finally { RegisterHotkeys(hkChanged); }
        }
        void ChangeSyncDir(string dir)
        {
            if (dir == "") { Store.SetSyncDir(""); SaveAll(); return; }
            var terr = Store.TestWrite(dir);
            if (terr != null) { ShowWriteHelp(this, dir, terr); return; }
            try
            {
                if (Store.HasData(dir))
                {
                    var r = MessageBox.Show(this, "这个文件夹已经有 PomoTodo 数据（可能来自另一台电脑）。\n\n是 = 合并本机和文件夹的数据（推荐）\n否 = 只用文件夹里的数据（本机列表会被替换）\n取消 = 不更改",
                        "云同步", MessageBoxButtons.YesNoCancel);
                    if (r == DialogResult.Cancel) return;
                    List<TodoItem> ct; List<RecordItem> cr;
                    Store.LoadFolder(dir, out ct, out cr);
                    if (r == DialogResult.Yes) Store.Merge(todos, records, ct, cr);
                    else { todos = ct; records = cr; current = null; }
                }
                Store.SetSyncDir(dir);
                SaveAll(); RefreshTodos(); RefreshRecords();
                exportPending = true; exportAt = DateTime.Now;
                MessageBox.Show(this, "同步已开启，数据保存到：\n" + dir + "\n\nOneDrive 会自动上传。\n另一台电脑选择同一个文件夹即可共用数据。", "云同步");
            }
            catch (Exception ex) { ShowWriteHelp(this, dir, ex); }
        }
        void ApplyStartup()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (key == null) return;
                    if (st.StartWithWindows) key.SetValue("PomoTodo", "\"" + Application.ExecutablePath + "\" --tray");
                    else key.DeleteValue("PomoTodo", false);
                }
            }
            catch { }
        }

        void SaveAll()
        {
            try { Store.SaveData(todos, records); Store.SaveSettings(st); }
            catch (Exception ex) { Store.LastSyncError = ex.Message; Store.LastSyncException = ex; }
            if (Store.LastSyncException != null) MaybeShowSyncHelp();
            exportPending = true; exportAt = DateTime.Now.AddSeconds(10);
            if (lblSync != null) UpdateSyncLabel();
        }
        void OnClosing(object sender, FormClosingEventArgs e)
        {
            if (!reallyExit && st.CloseToTray && e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
                if (!trayTipShown)
                {
                    trayTipShown = true;
                    ShowToastMsg("PomoTodo 仍在运行", "计时会在右下角通知区域继续。\n右键托盘图标 > 退出 才会真正关闭。", Accent, 5000);
                }
                return;
            }
            if (running && mode == "Focus" && e.CloseReason == CloseReason.UserClosing)
            {
                ShowMain();
                var r = MessageBox.Show(this, "正在专注中，确定退出？\n（已专注的时间会被记录为“中断”）", "PomoTodo", MessageBoxButtons.YesNo);
                if (r != DialogResult.Yes) { e.Cancel = true; reallyExit = false; return; }
            }
            LogInterrupted();
            SaveAll();
            if (st.AutoExport) AutoExportNow();
            UnregisterHotkeys();
            tray.Visible = false; tray.Dispose();
            if (lastHIcon != IntPtr.Zero) { try { DestroyIcon(lastHIcon); } catch { } }
        }
    }

    static class Program
    {
        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();

        public const string Version = "1.5.0";   // keep in sync with AssemblyVersion at the top

        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length == 2 && args[0] == "--selftest") { SelfTest(args[1]); return; }
            if (args.Length == 2 && args[0] == "--synctest") { SyncTest(args[1]); return; }
            AppDomain.CurrentDomain.UnhandledException += (a, e) => ShowCrash(e.ExceptionObject as Exception);
            Application.ThreadException += (a, e) => ShowCrash(e.Exception);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            bool startHidden = args.Any(a => a.Equals("--tray", StringComparison.OrdinalIgnoreCase));
            try { if (Environment.OSVersion.Platform == PlatformID.Win32NT) SetProcessDPIAware(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            bool created;
            using (var m = new Mutex(true, "PomoTodo_SingleInstance_Mutex", out created))
            {
                if (!created)
                {
                    bool shown = false;
                    try { using (var ev = EventWaitHandle.OpenExisting("PomoTodo_ShowWindow_Event")) { ev.Set(); shown = true; } } catch { }
                    if (!shown) MessageBox.Show("PomoTodo 已经在运行（可能是旧版本，在右下角托盘里）。\n请先在托盘图标上右键 > 退出，再打开新版本。", "PomoTodo");
                    return;
                }
                using (var showEv = new EventWaitHandle(false, EventResetMode.AutoReset, "PomoTodo_ShowWindow_Event"))
                {
                    var form = new MainForm(startHidden);
                    var th = new Thread(() =>
                    {
                        while (true)
                        {
                            try { showEv.WaitOne(); } catch { return; }
                            try { if (form.IsDisposed) return; form.BeginInvoke(new Action(form.ShowMain)); } catch { }
                        }
                    }) { IsBackground = true };
                    th.Start();
                    Application.Run(form);
                }
            }
        }

        // Simulates another PC + old conflict copies in a sync folder and checks the merge
        static void SyncTest(string dir)
        {
            Directory.CreateDirectory(dir);
            string oldSync = Store.SyncDir;
            Store.SetSyncDir(dir);
            var d0 = new DateTime(2026, 9, 1, 9, 0, 0);
            // old shared files + a OneDrive conflict copy (v1.3 format)
            File.WriteAllText(Path.Combine(dir, "todos.tsv"), "a\tTask A\t0\t2026-09-01 09:00:00\t\t1\t0\nb\tTask B\t0\t2026-09-01 09:00:00\t\t0\t0\n");
            File.WriteAllText(Path.Combine(dir, "todos-DESKTOP-X-2.tsv"), "a\tTask A\t0\t2026-09-01 09:00:00\t\t3\t0\nc\tTask C\t0\t2026-09-01 09:00:00\t\t0\t0\n");
            File.WriteAllText(Path.Combine(dir, "records.tsv"), "2026-09-01 09:00:00\t2026-09-01 09:25:00\t25\tFocus\tTask A\tCompleted\ta\n");
            File.WriteAllText(Path.Combine(dir, "records-Alexhe.tsv"), "2026-09-01 10:00:00\t2026-09-01 10:25:00\t25\tFocus\tTask A\tCompleted\ta\n");
            List<TodoItem> t; List<RecordItem> r;
            Store.LoadData(out t, out r);
            Console.WriteLine("load1 todos=" + string.Join(",", t.Select(x => x.Id + ":" + x.Pomos)) + " records=" + r.Count);
            t.RemoveAll(x => x.Id == "b");
            t.Add(new TodoItem { Id = "d", Text = "Task D" });
            Store.SaveData(t, r);
            Console.WriteLine("files=" + string.Join(",", Directory.GetFiles(dir).Select(Path.GetFileName).OrderBy(x => x)));
            Console.WriteLine("old_files=" + string.Join(",", Directory.GetFiles(Path.Combine(dir, "old_files")).Select(Path.GetFileName).OrderBy(x => x)));
            // the other PC (new version) edits A, deletes C, adds E and a record
            File.WriteAllText(Path.Combine(dir, "todos@OTHER.tsv"), "a\tTask A edited\t1\t2026-09-01 09:00:00\t2026-09-02 09:00:00\t4\t0\t2099-01-01 00:00:00\nc\tTask C\t0\t2026-09-01 09:00:00\t\t0\t0\t2026-09-01 09:00:00\ne\tTask E\t0\t2026-09-01 09:00:00\t\t0\t0\t2026-09-03 09:00:00\n");
            File.WriteAllText(Path.Combine(dir, "deleted@OTHER.tsv"), "Tc\t2026-09-03 09:00:00\n");
            File.WriteAllText(Path.Combine(dir, "records@OTHER.tsv"), "2026-09-03 11:00:00\t2026-09-03 11:25:00\t25\tFocus\tTask E\tCompleted\te\n");
            Console.WriteLine("changed=" + Store.SyncChangedExternally());
            Store.LoadData(out t, out r);
            Console.WriteLine("load2 todos=" + string.Join(",", t.Select(x => x.Id + ":" + x.Text + ":" + x.Pomos + (x.Done ? ":done" : ""))) + " records=" + r.Count);
            Store.SaveData(t, r);
            Console.WriteLine("changed-after-save=" + Store.SyncChangedExternally());
            Console.WriteLine("mine:\n" + File.ReadAllText(Path.Combine(dir, "todos@" + Store.PC + ".tsv")) + File.ReadAllText(Path.Combine(dir, "deleted@" + Store.PC + ".tsv")));
            // Excel: history that only exists in the workbook / its conflict copy is merged back; deleted records stay deleted
            string xl = Path.Combine(dir, "PomoTodo_Log.xlsx");
            var old = new List<RecordItem> {
                new RecordItem { Start = new DateTime(2026, 8, 1, 9, 0, 0), End = new DateTime(2026, 8, 1, 9, 25, 0), Minutes = 25, Task = "Old only in Excel" },
                new RecordItem { Start = new DateTime(2026, 9, 1, 9, 0, 0), End = new DateTime(2026, 9, 1, 9, 25, 0), Minutes = 25, Task = "Task A" } };
            Xlsx.Write(xl, DataIO.BuildSheets(new List<TodoItem>(), old, 25));
            var other = new List<RecordItem> { new RecordItem { Start = new DateTime(2026, 8, 2, 9, 0, 0), End = new DateTime(2026, 8, 2, 9, 25, 0), Minutes = 25, Task = "Other PC" },
                new RecordItem { Start = new DateTime(2026, 9, 3, 11, 0, 0), End = new DateTime(2026, 9, 3, 11, 25, 0), Minutes = 25, Task = "Task E" } };
            Xlsx.Write(Path.Combine(dir, "PomoTodo_Log-DESKTOP-X.xlsx"), DataIO.BuildSheets(new List<TodoItem>(), other, 25));
            r.RemoveAll(x => x.Task == "Task E");   // deleted on this PC -> must not come back
            Store.SaveData(t, r);
            int added = DataIO.AddRecords(r, DataIO.ReadRecords(xl)) + DataIO.AddRecords(r, DataIO.ReadRecords(Path.Combine(dir, "PomoTodo_Log-DESKTOP-X.xlsx")));
            Console.WriteLine("excel added=" + added + " records=" + string.Join(",", r.OrderBy(x => x.Start).Select(x => x.Task)));
            var sh = DataIO.BuildSheets(t, r, 25);
            Console.WriteLine("same-before-write=" + (DataIO.Signature(xl) == DataIO.Signature(sh)));
            Xlsx.Write(xl, sh);
            Console.WriteLine("same-after-write=" + (DataIO.Signature(xl) == DataIO.Signature(sh)));
            r.Add(new RecordItem { Start = new DateTime(2026, 8, 3, 9, 0, 0), End = new DateTime(2026, 8, 3, 9, 12, 30), Minutes = 12.5, Task = "中文 task" });
            Xlsx.Write(xl, DataIO.BuildSheets(t, r, 25));
            Console.WriteLine("same-after-write2=" + (DataIO.Signature(xl) == DataIO.Signature(DataIO.BuildSheets(t, r, 25))));
            Store.SetSyncDir(oldSync);
        }

        static void ShowCrash(Exception ex)
        {
            string log = Path.Combine(Path.GetTempPath(), "PomoTodo_crash.txt");
            try { log = Path.Combine(Store.AppDir, "crash.txt"); } catch { }
            try { File.WriteAllText(log, DateTime.Now.ToString("s") + "\r\n" + ex + "\r\n", Encoding.UTF8); } catch { }
            try { MessageBox.Show("PomoTodo 出错了：\n\n" + (ex == null ? "unknown" : ex.GetType().Name + ": " + ex.Message) + "\n\n详细信息已保存到：\n" + log + "\n请把这个文件发给开发者。", "PomoTodo"); } catch { }
        }

        static void SelfTest(string path)
        {
            var todos = new List<TodoItem> { new TodoItem { Text = "#日记", Pomos = 2, Target = 10 }, new TodoItem { Text = "rj", Done = true, DoneAt = DateTime.Now } };
            var now = new DateTime(2026, 9, 21, 18, 0, 0);
            var recs = new List<RecordItem> {
                new RecordItem { Start = now, End = now.AddMinutes(25), Minutes = 25, Task = "#日记", TaskId = todos[0].Id },
                new RecordItem { Start = now.AddMinutes(30), End = now.AddMinutes(42), Minutes = 12, Task = "rj", Status = "Interrupted" } };
            Xlsx.Write(path, DataIO.BuildSheets(todos, recs, 25));
            var t2 = new List<TodoItem>(); var r2 = new List<RecordItem>();
            Console.WriteLine(DataIO.Import(path, t2, r2));
            foreach (var t in t2) Console.WriteLine("todo " + t.Pomos + "/" + t.Target + " done=" + t.Done);
            uint mm; Keys kk; Console.WriteLine("hotkey " + Hotkey.Parse("Ctrl+Alt+P", out mm, out kk) + " " + mm + " " + kk);
            File.WriteAllBytes(Path.Combine(Path.GetTempPath(), "chime.wav"), Sound.TestWav());
        }
    }
}
