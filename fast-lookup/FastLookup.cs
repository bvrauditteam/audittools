// Fast Lookup: a floating sidebar for profit centers, GL accounts, bank details and users.
// Rest the mouse on the marked zone of the right screen edge to open it, move away to send it back.
// Ctrl+Shift+F toggles it. Works fully offline; data lives in %LOCALAPPDATA%\FastFinanceLookup\data.
//
// Written in C# 5 so it compiles with the csc.exe that ships with Windows (see build.bat).

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;

[assembly: System.Reflection.AssemblyTitle("Fast Lookup")]
[assembly: System.Reflection.AssemblyProduct("Fast Lookup")]
[assembly: System.Reflection.AssemblyVersion("1.3.0.0")]

namespace FastLookup
{
    static class Program
    {
        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();

        public const string Version = "1.3";

        [STAThread]
        static int Main(string[] args)
        {
            // FastLookup.exe --import file.xlsx [more files]: import without opening the sidebar.
            if (args.Length > 0 && args[0] == "--import")
            {
                var errors = new List<string>();
                var files = new string[args.Length - 1];
                Array.Copy(args, 1, files, 0, files.Length);
                DataStore.Import(files, errors);
                foreach (var e in errors) ReportError(new Exception("Import failed: " + e));
                return errors.Count == 0 ? 0 : 1;
            }

            bool created, replaced = false;
            using (var mutex = new Mutex(true, "FastFinanceLookup.SingleInstance", out created))
            using (var showSignal = new EventWaitHandle(false, EventResetMode.AutoReset, "FastFinanceLookup.Show"))
            using (var exitSignal = new EventWaitHandle(false, EventResetMode.AutoReset, "FastFinanceLookup.Exit"))
            {
                if (!created)
                {
                    // Another copy is running (maybe an older version): ask it to quit and take over,
                    // so starting a new exe always runs the new code.
                    exitSignal.Set();
                    bool acquired;
                    try { acquired = mutex.WaitOne(5000); }
                    catch (AbandonedMutexException) { acquired = true; }
                    if (!acquired)
                    {
                        showSignal.Set();   // the old copy didn't quit (very old version): just show it
                        return 0;
                    }
                    replaced = true;
                }
                try { SetProcessDPIAware(); } catch { }
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += (s, e) => ReportError(e.Exception);
                AppDomain.CurrentDomain.UnhandledException += (s, e) => ReportError(e.ExceptionObject as Exception);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                var form = new SidebarForm(showSignal, exitSignal, replaced);
                Application.Run(form);
                try { mutex.ReleaseMutex(); } catch { }
            }
            return 0;
        }

        // Errors go to %LOCALAPPDATA%\FastFinanceLookup\error.log so problems can be diagnosed.
        public static void ReportError(Exception ex)
        {
            try
            {
                Directory.CreateDirectory(DataStore.AppDir);
                File.AppendAllText(System.IO.Path.Combine(DataStore.AppDir, "error.log"),
                    DateTime.Now.ToString("s") + "  " + ex + Environment.NewLine + Environment.NewLine);
            }
            catch { }
        }
    }

    // ===================================================================== Data

    class Record
    {
        public string Source;
        public List<KeyValuePair<string, string>> Fields = new List<KeyValuePair<string, string>>();
        public string Hay;
    }

    class Dataset
    {
        public string Name;
        public string Path;
        public List<Record> Rows = new List<Record>();
    }

    static class DataStore
    {
        public static readonly string AppDir = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FastFinanceLookup");
        public static readonly string DataDir = System.IO.Path.Combine(AppDir, "data");

        // A "data" folder next to the exe is also read, handy for a shared/USB copy.
        static string ExeDataDir
        {
            get { return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), "data"); }
        }

        public static List<Dataset> LoadAll(List<string> errors)
        {
            Directory.CreateDirectory(DataDir);
            var result = new List<Dataset>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var dir in new[] { DataDir, ExeDataDir })
            {
                if (!Directory.Exists(dir)) continue;
                var files = Directory.GetFiles(dir, "*.json");
                Array.Sort(files, StringComparer.OrdinalIgnoreCase);
                foreach (var file in files)
                {
                    var name = PrettyName(file);
                    if (!seen.Add(name)) continue;
                    try { result.Add(Load(file)); }
                    catch (Exception ex) { errors.Add(System.IO.Path.GetFileName(file) + ": " + ex.Message); }
                }
            }
            return result;
        }

        public static string PrettyName(string file)
        {
            var n = System.IO.Path.GetFileNameWithoutExtension(file).Replace('_', ' ').Trim();
            return n.Length > 0 ? n : System.IO.Path.GetFileName(file);
        }

        public static Dataset Load(string file)
        {
            var ds = new Dataset { Name = PrettyName(file), Path = file };
            var parsed = MiniJson.Parse(ReadText(file)) as List<object>;
            if (parsed == null) throw new Exception("JSON must be an array of rows");
            var hay = new StringBuilder();
            foreach (var item in parsed)
            {
                var obj = item as List<KeyValuePair<string, object>>;
                if (obj == null) continue;
                var rec = new Record { Source = ds.Name };
                hay.Length = 0;
                var keys = new HashSet<string>();
                foreach (var kv in obj)
                {
                    var key = kv.Key.Trim();
                    if (key.Length == 0 || kv.Value == null) continue;
                    var val = Convert.ToString(kv.Value, CultureInfo.InvariantCulture).Trim();
                    if (val.Length == 0 || !keys.Add(key)) continue;
                    rec.Fields.Add(new KeyValuePair<string, string>(key, val));
                    hay.Append(val.ToLowerInvariant()).Append('\u0001');
                }
                if (rec.Fields.Count == 0) continue;
                rec.Hay = hay.ToString();
                ds.Rows.Add(rec);
            }
            return ds;
        }

        // Excel exports are often Windows-1252 rather than UTF-8.
        static string ReadText(string file)
        {
            var bytes = File.ReadAllBytes(file);
            try { return new UTF8Encoding(false, true).GetString(bytes).TrimStart('\uFEFF'); }
            catch (DecoderFallbackException) { return Encoding.GetEncoding(1252).GetString(bytes); }
        }

        // Excel/CSV sheets are converted to JSON; JSON files are validated and copied.
        // Importing the same file (or sheet) again replaces the old copy.
        public static int Import(string[] files, List<string> errors)
        {
            Directory.CreateDirectory(DataDir);
            int rows = 0;
            foreach (var f in files)
            {
                try
                {
                    var ext = System.IO.Path.GetExtension(f).ToLowerInvariant();
                    if (ext == ".json")
                    {
                        var ds = Load(f);
                        File.Copy(f, System.IO.Path.Combine(DataDir, System.IO.Path.GetFileName(f)), true);
                        rows += ds.Rows.Count;
                        continue;
                    }

                    List<Table> tables;
                    if (ext == ".xlsx" || ext == ".xlsm") tables = ExcelReader.Read(f);
                    else if (ext == ".csv" || ext == ".txt") tables = new List<Table> { CsvReader.Read(f, ReadText(f)) };
                    else if (ext == ".xls") throw new Exception("old .xls format. In Excel use File > Save As > Excel Workbook (*.xlsx)");
                    else throw new Exception("unsupported file type");

                    var baseName = System.IO.Path.GetFileNameWithoutExtension(f);
                    var withData = tables.Where(t => t.Records.Count > 0).ToList();
                    if (withData.Count == 0) throw new Exception("no data found (row 1 must have column headings)");
                    foreach (var t in withData)
                    {
                        var name = withData.Count == 1 ? baseName : baseName + " - " + t.Name;
                        var target = System.IO.Path.Combine(DataDir, SafeFileName(name) + ".json");
                        File.WriteAllText(target, JsonWriter.Write(t.Records), new UTF8Encoding(false));
                        rows += t.Records.Count;
                    }
                }
                catch (Exception ex) { errors.Add(System.IO.Path.GetFileName(f) + ": " + ex.Message); }
            }
            return rows;
        }

        static string SafeFileName(string name)
        {
            foreach (var c in System.IO.Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            return name.Trim();
        }

        public static void ClearAll()
        {
            if (!Directory.Exists(DataDir)) return;
            foreach (var f in Directory.GetFiles(DataDir, "*.json")) File.Delete(f);
        }
    }

    // Tiny key=value settings file in %LOCALAPPDATA%\FastFinanceLookup\settings.ini.
    static class Settings
    {
        static readonly string FilePath = System.IO.Path.Combine(DataStore.AppDir, "settings.ini");
        static Dictionary<string, string> values;

        static void Ensure()
        {
            if (values != null) return;
            values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (!File.Exists(FilePath)) return;
                foreach (var line in File.ReadAllLines(FilePath))
                {
                    int eq = line.IndexOf('=');
                    if (eq > 0) values[line.Substring(0, eq).Trim()] = line.Substring(eq + 1).Trim();
                }
            }
            catch { }
        }

        public static string Get(string key, string fallback)
        {
            Ensure();
            string v;
            return values.TryGetValue(key, out v) ? v : fallback;
        }

        public static void Set(string key, string value)
        {
            Ensure();
            values[key] = value;
            try
            {
                Directory.CreateDirectory(DataStore.AppDir);
                File.WriteAllLines(FilePath, values.Select(kv => kv.Key + "=" + kv.Value).ToArray());
            }
            catch { }
        }
    }

    // Minimal JSON reader. Keeps key order and tolerates duplicate keys (Excel exports have many "").
    class MiniJson
    {
        readonly string s;
        int i;
        MiniJson(string text) { s = text; }

        public static object Parse(string text)
        {
            var p = new MiniJson(text);
            p.Ws();
            var v = p.Value();
            p.Ws();
            if (p.i < p.s.Length) throw p.Err("Unexpected data after JSON");
            return v;
        }

        Exception Err(string msg) { return new FormatException(msg + " at position " + i); }
        void Ws() { while (i < s.Length && char.IsWhiteSpace(s[i])) i++; }

        object Value()
        {
            if (i >= s.Length) throw Err("Unexpected end of JSON");
            char c = s[i];
            if (c == '{') return Obj();
            if (c == '[') return Arr();
            if (c == '"') return Str();
            if (c == 't' && Match("true")) return "true";
            if (c == 'f' && Match("false")) return "false";
            if (c == 'n' && Match("null")) return null;
            if (c == '-' || char.IsDigit(c))
            {
                int start = i;
                while (i < s.Length && "+-0123456789.eE".IndexOf(s[i]) >= 0) i++;
                return s.Substring(start, i - start);
            }
            throw Err("Unexpected character '" + c + "'");
        }

        bool Match(string word)
        {
            if (string.CompareOrdinal(s, i, word, 0, word.Length) != 0) return false;
            i += word.Length;
            return true;
        }

        List<KeyValuePair<string, object>> Obj()
        {
            var list = new List<KeyValuePair<string, object>>();
            i++; Ws();
            if (i < s.Length && s[i] == '}') { i++; return list; }
            while (true)
            {
                Ws();
                if (i >= s.Length || s[i] != '"') throw Err("Expected property name");
                var key = Str();
                Ws();
                if (i >= s.Length || s[i] != ':') throw Err("Expected ':'");
                i++; Ws();
                list.Add(new KeyValuePair<string, object>(key, Value()));
                Ws();
                if (i < s.Length && s[i] == ',') { i++; continue; }
                if (i < s.Length && s[i] == '}') { i++; return list; }
                throw Err("Expected ',' or '}'");
            }
        }

        List<object> Arr()
        {
            var list = new List<object>();
            i++; Ws();
            if (i < s.Length && s[i] == ']') { i++; return list; }
            while (true)
            {
                Ws();
                list.Add(Value());
                Ws();
                if (i < s.Length && s[i] == ',') { i++; continue; }
                if (i < s.Length && s[i] == ']') { i++; return list; }
                throw Err("Expected ',' or ']'");
            }
        }

        string Str()
        {
            var sb = new StringBuilder();
            i++;
            while (i < s.Length)
            {
                char c = s[i++];
                if (c == '"') return sb.ToString();
                if (c != '\\') { sb.Append(c); continue; }
                if (i >= s.Length) break;
                char e = s[i++];
                switch (e)
                {
                    case 'n': sb.Append('\n'); break;
                    case 't': sb.Append('\t'); break;
                    case 'r': sb.Append('\r'); break;
                    case 'b': sb.Append('\b'); break;
                    case 'f': sb.Append('\f'); break;
                    case 'u':
                        if (i + 4 > s.Length) throw Err("Bad \\u escape");
                        sb.Append((char)int.Parse(s.Substring(i, 4), NumberStyles.HexNumber));
                        i += 4;
                        break;
                    default: sb.Append(e); break;
                }
            }
            throw Err("Unterminated string");
        }
    }

    // ===================================================================== Excel / CSV import

    class Table
    {
        public string Name;
        public List<List<KeyValuePair<string, string>>> Records = new List<List<KeyValuePair<string, string>>>();

        // Turns a grid into records. The first non-empty row holds the column headings.
        public static Table FromGrid(string name, List<List<string>> grid)
        {
            var t = new Table { Name = name };
            List<string> headers = null;
            foreach (var row in grid)
            {
                if (row.All(string.IsNullOrWhiteSpace)) continue;
                if (headers == null)
                {
                    headers = new List<string>();
                    var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < row.Count; i++)
                    {
                        var h = (row[i] ?? "").Trim();
                        if (h.Length == 0) h = "Column " + (i + 1);
                        var unique = h;
                        for (int n = 2; !used.Add(unique); n++) unique = h + " (" + n + ")";
                        headers.Add(unique);
                    }
                    continue;
                }
                var rec = new List<KeyValuePair<string, string>>();
                for (int i = 0; i < row.Count; i++)
                {
                    var v = (row[i] ?? "").Trim();
                    if (v.Length == 0) continue;
                    var key = i < headers.Count ? headers[i] : "Column " + (i + 1);
                    rec.Add(new KeyValuePair<string, string>(key, v));
                }
                if (rec.Count > 0) t.Records.Add(rec);
            }
            return t;
        }
    }

    // Reads .xlsx/.xlsm directly (they are zip files of XML), one Table per worksheet.
    static class ExcelReader
    {
        static readonly XNamespace M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        static readonly XNamespace PR = "http://schemas.openxmlformats.org/package/2006/relationships";

        public static List<Table> Read(string path)
        {
            var tables = new List<Table>();
            // FileShare.ReadWrite lets this work while the workbook is open in Excel.
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var shared = new List<string>();
                var sst = Load(zip, "xl/sharedStrings.xml");
                if (sst != null)
                {
                    foreach (var si in sst.Root.Elements(M + "si"))
                        shared.Add(string.Concat(si.Descendants(M + "t")
                            .Where(x => x.Parent.Name != M + "rPh")   // skip phonetic hints
                            .Select(x => x.Value)));
                }

                var wb = Load(zip, "xl/workbook.xml");
                var rels = Load(zip, "xl/_rels/workbook.xml.rels");
                if (wb == null || rels == null) throw new Exception("not a valid Excel workbook");
                var targets = new Dictionary<string, string>();
                foreach (var r in rels.Root.Elements(PR + "Relationship"))
                    targets[(string)r.Attribute("Id")] = (string)r.Attribute("Target");

                foreach (var sheet in wb.Root.Element(M + "sheets").Elements(M + "sheet"))
                {
                    string target;
                    var rid = (string)sheet.Attribute(R + "id");
                    if (rid == null || !targets.TryGetValue(rid, out target)) continue;
                    var entry = target.StartsWith("/") ? target.Substring(1) : "xl/" + target;
                    var doc = Load(zip, entry);
                    if (doc == null) continue;   // chart sheets etc.
                    tables.Add(Table.FromGrid((string)sheet.Attribute("name"), ReadGrid(doc, shared)));
                }
            }
            return tables;
        }

        static XDocument Load(ZipArchive zip, string entryName)
        {
            var e = zip.GetEntry(entryName) ?? zip.Entries.FirstOrDefault(x =>
                string.Equals(x.FullName, entryName, StringComparison.OrdinalIgnoreCase));
            if (e == null) return null;
            using (var s = e.Open()) return XDocument.Load(s);
        }

        static List<List<string>> ReadGrid(XDocument doc, List<string> shared)
        {
            var grid = new List<List<string>>();
            var data = doc.Root.Element(M + "sheetData");
            if (data == null) return grid;
            foreach (var row in data.Elements(M + "row"))
            {
                var cells = new List<string>();
                foreach (var c in row.Elements(M + "c"))
                {
                    var r = (string)c.Attribute("r");
                    int col = r != null ? ColumnIndex(r) : cells.Count;
                    while (cells.Count < col) cells.Add("");
                    var text = CellText(c, shared);
                    if (col < cells.Count) cells[col] = text; else cells.Add(text);
                }
                grid.Add(cells);
            }
            return grid;
        }

        static int ColumnIndex(string cellRef)
        {
            int n = 0;
            foreach (var ch in cellRef)
            {
                if (ch < 'A' || ch > 'Z') break;
                n = n * 26 + (ch - 'A' + 1);
            }
            return Math.Max(0, n - 1);
        }

        static string CellText(XElement c, List<string> shared)
        {
            var type = (string)c.Attribute("t");
            if (type == "inlineStr")
                return string.Concat(c.Descendants(M + "t").Select(x => x.Value));
            var v = c.Element(M + "v");
            if (v == null) return "";
            var raw = v.Value;
            switch (type)
            {
                case "s":
                    int idx;
                    return int.TryParse(raw, out idx) && idx >= 0 && idx < shared.Count ? shared[idx] : "";
                case "b": return raw == "1" ? "TRUE" : "FALSE";
                case "e": return "";
                case "str": return raw;
                default:
                    // Numbers: show 1000000010, not 1.00000001E+9; trim float noise like 0.1000000001.
                    double d;
                    if ((raw.IndexOf('E') >= 0 || raw.IndexOf('e') >= 0 || raw.IndexOf('.') >= 0) &&
                        double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                    {
                        if (Math.Abs(d) < 1e15 && d == Math.Floor(d)) return ((long)d).ToString(CultureInfo.InvariantCulture);
                        return d.ToString("G15", CultureInfo.InvariantCulture);
                    }
                    return raw;
            }
        }
    }

    static class CsvReader
    {
        public static Table Read(string path, string text)
        {
            // Pick the delimiter that appears most in the first line (Excel uses ; in some regions).
            int nl = text.IndexOf('\n');
            var first = nl >= 0 ? text.Substring(0, nl) : text;
            char delim = ',';
            int best = first.Count(ch => ch == ',');
            foreach (var cand in new[] { ';', '\t' })
            {
                int n = first.Count(ch => ch == cand);
                if (n > best) { best = n; delim = cand; }
            }

            var grid = new List<List<string>>();
            var row = new List<string>();
            var field = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (quoted)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                        else quoted = false;
                    }
                    else field.Append(ch);
                }
                else if (ch == '"') quoted = true;
                else if (ch == delim) { row.Add(field.ToString()); field.Length = 0; }
                else if (ch == '\n' || ch == '\r')
                {
                    if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    row.Add(field.ToString()); field.Length = 0;
                    grid.Add(row); row = new List<string>();
                }
                else field.Append(ch);
            }
            if (field.Length > 0 || row.Count > 0) { row.Add(field.ToString()); grid.Add(row); }
            return Table.FromGrid(System.IO.Path.GetFileNameWithoutExtension(path), grid);
        }
    }

    static class JsonWriter
    {
        public static string Write(List<List<KeyValuePair<string, string>>> records)
        {
            var sb = new StringBuilder("[\n");
            for (int r = 0; r < records.Count; r++)
            {
                sb.Append("  {");
                for (int i = 0; i < records[r].Count; i++)
                {
                    if (i > 0) sb.Append(", ");
                    Str(sb, records[r][i].Key);
                    sb.Append(": ");
                    Str(sb, records[r][i].Value);
                }
                sb.Append(r < records.Count - 1 ? "},\n" : "}\n");
            }
            return sb.Append("]\n").ToString();
        }

        static void Str(StringBuilder sb, string s)
        {
            sb.Append('"');
            foreach (var c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < ' ') sb.Append("\\u").Append(((int)c).ToString("x4"));
                        else sb.Append(c);
                        break;
                }
            }
            sb.Append('"');
        }
    }

    // ===================================================================== Theme

    static class Theme
    {
        public static readonly Color HeaderA = Color.FromArgb(0x4a, 0x54, 0xe1);
        public static readonly Color HeaderB = Color.FromArgb(0x6a, 0x11, 0xcb);
        public static readonly Color Back = Color.FromArgb(0xf5, 0xf7, 0xfa);
        public static readonly Color Card = Color.White;
        public static readonly Color Border = Color.FromArgb(0xe1, 0xe8, 0xed);
        public static readonly Color Chip = Color.FromArgb(0xed, 0xf2, 0xf7);
        public static readonly Color ChipBorder = Color.FromArgb(0xcb, 0xd5, 0xe0);
        public static readonly Color Text = Color.FromArgb(0x2c, 0x3e, 0x50);
        public static readonly Color Muted = Color.FromArgb(0x71, 0x80, 0x96);
        public static readonly Color Accent = Color.FromArgb(0x4a, 0x54, 0xe1);
        public static readonly Color Copied = Color.FromArgb(0x38, 0xa1, 0x69);
        public static readonly Color Error = Color.FromArgb(0xc5, 0x30, 0x30);

        public static GraphicsPath Round(Rectangle r, int radius)
        {
            var p = new GraphicsPath();
            int d = radius * 2;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }
    }

    // ===================================================================== Results view

    // Draws result cards with clickable "chips" (one per field). Click a chip to copy its value.
    class ResultsView : Panel
    {
        class Chip { public Rectangle Rect; public string Key; public string Value; }
        class Card { public Rectangle Rect; public string Source; public List<Chip> Chips = new List<Chip>(); }

        List<Record> records = new List<Record>();
        readonly List<Card> cards = new List<Card>();
        bool showSource;
        string message = "";
        string footer = "";
        Chip hover, copied;
        readonly System.Windows.Forms.Timer copiedTimer = new System.Windows.Forms.Timer { Interval = 800 };
        readonly float scale;
        readonly Font keyFont, valueFont, srcFont, msgFont;

        public event Action<string> Copied;

        public ResultsView(float scale)
        {
            this.scale = scale;
            DoubleBuffered = true;
            AutoScroll = true;
            BackColor = Theme.Back;
            SetStyle(ControlStyles.ResizeRedraw | ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
            keyFont = new Font("Segoe UI", 7f, FontStyle.Bold);
            valueFont = new Font("Consolas", 10f);
            srcFont = new Font("Segoe UI", 7f, FontStyle.Bold);
            msgFont = new Font("Segoe UI", 9.5f);
            copiedTimer.Tick += delegate { copiedTimer.Stop(); copied = null; Invalidate(); };
        }

        int S(float v) { return (int)Math.Round(v * scale); }

        public void SetResults(List<Record> recs, bool showSrc, string footerText)
        {
            records = recs;
            showSource = showSrc;
            message = "";
            footer = footerText ?? "";
            hover = copied = null;
            AutoScrollPosition = Point.Empty;
            Relayout();
        }

        public void SetMessage(string msg)
        {
            records = new List<Record>();
            message = msg;
            footer = "";
            Relayout();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Relayout();
        }

        void Relayout()
        {
            cards.Clear();
            int pad = S(10), gap = S(6), chipPadX = S(8), chipPadY = S(4);
            int width = ClientSize.Width;
            if (width <= 0) return;
            int cardW = width - pad * 2;
            int innerW = cardW - pad * 2;
            int y = pad;
            int keyH = TextRenderer.MeasureText("A", keyFont).Height;
            int valH = TextRenderer.MeasureText("A", valueFont).Height;
            int chipH = keyH + valH + chipPadY * 2 - S(2);
            var flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine;

            foreach (var r in records)
            {
                var card = new Card { Source = r.Source };
                int cy = pad;
                if (showSource) cy += TextRenderer.MeasureText(r.Source, srcFont).Height + S(4);
                int cx = 0;
                foreach (var f in r.Fields)
                {
                    int kw = TextRenderer.MeasureText(f.Key.ToUpperInvariant(), keyFont, Size.Empty, flags).Width;
                    int vw = TextRenderer.MeasureText(f.Value, valueFont, Size.Empty, flags).Width;
                    int w = Math.Min(innerW, Math.Max(kw, vw) + chipPadX * 2);
                    if (cx > 0 && cx + w > innerW) { cx = 0; cy += chipH + gap; }
                    card.Chips.Add(new Chip { Rect = new Rectangle(pad + cx, cy, w, chipH), Key = f.Key, Value = f.Value });
                    cx += w + gap;
                }
                cy += chipH + pad;
                card.Rect = new Rectangle(pad, y, cardW, cy);
                foreach (var c in card.Chips) c.Rect.Offset(pad, y);
                cards.Add(card);
                y += cy + S(8);
            }
            if (footer.Length > 0) y += S(30);
            AutoScrollMinSize = new Size(0, y);
            Invalidate();
        }

        protected override void OnScroll(ScrollEventArgs se)
        {
            base.OnScroll(se);
            Invalidate();
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            Invalidate();
        }

        public void ScrollByWheel(int delta)
        {
            var pos = AutoScrollPosition;
            int y = Math.Max(0, -pos.Y - delta / 120 * S(60));
            AutoScrollPosition = new Point(0, y);
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(BackColor);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            if (message.Length > 0)
            {
                var rect = new Rectangle(S(20), S(30), ClientSize.Width - S(40), ClientSize.Height - S(40));
                TextRenderer.DrawText(g, message, msgFont, rect, Theme.Muted,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.WordBreak);
                return;
            }

            int oy = AutoScrollPosition.Y;
            g.TranslateTransform(0, oy);
            var visible = new Rectangle(0, -oy, ClientSize.Width, ClientSize.Height);
            var flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis;

            using (var cardBrush = new SolidBrush(Theme.Card))
            using (var borderPen = new Pen(Theme.Border))
            {
                foreach (var card in cards)
                {
                    if (!card.Rect.IntersectsWith(visible)) continue;
                    using (var path = Theme.Round(card.Rect, S(6)))
                    {
                        g.FillPath(cardBrush, path);
                        g.DrawPath(borderPen, path);
                    }
                    if (showSource)
                    {
                        TextRenderer.DrawText(g, card.Source.ToUpperInvariant(), srcFont,
                            new Point(card.Rect.X + S(10), card.Rect.Y + S(8)), Theme.Muted, TextFormatFlags.NoPadding);
                    }
                    foreach (var chip in card.Chips) DrawChip(g, chip, flags);
                }
            }

            if (footer.Length > 0)
            {
                var r = new Rectangle(0, AutoScrollMinSize.Height - S(34), ClientSize.Width, S(24));
                TextRenderer.DrawText(g, footer, msgFont, r, Theme.Muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
        }

        void DrawChip(Graphics g, Chip chip, TextFormatFlags flags)
        {
            bool isCopied = chip == copied, isHover = chip == hover;
            Color back = isCopied ? Theme.Copied : isHover ? Theme.Accent : Theme.Chip;
            Color border = isCopied ? Theme.Copied : isHover ? Theme.Accent : Theme.ChipBorder;
            Color keyColor = (isCopied || isHover) ? Color.FromArgb(220, 255, 255, 255) : Theme.Muted;
            Color valColor = (isCopied || isHover) ? Color.White : Theme.Text;

            using (var path = Theme.Round(chip.Rect, S(4)))
            using (var b = new SolidBrush(back))
            using (var p = new Pen(border))
            {
                g.FillPath(b, path);
                g.DrawPath(p, path);
            }
            int x = chip.Rect.X + S(8), w = chip.Rect.Width - S(16);
            int keyH = TextRenderer.MeasureText("A", keyFont).Height;
            TextRenderer.DrawText(g, chip.Key.ToUpperInvariant(), keyFont, new Rectangle(x, chip.Rect.Y + S(4), w, keyH), keyColor, flags);
            var valueText = isCopied ? "\u2713 Copied!" : chip.Value;
            TextRenderer.DrawText(g, valueText, valueFont,
                new Rectangle(x, chip.Rect.Y + S(2) + keyH, w, chip.Rect.Height - keyH - S(4)), valColor, flags);
        }

        Chip HitTest(Point clientPt)
        {
            var p = new Point(clientPt.X, clientPt.Y - AutoScrollPosition.Y);
            foreach (var card in cards)
            {
                if (!card.Rect.Contains(p)) continue;
                foreach (var c in card.Chips) if (c.Rect.Contains(p)) return c;
            }
            return null;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            var h = HitTest(e.Location);
            if (h != hover)
            {
                hover = h;
                Cursor = h != null ? Cursors.Hand : Cursors.Default;
                Invalidate();
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (hover != null) { hover = null; Invalidate(); }
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            if (e.Button != MouseButtons.Left) return;
            var c = HitTest(e.Location);
            if (c == null) return;
            try { Clipboard.SetText(c.Value); }
            catch { return; }
            copied = c;
            copiedTimer.Stop();
            copiedTimer.Start();
            Invalidate();
            if (Copied != null) Copied(c.Value);
        }

        // Copies the first chip of the top result (Enter in the search box).
        public string CopyFirst()
        {
            if (cards.Count == 0 || cards[0].Chips.Count == 0) return null;
            var c = cards[0].Chips[0];
            try { Clipboard.SetText(c.Value); } catch { return null; }
            copied = c;
            copiedTimer.Stop();
            copiedTimer.Start();
            Invalidate();
            return c.Value;
        }
    }

    // ===================================================================== Sidebar window

    class SidebarForm : Form, IMessageFilter
    {
        // ---- Win32
        [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hWnd, int id, uint mods, uint vk);
        [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr pid);
        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool attach);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, string l);
        [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point p);

        const int WM_HOTKEY = 0x0312, WM_MOUSEWHEEL = 0x020A, EM_SETCUEBANNER = 0x1501;
        const int HOTKEY_ID = 0xF00D;
        const uint MOD_CONTROL = 0x2, MOD_SHIFT = 0x4, MOD_NOREPEAT = 0x4000;
        const int EDGE_DWELL_MS = 250, LEAVE_DELAY_MS = 600, MAX_RESULTS = 150, EMPTY_QUERY_RESULTS = 50;
        const string RUN_KEY = @"Software\Microsoft\Windows\CurrentVersion\Run", RUN_NAME = "FastFinanceLookup";

        enum State { Closed, Opening, Open, Closing }

        readonly float scale;
        readonly EventWaitHandle showSignal, exitSignal;
        readonly bool openOnStart;
        readonly TextBox search;
        readonly Label countLabel, statusLabel;
        readonly ResultsView results;
        readonly HeaderButton pinButton;
        readonly NotifyIcon tray;
        readonly ContextMenuStrip menu;
        readonly ToolStripMenuItem startupItem;
        readonly System.Windows.Forms.Timer edgeTimer, animTimer, searchTimer, statusTimer;

        List<Dataset> datasets = new List<Dataset>();
        List<Record> all = new List<Record>();
        State state = State.Closed;
        bool pinned, busy, holdUntilMouseEnters, exiting, started;
        int edgeSince, leaveSince;
        int animFrom, animTo, animStart;
        IntPtr previousWindow = IntPtr.Zero;
        Screen screen;
        readonly EdgeMarker marker;
        readonly ToolStripMenuItem markerItem;
        readonly Dictionary<string, ToolStripMenuItem> zoneItems = new Dictionary<string, ToolStripMenuItem>();
        string zonePosition = Settings.Get("zone", "middle");
        bool showMarker = Settings.Get("marker", "on") != "off";
        Rectangle hotZone;

        public SidebarForm(EventWaitHandle showSignal, EventWaitHandle exitSignal, bool openOnStart)
        {
            this.showSignal = showSignal;
            this.exitSignal = exitSignal;
            this.openOnStart = openOnStart;
            // Don't use CreateGraphics() here: it would create the window handle too early.
            using (var g = Graphics.FromHwnd(IntPtr.Zero)) scale = g.DpiX / 96f;

            Text = "Fast Lookup";
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Theme.Border;
            Padding = new Padding(S(1), 0, 0, 0); // thin left border
            KeyPreview = true;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            // ---- Header
            var header = new GradientPanel { Dock = DockStyle.Top, Height = S(46) };
            var title = new Label
            {
                Text = "Fast Lookup",
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI Semibold", 11f),
                AutoSize = true,
                Location = new Point(S(12), S(12))
            };
            var closeButton = new HeaderButton("\uE76C", "Send back (Esc)", S(32));   // ChevronRight
            var menuButton = new HeaderButton("\uE712", "Menu", S(32));              // More
            pinButton = new HeaderButton("\uE718", "Pin open", S(32));               // Pin
            header.Controls.Add(title);
            foreach (var b in new[] { closeButton, menuButton, pinButton })
            {
                header.Controls.Add(b);
                b.Top = S(7);
            }
            header.Resize += delegate
            {
                closeButton.Left = header.Width - S(40);
                menuButton.Left = header.Width - S(74);
                pinButton.Left = header.Width - S(108);
            };
            closeButton.Click += delegate { CloseSidebar(); };
            pinButton.Click += delegate { SetPinned(!pinned); };

            // ---- Search box
            var searchWrap = new Panel { Dock = DockStyle.Top, Height = S(62), BackColor = Color.White, Padding = new Padding(S(12), S(10), S(12), 0) };
            var searchFrame = new RoundedFrame { Dock = DockStyle.Top, Height = S(36), Padding = new Padding(S(10), S(8), S(10), S(4)) };
            search = new TextBox
            {
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 11f),
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(0xf8, 0xf9, 0xfa),
                ForeColor = Theme.Text
            };
            searchFrame.Inner = search;
            searchFrame.Controls.Add(search);
            countLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = S(16),
                ForeColor = Theme.Muted,
                Font = new Font("Segoe UI", 8f),
                Padding = new Padding(S(2), 0, 0, 0)
            };
            searchWrap.Controls.Add(countLabel);
            searchWrap.Controls.Add(searchFrame);
            var searchLine = new Panel { Dock = DockStyle.Top, Height = S(1), BackColor = Theme.Border };

            // ---- Results + status bar
            results = new ResultsView(scale) { Dock = DockStyle.Fill };
            statusLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = S(26),
                BackColor = Color.White,
                ForeColor = Theme.Muted,
                Font = new Font("Segoe UI", 8.5f),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(S(10), 0, 0, 0),
                Text = "Click a value to copy it"
            };

            Controls.Add(results);
            Controls.Add(statusLabel);
            Controls.Add(searchLine);
            Controls.Add(searchWrap);
            Controls.Add(header);

            // ---- Menu (shared by the header menu button and the tray icon)
            menu = new ContextMenuStrip();
            menu.Items.Add(new ToolStripMenuItem("Fast Lookup v" + Program.Version) { Enabled = false });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Open sidebar", null, delegate { OpenSidebar(true, true); });
            menu.Items.Add("Import Excel / CSV / JSON\u2026", null, delegate { ImportFiles(); });
            menu.Items.Add("Reload data", null, delegate { ReloadData(true); });
            menu.Items.Add("Open data folder", null, delegate
            {
                Directory.CreateDirectory(DataStore.DataDir);
                Process.Start("explorer.exe", "\"" + DataStore.DataDir + "\"");
            });
            menu.Items.Add("Clear database\u2026", null, delegate { ClearDatabase(); });
            menu.Items.Add(new ToolStripSeparator());
            var zoneMenu = new ToolStripMenuItem("Hover zone on right edge");
            foreach (var z in new[] { "top", "middle", "bottom" })
            {
                var pos = z;
                var item = new ToolStripMenuItem(char.ToUpper(z[0]) + z.Substring(1), null, delegate { SetZone(pos); });
                zoneItems[z] = item;
                zoneMenu.DropDownItems.Add(item);
            }
            menu.Items.Add(zoneMenu);
            markerItem = new ToolStripMenuItem("Show edge marker", null, delegate
            {
                showMarker = !showMarker;
                Settings.Set("marker", showMarker ? "on" : "off");
                UpdateMarker();
            });
            menu.Items.Add(markerItem);
            startupItem = new ToolStripMenuItem("Start with Windows", null, delegate { ToggleStartup(); });
            menu.Items.Add(startupItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Exit", null, delegate { exiting = true; Close(); });
            menu.Opening += delegate
            {
                busy = true;
                startupItem.Checked = IsStartupEnabled();
                markerItem.Checked = showMarker;
                foreach (var kv in zoneItems) kv.Value.Checked = kv.Key == zonePosition;
            };
            marker = new EdgeMarker();
            menu.Closed += delegate { busy = false; };
            menuButton.Click += delegate { menu.Show(menuButton, new Point(0, menuButton.Height)); };

            tray = new NotifyIcon
            {
                Icon = Icon ?? SystemIcons.Application,
                Text = "Fast Lookup (right-edge marker or Ctrl+Shift+F)",
                ContextMenuStrip = menu,
                Visible = true
            };
            tray.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) ToggleSidebar(); };

            // ---- Events
            search.TextChanged += delegate { searchTimer.Stop(); searchTimer.Start(); };
            search.KeyDown += OnSearchKeyDown;
            results.Copied += v => ShowStatus("Copied: " + v, false);
            results.MouseDown += delegate { holdUntilMouseEnters = false; };

            searchTimer = new System.Windows.Forms.Timer { Interval = 60 };
            searchTimer.Tick += delegate { searchTimer.Stop(); RunSearch(); };
            statusTimer = new System.Windows.Forms.Timer { Interval = 2500 };
            statusTimer.Tick += delegate { statusTimer.Stop(); statusLabel.ForeColor = Theme.Muted; statusLabel.Text = "Click a value to copy it"; };
            animTimer = new System.Windows.Forms.Timer { Interval = 10 };
            animTimer.Tick += delegate { AnimateStep(); };
            edgeTimer = new System.Windows.Forms.Timer { Interval = 60 };
            edgeTimer.Tick += delegate { WatchMouse(); };

            // Drop Excel/CSV/JSON files onto the sidebar to import them.
            AllowDrop = true;
            DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            DragDrop += (s, e) =>
            {
                var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (paths != null && paths.Length > 0) ImportPaths(paths);
            };

            Application.AddMessageFilter(this);
            ReloadData(false);
        }

        int S(float v) { return (int)Math.Round(v * scale); }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x80;        // WS_EX_TOOLWINDOW: keep out of Alt+Tab
                cp.ClassStyle |= 0x20000;  // CS_DROPSHADOW
                return cp;
            }
        }

        protected override void SetVisibleCore(bool value)
        {
            // Start hidden; the sidebar only appears when summoned.
            if (!started)
            {
                started = true;
                if (!IsHandleCreated) CreateHandle();
                OnCreatedHidden();
                value = false;
            }
            base.SetVisibleCore(value);
        }

        void OnCreatedHidden()
        {
            SendMessage(search.Handle, EM_SETCUEBANNER, (IntPtr)1, "Search GL, profit center, plant, bank\u2026");
            if (!RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT | MOD_NOREPEAT, (uint)Keys.F))
                tray.ShowBalloonTip(4000, "Fast Lookup", "Ctrl+Shift+F is used by another program. Use the screen edge or the tray icon.", ToolTipIcon.Warning);
            edgeTimer.Start();
            UpdateMarker();
            new Thread(() =>
            {
                var handles = new WaitHandle[] { showSignal, exitSignal };
                while (!exiting)
                {
                    int which = WaitHandle.WaitAny(handles, 1000);
                    if (which == WaitHandle.WaitTimeout) continue;
                    try
                    {
                        if (which == 0) BeginInvoke((Action)(() => OpenSidebar(true, true)));
                        else { BeginInvoke((Action)(() => { exiting = true; Close(); })); return; }
                    }
                    catch { return; }
                }
            }) { IsBackground = true }.Start();
            if (openOnStart) BeginInvoke((Action)(() => OpenSidebar(true, true)));
            tray.ShowBalloonTip(3000, "Fast Lookup " + Program.Version + " is running",
                "Rest the mouse on the purple strip at the right edge of the screen, or press Ctrl+Shift+F.", ToolTipIcon.Info);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && (int)m.WParam == HOTKEY_ID)
            {
                ToggleSidebar();
                return;
            }
            base.WndProc(ref m);
        }

        // Send the mouse wheel to the results list whenever the cursor is over it.
        public bool PreFilterMessage(ref Message m)
        {
            if (m.Msg != WM_MOUSEWHEEL || !Visible) return false;
            var pt = Cursor.Position;
            if (!results.RectangleToScreen(results.ClientRectangle).Contains(pt)) return false;
            int delta = (short)((long)m.WParam >> 16);
            results.ScrollByWheel(delta);
            return true;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!exiting && e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                CloseSidebar();
                return;
            }
            exiting = true;
            edgeTimer.Stop();
            marker.Close();
            UnregisterHotKey(Handle, HOTKEY_ID);
            tray.Visible = false;
            tray.Dispose();
            base.OnFormClosing(e);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape) { CloseSidebar(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ------------------------------------------------------------ Open / close

        void ToggleSidebar()
        {
            if (state == State.Open || state == State.Opening) CloseSidebar();
            else OpenSidebar(true, true);
        }

        // activate: take keyboard focus so you can type right away.
        // hold: don't auto-hide until the mouse has visited the sidebar once (used for hotkey/tray opens).
        void OpenSidebar(bool activate, bool hold)
        {
            holdUntilMouseEnters = hold;
            leaveSince = 0;
            if (state == State.Open || state == State.Opening)
            {
                if (activate) FocusSearch();
                return;
            }

            var fg = GetForegroundWindow();
            if (fg != Handle) previousWindow = fg;

            screen = hold ? Screen.FromPoint(Cursor.Position) : EdgeScreen();
            var wa = screen.WorkingArea;
            int width = Math.Min(S(420), wa.Width);
            Bounds = new Rectangle(wa.Right, wa.Top, width, wa.Height);
            if (!Visible) Show();
            if (activate) FocusSearch();
            StartAnim(wa.Right - width);
            state = State.Opening;
            UpdateMarker();
        }

        void CloseSidebar()
        {
            if (state == State.Closed || state == State.Closing) return;
            bool wasForeground = GetForegroundWindow() == Handle;
            var wa = (screen ?? Screen.FromControl(this)).WorkingArea;
            StartAnim(wa.Right);
            state = State.Closing;
            if (wasForeground && previousWindow != IntPtr.Zero) SetForegroundWindow(previousWindow);
        }

        void StartAnim(int targetLeft)
        {
            animFrom = Left;
            animTo = targetLeft;
            animStart = Environment.TickCount;
            animTimer.Start();
        }

        void AnimateStep()
        {
            const int duration = 160;
            double t = Math.Min(1.0, (Environment.TickCount - animStart) / (double)duration);
            double eased = 1 - Math.Pow(1 - t, 3);
            Left = animFrom + (int)Math.Round((animTo - animFrom) * eased);
            if (t < 1) return;
            animTimer.Stop();
            if (state == State.Opening) state = State.Open;
            else if (state == State.Closing) { state = State.Closed; Hide(); }
            UpdateMarker();
        }

        void FocusSearch()
        {
            // Windows only lets the foreground app steal focus; borrow its input queue briefly.
            var fg = GetForegroundWindow();
            uint fgThread = GetWindowThreadProcessId(fg, IntPtr.Zero), me = GetCurrentThreadId();
            if (fg != Handle && fgThread != me)
            {
                AttachThreadInput(me, fgThread, true);
                SetForegroundWindow(Handle);
                AttachThreadInput(me, fgThread, false);
            }
            else SetForegroundWindow(Handle);
            Activate();
            search.Focus();
            search.SelectAll();
        }

        void SetPinned(bool value)
        {
            pinned = value;
            pinButton.Glyph = pinned ? "\uE840" : "\uE718"; // Pinned / Pin
            pinButton.Active = pinned;
            pinButton.Tip = pinned ? "Unpin (auto-hide)" : "Pin open";
            ShowStatus(pinned ? "Pinned: stays open until you unpin or press Esc" : "Auto-hide on", false);
        }

        // The monitor whose right edge is the far right of the desktop.
        static Screen EdgeScreen()
        {
            Screen best = Screen.PrimaryScreen;
            foreach (var sc in Screen.AllScreens) if (sc.Bounds.Right > best.Bounds.Right) best = sc;
            return best;
        }

        // Only this band of the right edge opens the sidebar (top, middle or bottom).
        Rectangle ComputeHotZone()
        {
            var wa = EdgeScreen().WorkingArea;
            int h = Math.Min(S(180), wa.Height);
            int top;
            if (zonePosition == "top") top = wa.Top + S(60);          // below window close buttons
            else if (zonePosition == "bottom") top = wa.Bottom - h - S(20);
            else top = wa.Top + (wa.Height - h) / 2;
            return new Rectangle(wa.Right - 1, top, 1, h);
        }

        void SetZone(string pos)
        {
            zonePosition = pos;
            Settings.Set("zone", pos);
            UpdateMarker();
            ShowStatus("Hover zone: " + pos + " of the right edge", false);
            if (state == State.Closed) marker.Flash();
        }

        void UpdateMarker()
        {
            hotZone = ComputeHotZone();
            bool visible = showMarker && (state == State.Closed || state == State.Closing);
            marker.Place(new Rectangle(hotZone.Right - S(4), hotZone.Top, S(4), hotZone.Height), visible);
        }

        // Polls the cursor: dwell on the right edge opens, leaving the sidebar closes.
        void WatchMouse()
        {
            if (busy) { leaveSince = 0; return; }
            var p = Cursor.Position;
            int now = Environment.TickCount;

            if (state == State.Closed || state == State.Closing)
            {
                var zone = ComputeHotZone();
                if (zone != hotZone || marker.Visible != (showMarker && state == State.Closed)) UpdateMarker();  // screens changed
                // Works while dragging a file too, so you can drop it straight onto the sidebar.
                bool atEdge = p.X >= zone.Left && p.Y >= zone.Top && p.Y < zone.Bottom;
                marker.Highlight(atEdge);
                if (!atEdge) { edgeSince = 0; return; }
                if (edgeSince == 0) { edgeSince = now; return; }
                if (now - edgeSince >= EDGE_DWELL_MS)
                {
                    edgeSince = 0;
                    OpenSidebar(true, false);
                }
                return;
            }

            if (pinned) { leaveSince = 0; return; }
            var area = Bounds;
            area.Inflate(S(20), S(4));
            bool inside = area.Contains(p);
            if (inside) { holdUntilMouseEnters = false; leaveSince = 0; return; }
            if (holdUntilMouseEnters || Control.MouseButtons != MouseButtons.None) { leaveSince = 0; return; }
            if (leaveSince == 0) { leaveSince = now; return; }
            if (now - leaveSince >= LEAVE_DELAY_MS)
            {
                leaveSince = 0;
                CloseSidebar();
            }
        }

        // ------------------------------------------------------------ Search

        void OnSearchKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                searchTimer.Stop();
                RunSearch();
                var v = results.CopyFirst();
                if (v != null) ShowStatus("Copied: " + v, false);
            }
            else if (e.KeyCode == Keys.Down || e.KeyCode == Keys.PageDown)
            {
                e.SuppressKeyPress = true;
                results.ScrollByWheel(e.KeyCode == Keys.Down ? -120 : -600);
            }
            else if (e.KeyCode == Keys.Up || e.KeyCode == Keys.PageUp)
            {
                e.SuppressKeyPress = true;
                results.ScrollByWheel(e.KeyCode == Keys.Up ? 120 : 600);
            }
        }

        void RunSearch()
        {
            if (all.Count == 0)
            {
                countLabel.Text = "";
                results.SetMessage("No database loaded.\n\nClick \u22EF \u2192 Import Excel / CSV / JSON, or drag your Excel file onto this panel.\nRow 1 of each sheet must hold the column headings. You only need to do this once.");
                return;
            }

            var query = search.Text.Trim().ToLowerInvariant();
            bool multi = datasets.Count > 1;
            if (query.Length == 0)
            {
                var first = all.GetRange(0, Math.Min(EMPTY_QUERY_RESULTS, all.Count));
                countLabel.Text = string.Format("Showing first {0} of {1:N0} records", first.Count, all.Count);
                results.SetResults(first, multi, null);
                return;
            }

            // Every word you type must appear somewhere in the row.
            var terms = query.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            var scored = new List<KeyValuePair<int, Record>>();
            int order = 0;
            foreach (var r in all)
            {
                bool ok = true;
                foreach (var t in terms) if (r.Hay.IndexOf(t, StringComparison.Ordinal) < 0) { ok = false; break; }
                if (!ok) continue;
                // Exact and prefix hits on a field rank first; ties keep file order.
                int score = 0;
                foreach (var f in r.Fields)
                {
                    var lv = f.Value.ToLowerInvariant();
                    foreach (var t in terms)
                    {
                        if (lv == t) score += 100;
                        else if (lv.StartsWith(t, StringComparison.Ordinal)) score += 10;
                    }
                }
                scored.Add(new KeyValuePair<int, Record>(-score * 1000000 + order++, r));
            }
            scored.Sort((a, b) => a.Key.CompareTo(b.Key));

            int total = scored.Count;
            var top = new List<Record>();
            for (int i = 0; i < Math.Min(MAX_RESULTS, total); i++) top.Add(scored[i].Value);

            if (total == 0)
            {
                countLabel.Text = "";
                results.SetMessage("No matches found");
                return;
            }
            countLabel.Text = string.Format("{0:N0} result{1}", total, total == 1 ? "" : "s");
            results.SetResults(top, multi, total > MAX_RESULTS
                ? string.Format("Showing {0} of {1:N0}. Keep typing to narrow down.", MAX_RESULTS, total)
                : null);
        }

        // ------------------------------------------------------------ Data actions

        void ReloadData(bool announce)
        {
            var errors = new List<string>();
            datasets = DataStore.LoadAll(errors);
            all = new List<Record>();
            foreach (var d in datasets) all.AddRange(d.Rows);
            RunSearch();
            if (errors.Count > 0) ShowStatus(string.Join(" \u00B7 ", errors.ToArray()), true);
            else if (announce) ShowStatus(string.Format("Loaded {0:N0} records from {1} file(s)", all.Count, datasets.Count), false);
        }

        void ImportFiles()
        {
            busy = true;
            try
            {
                using (var dlg = new OpenFileDialog
                {
                    Title = "Import lookup data",
                    Filter = "Lookup data (*.xlsx;*.xlsm;*.csv;*.json)|*.xlsx;*.xlsm;*.csv;*.json|Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|CSV files (*.csv)|*.csv|JSON files (*.json)|*.json|All files (*.*)|*.*",
                    Multiselect = true
                })
                {
                    if (dlg.ShowDialog(this) != DialogResult.OK) return;
                    ImportPaths(dlg.FileNames);
                }
            }
            finally { busy = false; }
        }

        void ImportPaths(string[] paths)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                var errors = new List<string>();
                int rows = DataStore.Import(paths, errors);
                ReloadData(false);
                if (errors.Count > 0)
                {
                    ShowStatus("Import problem, see message", true);
                    foreach (var e in errors) Program.ReportError(new Exception("Import failed: " + e));
                    bool wasBusy = busy;
                    busy = true;
                    MessageBox.Show(this, "Some files could not be imported:\n\n" + string.Join("\n\n", errors.ToArray()) +
                        (rows > 0 ? string.Format("\n\nOther files added {0:N0} records.", rows) : ""),
                        "Fast Lookup", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    busy = wasBusy;
                }
                else ShowStatus(string.Format("\u2713 Added {0:N0} records. Total is now {1:N0}.", rows, all.Count), false);
            }
            finally { Cursor = Cursors.Default; }
        }

        void ClearDatabase()
        {
            busy = true;
            try
            {
                if (MessageBox.Show(this, "Are you sure you want to completely clear the lookup database?",
                        "Fast Lookup", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
                DataStore.ClearAll();
                ReloadData(false);
                ShowStatus("Database cleared completely.", true);
            }
            finally { busy = false; }
        }

        static bool IsStartupEnabled()
        {
            using (var k = Registry.CurrentUser.OpenSubKey(RUN_KEY))
                return k != null && k.GetValue(RUN_NAME) != null;
        }

        void ToggleStartup()
        {
            using (var k = Registry.CurrentUser.CreateSubKey(RUN_KEY))
            {
                if (IsStartupEnabled()) { k.DeleteValue(RUN_NAME, false); ShowStatus("Won't start with Windows", false); }
                else { k.SetValue(RUN_NAME, "\"" + Application.ExecutablePath + "\""); ShowStatus("Will start with Windows", false); }
            }
        }

        void ShowStatus(string msg, bool error)
        {
            statusLabel.Text = msg;
            statusLabel.ForeColor = error ? Theme.Error : Theme.Copied;
            statusTimer.Stop();
            statusTimer.Start();
        }
    }

    // ===================================================================== Small controls

    // Thin strip showing where the hover zone is. Clicks pass straight through it.
    class EdgeMarker : Form
    {
        const double IdleOpacity = 0.45, HotOpacity = 1.0;
        readonly System.Windows.Forms.Timer flashTimer = new System.Windows.Forms.Timer { Interval = 1200 };

        public EdgeMarker()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Theme.Accent;
            Opacity = IdleOpacity;
            flashTimer.Tick += delegate { flashTimer.Stop(); Opacity = IdleOpacity; };
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                // TOOLWINDOW | LAYERED | TRANSPARENT (click-through) | NOACTIVATE
                cp.ExStyle |= 0x80 | 0x80000 | 0x20 | 0x08000000;
                return cp;
            }
        }

        public void Place(Rectangle r, bool visible)
        {
            if (Bounds != r) Bounds = r;
            if (visible && !Visible) Show();
            else if (!visible && Visible) Hide();
        }

        public void Highlight(bool hot)
        {
            if (flashTimer.Enabled) return;
            double o = hot ? HotOpacity : IdleOpacity;
            if (Math.Abs(Opacity - o) > 0.01) Opacity = o;
        }

        public void Flash()
        {
            Opacity = HotOpacity;
            flashTimer.Stop();
            flashTimer.Start();
        }
    }

    class GradientPanel : Panel
    {
        public GradientPanel() { DoubleBuffered = true; }
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            if (Width <= 0 || Height <= 0) return;
            using (var b = new LinearGradientBrush(ClientRectangle, Theme.HeaderA, Theme.HeaderB, 20f))
                e.Graphics.FillRectangle(b, ClientRectangle);
        }
    }

    class RoundedFrame : Panel
    {
        public Control Inner;
        public RoundedFrame()
        {
            DoubleBuffered = true;
            BackColor = Color.White;
            SetStyle(ControlStyles.ResizeRedraw, true);
        }
        protected override void OnControlAdded(ControlEventArgs e)
        {
            base.OnControlAdded(e);
            e.Control.GotFocus += delegate { Invalidate(); };
            e.Control.LostFocus += delegate { Invalidate(); };
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            bool focused = Inner != null && Inner.Focused;
            using (var path = Theme.Round(r, 6))
            using (var b = new SolidBrush(Inner != null ? Inner.BackColor : BackColor))
            using (var p = new Pen(focused ? Theme.Accent : Theme.Border, 2f))
            {
                e.Graphics.FillPath(b, path);
                e.Graphics.DrawPath(p, path);
            }
        }
    }

    class HeaderButton : Control
    {
        string glyph;
        bool hover, active;
        readonly ToolTip tip = new ToolTip();
        readonly Font font;

        public HeaderButton(string glyph, string tooltip, int size)
        {
            this.glyph = glyph;
            Size = new Size(size, size);
            Cursor = Cursors.Hand;
            SetStyle(ControlStyles.SupportsTransparentBackColor | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
            BackColor = Color.Transparent;
            font = new Font(IconFontName(), 11f);
            tip.SetToolTip(this, tooltip);
        }

        static string IconFontName()
        {
            foreach (var f in FontFamily.Families)
                if (f.Name == "Segoe Fluent Icons") return f.Name;
            return "Segoe MDL2 Assets";
        }

        public string Glyph { set { glyph = value; Invalidate(); } }
        public bool Active { set { active = value; Invalidate(); } }
        public string Tip { set { tip.SetToolTip(this, value); } }

        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            if (hover || active)
            {
                using (var path = Theme.Round(new Rectangle(0, 0, Width - 1, Height - 1), 5))
                using (var b = new SolidBrush(Color.FromArgb(active ? 90 : 50, 255, 255, 255)))
                    e.Graphics.FillPath(b, path);
            }
            TextRenderer.DrawText(e.Graphics, glyph, font, ClientRectangle, Color.White,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }
    }
}
