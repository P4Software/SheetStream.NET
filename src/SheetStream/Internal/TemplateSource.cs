using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace SheetStream.Internal
{
    /// <summary>
    /// A template workbook read into memory, with the structure SheetStream needs to fill it: sheet names and
    /// part paths, the shared string table, the style counts, and the parts that must be rewritten. Templates
    /// are small (a styled header and a data area), so holding the bytes is cheap and lets several readers
    /// open parts independently.
    /// </summary>
    /// <remarks>
    /// The packed file is held as it arrived, so its size is the caller's to bound (an upload size limit).
    /// What this class bounds is what unpacking can add: every part is opened through <see cref="OpenXml"/>,
    /// which applies the same limits as <see cref="Reading.ExcelReader"/>, because a workbook handed to
    /// <c>OpenTemplate</c> or <c>Edit</c> may be one a user uploaded.
    /// </remarks>
    internal sealed class TemplateSource
    {
        private readonly byte[] _bytes;

        private TemplateSource(byte[] bytes, Reading.ReaderOptions limits)
        {
            _bytes = bytes;
            Limits = limits;
            Budget = new MemoryBudget(limits.MaxLoadedPartBytes);
        }

        /// <summary>
        /// What the template makes the package keep in memory: its shared strings, styles, sheet list and any
        /// sheets loaded for editing, all charged against <see cref="Reading.ReaderOptions.MaxLoadedPartBytes"/>.
        /// </summary>
        public MemoryBudget Budget { get; }

        /// <summary>Limits applied to everything read out of the template.</summary>
        public Reading.ReaderOptions Limits { get; }

        /// <summary>
        /// Called when reading any part of the template fails. The package uses it to give up on saving: a
        /// workbook built from a template that could not be read is not one to write out, and trying to would
        /// replace the error that explains the problem with a confusing one from the half-written output.
        /// </summary>
        public Action? ReadFailed { get; set; }

        /// <summary>Path of the workbook part, normally <c>xl/workbook.xml</c>.</summary>
        public string WorkbookPart { get; private set; } = "xl/workbook.xml";

        /// <summary>Folder of the workbook part with trailing slash, e.g. <c>xl/</c>.</summary>
        public string WorkbookFolder { get; private set; } = "xl/";

        /// <summary>Sheets in tab order: name, part path, zero-based position.</summary>
        public List<(string Name, string Part)> Sheets { get; } = new List<(string, string)>();

        /// <summary>Shared strings part, or null when the template has none.</summary>
        public string? SharedStringsPart { get; private set; }

        /// <summary>Styles part. Every Excel-saved workbook has one.</summary>
        public string? StylesPart { get; private set; }

        /// <summary>Calculation chain part, dropped on save because filled cells make it stale.</summary>
        public string? CalcChainPart { get; private set; }

        /// <summary>Number of <c>si</c> items in the template's shared string table.</summary>
        public int SharedStringCount { get; private set; }

        /// <summary>Plain-text shared strings and their indexes, so new cells can reuse them.</summary>
        public Dictionary<string, int> PlainSharedStrings { get; } = new Dictionary<string, int>(StringComparer.Ordinal);

        /// <summary>True when the workbook counts dates from 1904 (the old Mac convention) instead of 1900.</summary>
        public bool Date1904 { get; private set; }

        // Text of every shared string by index, loaded on first use because only edited sheets need it.
        private List<string>? _sharedStringTexts;

        /// <summary>
        /// The text of every shared string item (rich text flattened), by index. Needed to show the value of a
        /// <c>t="s"</c> cell when a sheet is edited; streaming fills never read cell values, so this is lazy.
        /// </summary>
        public IReadOnlyList<string> GetSharedStringTexts()
        {
            if (_sharedStringTexts != null) return _sharedStringTexts;
            var texts = new List<string>(SharedStringCount);
            if (SharedStringsPart != null)
            {
                // The table as a whole is bounded by MaxLoadedPartBytes; each item by the cell text limit.
                var state = new Reading.ExcelReader.TextState(Limits.MaxCellTextLength, long.MaxValue);
                using (var archive = OpenArchive())
                using (var r = OpenXml(archive, SharedStringsPart, held: true))
                {
                    while (r != null && r.Read())
                    {
                        // Depth 1: the table's own items. An "si" nested in an extension element is not one.
                        if (r.NodeType != XmlNodeType.Element || r.LocalName != "si" || r.Depth != 1) continue;
                        state.BeginCell();
                        var item = r.IsEmptyElement ? "" : Reading.ExcelReader.ReadRichText(r, state);
                        Budget.ChargeText(item);
                        texts.Add(item);
                    }
                }
            }
            return _sharedStringTexts = texts;
        }

        /// <summary>Loads a template from a stream (copied into memory) and parses its structure.</summary>
        public static TemplateSource Load(Stream stream, Reading.ReaderOptions limits)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            byte[] bytes;
            if (stream.CanSeek)
            {
                // The size is known, so read into an array of exactly that size: growing a buffer and then
                // copying it out would hold the file three times over for a moment.
                var length = stream.Length - stream.Position;
                if (length > int.MaxValue - 64)
                    throw new InvalidDataException("The workbook is larger than 2 GB, which cannot be loaded as a template or for editing.");
                bytes = new byte[(int)length];
                var filled = 0;
                int read;
                while (filled < bytes.Length && (read = stream.Read(bytes, filled, bytes.Length - filled)) > 0) filled += read;
                if (filled < bytes.Length) Array.Resize(ref bytes, filled); // The stream ended before its reported length.
            }
            else
            {
                using (var ms = new MemoryStream())
                {
                    stream.CopyTo(ms);
                    bytes = ms.ToArray();
                }
            }
            var source = new TemplateSource(bytes, limits);
            source.Parse();
            return source;
        }

        /// <summary>Opens a fresh read-only archive over the template bytes. Each caller gets its own.</summary>
        public ZipArchive OpenArchive()
        {
            try
            {
                return ZipDirectory.Open(new MemoryStream(_bytes, writable: false));
            }
            catch (InvalidDataException ex)
            {
                throw new InvalidDataException("The template is not an .xlsx workbook (it is not a valid zip package). Save it from Excel as .xlsx first.", ex);
            }
        }

        /// <summary>
        /// Opens a part of the template as an XML reader, or returns null if it does not exist. This is the only
        /// way template XML is read, so the limits hold for every caller: long tags, long CDATA sections and
        /// deep nesting are refused by the reader itself.
        /// </summary>
        /// <param name="archive">An archive from <see cref="OpenArchive"/>.</param>
        /// <param name="part">Part path inside the package.</param>
        /// <param name="held">
        /// True when the caller keeps what it reads (shared strings, styles, the workbook, relationships, a sheet
        /// loaded for editing): the part is then capped at <see cref="Reading.ReaderOptions.MaxLoadedPartBytes"/>.
        /// False for parts that are streamed through to the output, capped at
        /// <see cref="Reading.ReaderOptions.MaxStreamedPartBytes"/>.
        /// </param>
        public XmlReader? OpenXml(ZipArchive archive, string part, bool held = false)
        {
            var entry = archive.GetEntry(part);
            if (entry == null) return null;
            var text = new GuardedPartReader(entry.Open(), part, held ? Limits.MaxLoadedPartBytes : Limits.MaxStreamedPartBytes, held, Limits.MaxTokenChars, Limits.MaxTagChars);
            return new DepthGuardedXmlReader(XmlReader.Create(text, XmlCopy.ReaderSettings()), () => ReadFailed?.Invoke());
        }

        // Names of the entries in the package, read once: looking a part up must not cost a pass over the
        // archive's directory each time.
        private HashSet<string>? _partNames;

        /// <summary>True when the package contains a part with exactly this name.</summary>
        public bool HasPart(string part)
        {
            if (_partNames == null)
            {
                var names = new HashSet<string>(StringComparer.Ordinal);
                using (var archive = OpenArchive())
                    foreach (var entry in archive.Entries) names.Add(entry.FullName);
                _partNames = names;
            }
            return part.Length > 0 && _partNames.Contains(part);
        }

        /// <summary>
        /// True when <paramref name="name"/> is a name a package part can have and this library is willing to
        /// write into its output: a relative path of plain segments. Refused are empty names, names that
        /// start with a slash or contain a backslash, "." and ".." segments (a zip written with those can
        /// place a file outside the folder it is unpacked into), control characters, and names too long to be
        /// a part of any real workbook.
        /// </summary>
        public static bool IsWritablePartName(string name)
        {
            if (name.Length == 0 || name.Length > 1024 || name[0] == '/') return false;
            var segmentStart = 0;
            for (var i = 0; i <= name.Length; i++)
            {
                if (i < name.Length)
                {
                    var c = name[i];
                    if (c == '\\' || c < ' ') return false;
                    if (c != '/') continue;
                }
                var length = i - segmentStart;
                if (length == 0) return false;
                if (name[segmentStart] == '.' && (length == 1 || (length == 2 && name[segmentStart + 1] == '.'))) return false;
                segmentStart = i + 1;
            }
            return true;
        }

        /// <summary>
        /// Opens a part the workbook says it has. A relationship that points at a part which is not in the
        /// package is damage, and is reported as such instead of surfacing later as a null reference.
        /// </summary>
        public XmlReader OpenRequiredXml(ZipArchive archive, string part, bool held = false)
            => OpenXml(archive, part, held)
               ?? throw new InvalidDataException($"The workbook refers to a part '{part}' that is not in the file. The file is damaged.");

        /// <summary>Relationships of a part: id, type and target resolved to a package path.</summary>
        public List<(string Id, string Type, string Target)> ReadRelationships(ZipArchive archive, string sourcePart)
        {
            var folder = sourcePart.Contains("/") ? sourcePart.Substring(0, sourcePart.LastIndexOf('/') + 1) : "";
            var relsPart = folder + "_rels/" + sourcePart.Substring(folder.Length) + ".rels";
            var list = new List<(string, string, string)>();
            using (var r = OpenXml(archive, relsPart, held: true))
            {
                if (r == null) return list;
                while (r.Read())
                {
                    if (r.NodeType != XmlNodeType.Element || r.LocalName != "Relationship") continue;
                    if (r.GetAttribute("TargetMode") == "External") continue;
                    var id = r.GetAttribute("Id") ?? "";
                    var type = r.GetAttribute("Type") ?? "";
                    var target = Resolve(folder, r.GetAttribute("Target") ?? "");
                    Budget.Charge(2L * (id.Length + type.Length + target.Length) + 128);
                    list.Add((id, type, target));
                }
            }
            return list;
        }

        /// <summary>Path of the <c>.rels</c> part that belongs to <paramref name="part"/>.</summary>
        public static string RelsPartOf(string part)
        {
            var folder = part.Contains("/") ? part.Substring(0, part.LastIndexOf('/') + 1) : "";
            return folder + "_rels/" + part.Substring(folder.Length) + ".rels";
        }

        private void Parse()
        {
            using (var archive = OpenArchive())
            {
                foreach (var rel in ReadRelationships(archive, ""))
                    if (rel.Type.EndsWith("/officeDocument", StringComparison.Ordinal)) WorkbookPart = rel.Target;
                // ReadRelationships("") looks for "_rels/.rels", which is exactly the package root relationships.
                WorkbookFolder = WorkbookPart.Contains("/") ? WorkbookPart.Substring(0, WorkbookPart.LastIndexOf('/') + 1) : "";

                var targets = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (var rel in ReadRelationships(archive, WorkbookPart))
                {
                    targets[rel.Id] = rel.Target;
                    if (rel.Type.EndsWith("/sharedStrings", StringComparison.Ordinal)) SharedStringsPart = rel.Target;
                    else if (rel.Type.EndsWith("/styles", StringComparison.Ordinal)) StylesPart = rel.Target;
                    else if (rel.Type.EndsWith("/calcChain", StringComparison.Ordinal)) CalcChainPart = rel.Target;
                }

                var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var sheetParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using (var r = OpenXml(archive, WorkbookPart, held: true) ?? throw new InvalidDataException("The template has no workbook part."))
                {
                    while (r.Read())
                    {
                        if (r.NodeType != XmlNodeType.Element) continue;
                        // Only the workbook's own workbookPr: Excel writes an x15:workbookPr in an extension
                        // list as well, and that one must not reset the date system. See ExcelReader.LoadWorkbook.
                        if (r.LocalName == "workbookPr")
                        {
                            if (r.Depth == 1 && r.NamespaceURI == OpenXmlNamespaces.Main)
                            {
                                var d = r.GetAttribute("date1904");
                                Date1904 = d == "1" || string.Equals(d, "true", StringComparison.OrdinalIgnoreCase);
                            }
                            continue;
                        }
                        if (r.LocalName != "sheet" || r.NamespaceURI != OpenXmlNamespaces.Main) continue;
                        var rid = r.GetAttribute("id", OpenXmlNamespaces.Relationships) ?? "";
                        if (targets.TryGetValue(rid, out var part))
                        {
                            var name = r.GetAttribute("name") ?? "";
                            Budget.Charge(4L * (name.Length + part.Length) + 192);
                            // Excel allows neither; here two sheets with one name could not be told apart, and
                            // two sheets sharing a part would be written to the output twice.
                            if (!sheetNames.Add(name))
                                throw new InvalidDataException($"The workbook has more than one sheet named '{name}'. The file is damaged.");
                            if (!sheetParts.Add(part))
                                throw new InvalidDataException($"Two sheets of the workbook point at the same part '{part}'. The file is damaged.");
                            Sheets.Add((name, part));
                        }
                    }
                }

                if (StylesPart == null) throw new InvalidDataException("The template has no styles part. Open it in Excel and save it again.");
                if (SharedStringsPart != null) ParseSharedStrings(archive);
            }
        }

        /// <summary>Counts the template's shared strings and indexes the plain ones for reuse.</summary>
        private void ParseSharedStrings(ZipArchive archive)
        {
            using (var r = OpenXml(archive, SharedStringsPart!, held: true))
            {
                if (r == null) return;
                // The table as a whole is bounded by MaxLoadedPartBytes; each item by the cell text limit.
                var state = new Reading.ExcelReader.TextState(Limits.MaxCellTextLength, long.MaxValue);
                var index = 0;
                while (r.Read())
                {
                    // Depth 1 only, matching what SharedStringTable.WriteTo copies: counting an "si" nested in
                    // an extension element would shift the index of every string appended after it.
                    if (r.NodeType != XmlNodeType.Element || r.LocalName != "si" || r.Depth != 1) continue;
                    if (!r.IsEmptyElement)
                    {
                        // A plain item is exactly <si><t>text</t></si>; rich-text items are counted but not reused.
                        var itemDepth = r.Depth;
                        var tCount = 0;
                        var other = false;
                        string? text = null;
                        state.BeginCell();
                        r.Read();
                        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == itemDepth))
                        {
                            if (r.NodeType != XmlNodeType.Element) { r.Read(); continue; }
                            if (r.LocalName != "t") { other = true; r.Skip(); continue; }
                            tCount++;
                            if (r.IsEmptyElement) { text = ""; r.Read(); }
                            else text = Reading.ExcelReader.ReadElementText(r, state); // Leaves the reader after </t>.
                        }
                        if (tCount == 1 && !other && text != null && !PlainSharedStrings.ContainsKey(text))
                        {
                            Budget.ChargeText(text);
                            Budget.Charge(64); // The dictionary entry.
                            PlainSharedStrings[text] = index;
                        }
                    }
                    index++;
                }
                SharedStringCount = index;
            }
        }

        /// <summary>
        /// Resolves a relationship target against the source part's folder ("worksheets/sheet1.xml",
        /// "/xl/x.xml", "../x.xml") to a path inside the package.
        /// </summary>
        /// <remarks>
        /// One pass, building the result as a stack. An earlier version removed segments from the front of a
        /// list, which made a target of half a million "./" segments (six kilobytes once zipped) take most of
        /// a minute. ".." above the root is dropped, so the result never climbs out of the package.
        /// </remarks>
        public static string Resolve(string folder, string target)
        {
            var path = target.StartsWith("/", StringComparison.Ordinal) ? target : folder + target;
            var kept = new List<string>();
            var start = 0;
            while (start <= path.Length)
            {
                var slash = path.IndexOf('/', start);
                if (slash < 0) slash = path.Length;
                var length = slash - start;
                if (length == 2 && path[start] == '.' && path[start + 1] == '.')
                {
                    if (kept.Count > 0) kept.RemoveAt(kept.Count - 1);
                }
                else if (length > 0 && !(length == 1 && path[start] == '.'))
                {
                    kept.Add(path.Substring(start, length));
                }
                start = slash + 1;
            }
            return string.Join("/", kept);
        }

        /// <summary>Parses an integer attribute value, or returns null.</summary>
        public static int? Int(string? s)
            => s != null && int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : (int?)null;
    }
}
