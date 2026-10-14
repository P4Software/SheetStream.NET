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
    internal sealed class TemplateSource
    {
        private readonly byte[] _bytes;

        private TemplateSource(byte[] bytes)
        {
            _bytes = bytes;
        }

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

        /// <summary>Loads a template from a stream (copied into memory) and parses its structure.</summary>
        public static TemplateSource Load(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            byte[] bytes;
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                bytes = ms.ToArray();
            }
            var source = new TemplateSource(bytes);
            source.Parse();
            return source;
        }

        /// <summary>Opens a fresh read-only archive over the template bytes. Each caller gets its own.</summary>
        public ZipArchive OpenArchive()
        {
            try
            {
                return new ZipArchive(new MemoryStream(_bytes, writable: false), ZipArchiveMode.Read);
            }
            catch (InvalidDataException ex)
            {
                throw new InvalidDataException("The template is not an .xlsx workbook (it is not a valid zip package). Save it from Excel as .xlsx first.", ex);
            }
        }

        /// <summary>Opens a part of the template as an XML reader, or returns null if it does not exist.</summary>
        public XmlReader? OpenXml(ZipArchive archive, string part)
        {
            var entry = archive.GetEntry(part);
            return entry == null ? null : XmlReader.Create(entry.Open(), XmlCopy.ReaderSettings);
        }

        /// <summary>Relationships of a part: id, type and target resolved to a package path.</summary>
        public List<(string Id, string Type, string Target)> ReadRelationships(ZipArchive archive, string sourcePart)
        {
            var folder = sourcePart.Contains("/") ? sourcePart.Substring(0, sourcePart.LastIndexOf('/') + 1) : "";
            var relsPart = folder + "_rels/" + sourcePart.Substring(folder.Length) + ".rels";
            var list = new List<(string, string, string)>();
            using (var r = OpenXml(archive, relsPart))
            {
                if (r == null) return list;
                while (r.Read())
                {
                    if (r.NodeType != XmlNodeType.Element || r.LocalName != "Relationship") continue;
                    if (r.GetAttribute("TargetMode") == "External") continue;
                    list.Add((r.GetAttribute("Id") ?? "", r.GetAttribute("Type") ?? "", Resolve(folder, r.GetAttribute("Target") ?? "")));
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

                using (var r = OpenXml(archive, WorkbookPart) ?? throw new InvalidDataException("The template has no workbook part."))
                {
                    while (r.Read())
                    {
                        if (r.NodeType != XmlNodeType.Element || r.LocalName != "sheet") continue;
                        var rid = r.GetAttribute("id", OpenXmlNamespaces.Relationships) ?? "";
                        if (targets.TryGetValue(rid, out var part)) Sheets.Add((r.GetAttribute("name") ?? "", part));
                    }
                }

                if (StylesPart == null) throw new InvalidDataException("The template has no styles part. Open it in Excel and save it again.");
                if (SharedStringsPart != null) ParseSharedStrings(archive);
            }
        }

        /// <summary>Counts the template's shared strings and indexes the plain ones for reuse.</summary>
        private void ParseSharedStrings(ZipArchive archive)
        {
            using (var r = OpenXml(archive, SharedStringsPart!))
            {
                if (r == null) return;
                var index = 0;
                while (r.Read())
                {
                    if (r.NodeType != XmlNodeType.Element || r.LocalName != "si") continue;
                    if (!r.IsEmptyElement)
                    {
                        // A plain item is exactly <si><t>text</t></si>; rich-text items are counted but not reused.
                        var sub = r.ReadSubtree();
                        sub.Read();
                        var tCount = 0;
                        var other = false;
                        string? text = null;
                        while (sub.Read())
                        {
                            if (sub.NodeType != XmlNodeType.Element) continue;
                            if (sub.LocalName == "t" && sub.Depth == 1) { tCount++; text = sub.ReadElementContentAsString(); }
                            else if (sub.Depth == 1) other = true;
                        }
                        sub.Dispose();
                        if (tCount == 1 && !other && text != null && !PlainSharedStrings.ContainsKey(text)) PlainSharedStrings[text] = index;
                    }
                    index++;
                }
                SharedStringCount = index;
            }
        }

        /// <summary>Resolves a relationship target against the source part's folder.</summary>
        public static string Resolve(string folder, string target)
        {
            if (target.StartsWith("/", StringComparison.Ordinal)) return target.Substring(1);
            var parts = new List<string>((folder + target).Split('/'));
            for (var i = 0; i < parts.Count; i++)
            {
                if (parts[i] == "..") { parts.RemoveAt(i); if (i > 0) { parts.RemoveAt(i - 1); i--; } i--; }
                else if (parts[i] == "." || parts[i].Length == 0) { parts.RemoveAt(i); i--; }
            }
            return string.Join("/", parts);
        }

        /// <summary>Parses an integer attribute value, or returns null.</summary>
        public static int? Int(string? s)
            => s != null && int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : (int?)null;
    }
}
