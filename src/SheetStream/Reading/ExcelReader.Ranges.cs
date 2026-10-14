using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream.Reading
{
    /// <summary>A rectangle of merged cells read from a sheet. Only the top-left cell of a merge carries a value.</summary>
    public sealed class MergedRange
    {
        internal MergedRange(int fromRow, int fromColumn, int toRow, int toColumn, string range)
        {
            FromRow = fromRow;
            FromColumn = fromColumn;
            ToRow = toRow;
            ToColumn = toColumn;
            Range = range;
        }

        /// <summary>1-based first row of the merge.</summary>
        public int FromRow { get; }

        /// <summary>1-based first column of the merge (column A is 1).</summary>
        public int FromColumn { get; }

        /// <summary>1-based last row of the merge.</summary>
        public int ToRow { get; }

        /// <summary>1-based last column of the merge.</summary>
        public int ToColumn { get; }

        /// <summary>The merge in A1 form, such as <c>"A1:D1"</c>.</summary>
        public string Range { get; }

        /// <summary>True when the 1-based cell lies inside the merge (its top-left cell included).</summary>
        public bool Contains(int row, int column)
            => row >= FromRow && row <= ToRow && column >= FromColumn && column <= ToColumn;

        /// <inheritdoc />
        public override string ToString() => Range;
    }

    /// <summary>
    /// A hyperlink read from a sheet. Exactly one of <see cref="Url"/> (web, mail or file target) and
    /// <see cref="Location"/> (a place in the same workbook) is normally set.
    /// </summary>
    public sealed class ExcelHyperlink
    {
        internal ExcelHyperlink(int row, int column, string? url, string? location, string? tooltip, string? display)
        {
            Row = row;
            Column = column;
            Address = CellReference.ToA1(row, column);
            Url = url;
            Location = location;
            Tooltip = tooltip;
            Display = display;
        }

        /// <summary>1-based row of the linked cell.</summary>
        public int Row { get; }

        /// <summary>1-based column of the linked cell (column A is 1).</summary>
        public int Column { get; }

        /// <summary>The linked cell in A1 form, such as <c>"B7"</c>.</summary>
        public string Address { get; }

        /// <summary>External target (<c>https://</c>, <c>mailto:</c>, a file path), resolved from the sheet's relationships. Null for in-workbook links.</summary>
        public string? Url { get; }

        /// <summary>In-workbook target such as <c>"'Summary'!A1"</c> or a defined name. Null for external links.</summary>
        public string? Location { get; }

        /// <summary>Hover text, when the file carries one.</summary>
        public string? Tooltip { get; }

        /// <summary>
        /// Display text stored on the link itself, when present. Excel writes it for links created through its
        /// UI; files written by code usually leave it out, in which case the cell value is the visible text.
        /// </summary>
        public string? Display { get; }

        /// <inheritdoc />
        public override string ToString() => Address + " -> " + (Url ?? Location ?? "");
    }

    // Merged cells and hyperlinks: the sheet metadata that follows sheetData. (The class summary lives on the
    // main part in ExcelReader.cs; a second one here would be merged into it by the doc generator.)
    public sealed partial class ExcelReader
    {
        /// <summary>
        /// Lists the merged cell ranges of a sheet by name (case-insensitive). The sheet's rows are skipped, not
        /// parsed, so this is cheap even on a sheet with a million rows. A sheet without merges yields an empty list.
        /// </summary>
        public IReadOnlyList<MergedRange> ReadMergedRanges(string sheetName) => ReadMergedRanges(FindSheet(sheetName));

        /// <summary>Lists the merged cell ranges of a sheet by zero-based position.</summary>
        public IReadOnlyList<MergedRange> ReadMergedRanges(int sheetIndex)
        {
            var merges = new List<MergedRange>();
            ScanAfterSheetData(sheetIndex, merges, null);
            return merges;
        }

        /// <summary>
        /// Lists the hyperlinks of a sheet by name (case-insensitive). External targets are resolved through the
        /// sheet's relationships; in-workbook links surface as <see cref="ExcelHyperlink.Location"/>. The sheet's
        /// rows are skipped, not parsed. A sheet without links yields an empty list.
        /// </summary>
        public IReadOnlyList<ExcelHyperlink> ReadHyperlinks(string sheetName) => ReadHyperlinks(FindSheet(sheetName));

        /// <summary>Lists the hyperlinks of a sheet by zero-based position.</summary>
        public IReadOnlyList<ExcelHyperlink> ReadHyperlinks(int sheetIndex)
        {
            var links = new List<ExcelHyperlink>();
            ScanAfterSheetData(sheetIndex, null, links);
            return links;
        }

        /// <summary>
        /// Walks a sheet part collecting <c>mergeCells</c> and/or <c>hyperlinks</c> into whichever lists are
        /// non-null. <c>sheetData</c> is skipped as a whole so the cost does not grow with the row count.
        /// Malformed entries are ignored rather than failing the whole call: a damaged merge reference should
        /// not stop an import that only wanted the links.
        /// </summary>
        private void ScanAfterSheetData(int sheetIndex, List<MergedRange>? merges, List<ExcelHyperlink>? links)
        {
            if (sheetIndex < 0 || sheetIndex >= _sheets.Count) throw new ArgumentOutOfRangeException(nameof(sheetIndex), sheetIndex, $"The workbook has {_sheets.Count} sheet(s).");
            var part = _sheets[sheetIndex].Part;
            Dictionary<string, string>? externalTargets = null; // Loaded lazily: only sheets with external links pay for the .rels read.

            using (var xr = OpenXml(part) ?? throw new InvalidDataException($"Sheet '{_sheets[sheetIndex].Name}' points at missing part '{part}'."))
            {
                // Cursor discipline: every branch leaves the reader on the next unread node, so the loop never
                // calls Read() after a method that already advanced it (Skip and the collectors both do).
                xr.Read();
                while (!xr.EOF)
                {
                    if (xr.NodeType != XmlNodeType.Element) { xr.Read(); continue; }

                    if (xr.LocalName == "sheetData")
                    {
                        xr.Skip(); // The bulk of the file; Skip advances past </sheetData> without materialising cells.
                        continue;
                    }
                    if (merges != null && xr.LocalName == "mergeCells")
                    {
                        CollectMerges(xr, merges);
                        continue;
                    }
                    if (links != null && xr.LocalName == "hyperlinks")
                    {
                        CollectHyperlinks(xr, links, part, ref externalTargets);
                        continue;
                    }
                    xr.Read();
                }
            }
        }

        /// <summary>Reads the children of a <c>mergeCells</c> element (positioned on its start) and leaves the reader on the node after it.</summary>
        private static void CollectMerges(XmlReader xr, List<MergedRange> merges)
        {
            if (xr.IsEmptyElement) { xr.Read(); return; }
            var depth = xr.Depth;
            xr.Read();
            while (!xr.EOF && !(xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth))
            {
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "mergeCell")
                {
                    var reference = xr.GetAttribute("ref");
                    if (reference != null && TryParseRange(reference, out var r))
                        merges.Add(new MergedRange(r.FromRow, r.FromColumn, r.ToRow, r.ToColumn, r.ToString()));
                    xr.Skip(); // A mergeCell has no children we care about; Skip also handles the empty-element form.
                    continue;
                }
                xr.Read();
            }
            xr.Read(); // Step past </mergeCells>.
        }

        /// <summary>Reads the children of a <c>hyperlinks</c> element (positioned on its start) and leaves the reader on the node after it.</summary>
        private void CollectHyperlinks(XmlReader xr, List<ExcelHyperlink> links, string sheetPart, ref Dictionary<string, string>? externalTargets)
        {
            if (xr.IsEmptyElement) { xr.Read(); return; }
            var depth = xr.Depth;
            xr.Read();
            while (!xr.EOF && !(xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth))
            {
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "hyperlink")
                {
                    // Excel allows a ranged ref ("A1:B2"); the link is reported on its top-left cell.
                    var reference = xr.GetAttribute("ref");
                    if (reference != null && TryParseRange(reference, out var r))
                    {
                        string? url = null;
                        var rid = xr.GetAttribute("id", OpenXmlNamespaces.Relationships);
                        if (rid != null)
                        {
                            if (externalTargets == null) externalTargets = LoadHyperlinkTargets(sheetPart);
                            externalTargets.TryGetValue(rid, out url); // A dangling id simply yields no URL.
                        }
                        var location = xr.GetAttribute("location");
                        var tooltip = xr.GetAttribute("tooltip");
                        var display = xr.GetAttribute("display");
                        links.Add(new ExcelHyperlink(r.FromRow, r.FromColumn, url,
                            string.IsNullOrEmpty(location) ? null : location,
                            tooltip == null ? null : Decode(tooltip),
                            display == null ? null : Decode(display)));
                    }
                    xr.Skip();
                    continue;
                }
                xr.Read();
            }
            xr.Read(); // Step past </hyperlinks>.
        }

        /// <summary>
        /// Maps relationship ids to the external targets of a sheet's <c>/hyperlink</c> relationships. The
        /// targets are kept verbatim: they are absolute URIs (or user file paths), not package-relative parts,
        /// so <see cref="Resolve"/> must not touch them.
        /// </summary>
        private Dictionary<string, string> LoadHyperlinkTargets(string sheetPart)
        {
            // Same convention LoadWorkbook uses for the workbook: "<folder>/_rels/<file>.rels".
            var folder = sheetPart.Contains("/") ? sheetPart.Substring(0, sheetPart.LastIndexOf('/') + 1) : "";
            var relsPart = folder + "_rels/" + sheetPart.Substring(folder.Length) + ".rels";
            var targets = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var rel in ReadRelationships(relsPart))
                if (rel.Type.EndsWith("/hyperlink", StringComparison.Ordinal)) targets[rel.Id] = rel.Target;
            return targets;
        }

        /// <summary>Parses an A1 cell or range, returning false (instead of throwing) for anything malformed.</summary>
        private static bool TryParseRange(string text, out RangeReference range)
        {
            try
            {
                range = RangeReference.Parse(text);
                return true;
            }
            catch (ArgumentException)
            {
                // RangeReference.Parse reports bad input through ArgumentException (and its subclass
                // ArgumentOutOfRangeException for out-of-sheet corners); both mean "ignore this entry".
                range = default;
                return false;
            }
        }
    }
}
