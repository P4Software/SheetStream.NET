using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>
    /// The workbook-wide shared string table (SST). Every distinct string that goes through it is stored once and cells
    /// refer to it by index, which is both what the specification recommends and what keeps repeated values
    /// (customer names, product descriptions) from bloating the file.
    /// </summary>
    /// <remarks>
    /// The table lives in memory until the package is saved because its index must be known while the
    /// sheets are being streamed. It stops growing at <see cref="PackageOptions.MaxSharedStringChars"/>;
    /// new strings beyond that are written inline, so an export dominated by unique strings cannot run the
    /// process out of memory. With <see cref="PackageOptions.StringStorage"/> at its default (<see cref="StringStorage.Auto"/>)
    /// only the columns whose text repeats feed the table, besides the first cells of every column, which are looked up to
    /// judge it; <see cref="StringStorage.Inline"/> (or <see cref="PackageOptions.UseSharedStrings"/> set to <c>false</c>)
    /// writes every string inline from the start. When filling a template, the template's own strings keep their indexes
    /// and new strings are appended after them. Not thread-safe.
    /// </remarks>
    public sealed class SharedStringTable
    {
        private readonly Dictionary<string, int> _index = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<string> _strings = new List<string>();
        private long _references;

        // Template mode: the template's items occupy indexes 0..._templateCount-1 and are copied verbatim.
        private int _templateCount;

        // Memory budget, in characters, for the strings this package adds, and how much of it is used.
        private readonly long _budget;
        private long _used;

        /// <summary>
        /// What one entry costs beyond its characters, in the budget's character units: the dictionary
        /// entry, the list slot and the string header. Without it a flood of very short unique strings
        /// (ids, codes) would use far more memory than the budget suggests.
        /// </summary>
        private const int EntryOverhead = 32;

        private const string ItemOpen = "<si><t>";
        private const string ItemOpenPreserve = "<si><t xml:space=\"preserve\">";
        private const string ItemClose = "</t></si>";

        internal SharedStringTable(long budget)
        {
            _budget = budget;
        }

        /// <summary>Number of distinct strings stored (template strings included).</summary>
        public int UniqueCount => _templateCount + _strings.Count;

        /// <summary>Total number of cells written through the table by this package.</summary>
        public long ReferenceCount => _references;

        /// <summary>
        /// Switches to template mode: new strings are numbered after the template's, and identical plain
        /// template strings are reused instead of being added again.
        /// </summary>
        internal void AttachTemplate(int templateCount, IDictionary<string, int> plainTemplateStrings)
        {
            _templateCount = templateCount;
            foreach (var pair in plainTemplateStrings) _index[pair.Key] = pair.Value;
        }

        /// <summary>
        /// Looks up the index of <paramref name="value"/>, adding it on first sight. The value must already be
        /// sanitised with <see cref="XmlTextUtil.Sanitize"/>.
        /// </summary>
        /// <returns>
        /// False when <paramref name="value"/> is new and the table has used up
        /// <see cref="PackageOptions.MaxSharedStringChars"/>; the caller then writes the string inline, which
        /// keeps memory flat however many unique strings the export has.
        /// </returns>
        internal bool TryGetOrAdd(string value, out int id) => TryGetOrAdd(value, out id, out _);

        /// <summary>
        /// As <see cref="TryGetOrAdd(string, out int)"/>, and says through <paramref name="existed"/> whether the
        /// value was already in the table (a repeat) or has just been added, which is what tells a column of
        /// repeating values from a column of unique ones.
        /// </summary>
        internal bool TryGetOrAdd(string value, out int id, out bool existed)
        {
#if NET8_0_OR_GREATER
            // One hash of the text and one probe for both the lookup and the insert, where TryGetValue followed by the indexer
            // hashes and probes twice; most cells that reach the table are new texts (a column is sampled before it is judged).
            ref var slot = ref System.Runtime.InteropServices.CollectionsMarshal.GetValueRefOrAddDefault(_index, value, out var found);
            if (found)
            {
                id = slot;
                _references++;
                existed = true;
                return true;
            }
            existed = false;
            var newCost = value.Length + EntryOverhead;
            if (_used + newCost > _budget)
            {
                _index.Remove(value); // Leaves no trace of the placeholder entry the call above made.
                id = 0;
                return false;
            }
            _used += newCost;
            _references++;
            id = slot = _templateCount + _strings.Count;
            _strings.Add(value);
            return true;
#else
            if (_index.TryGetValue(value, out id))
            {
                _references++;
                existed = true;
                return true;
            }
            existed = false;
            var cost = value.Length + EntryOverhead;
            if (_used + cost > _budget) return false;
            _used += cost;
            _references++;
            id = _templateCount + _strings.Count;
            _strings.Add(value);
            _index[value] = id;
            return true;
#endif
        }

        // ---- What a sheet counts as a repeat (StringStorage.Auto) ---------------------------------------------------------
        //
        // The table is shared by every sheet of the package, but the policy that judges a column asks "did this text
        // repeat?" about the sheet it is writing. A text that an EARLIER sheet put in the table is not a repeat in this
        // sheet's column the first time the sheet uses it: if it were, a unique column written on a second sheet (the
        // same ids listed twice) would be judged shared because of the first sheet's entries, and its whole length would
        // go into the table. So each sheet takes a scope: the table's size when the sheet's first text arrives (its
        // watermark) and a stamp of its own. An entry at or above the watermark was added by this sheet, so a lookup that
        // finds it is a repeat, as always. An entry below it is an earlier sheet's (or the template's): its first use in
        // this sheet is reported as new and marks it with the stamp, and from then on it is a repeat. The judge therefore
        // sees what it would see if every sheet started with an empty table, while the cells still share the entries.
        // One sheet of a new package has watermark 0, so nothing changes for it.

        // Per entry below the current watermark: the stamp of the last scope that used it (0 = none). Allocated when a
        // scope starts with a non-empty table, so a package with one sheet never has it. It grows geometrically, never
        // to the exact size: sized exactly at every sheet it would be reallocated and copied for each of 1,000 small
        // sheets, which is quadratic. It is 4 bytes per entry, 8 at worst while it has room to spare, for the entries this
        // package adds, which the entry's EntryOverhead in the budget (32 characters) more than covers; for the entries
        // of a template, which the budget does not charge, it is the same few bytes on top of what the template itself costs.
        private int[] _usedBy = Array.Empty<int>();
        private int _scopes;

        /// <summary>
        /// Starts the scope of a sheet that is about to look texts up (see the notes above) and returns its stamp.
        /// <paramref name="watermark"/> is the number of entries the table holds now: those are not this sheet's.
        /// </summary>
        internal int BeginScope(out int watermark)
        {
            watermark = UniqueCount;
            if (_usedBy.Length < watermark)
            {
                // Twice the old length (at least 16, at least what is needed now), but never more than the table can ever hold:
                // the budget allows one entry per EntryOverhead characters, plus the template's own.
                var most = Math.Min((long)int.MaxValue / 2, _templateCount + _budget / EntryOverhead + 1);
                var grown = (int)Math.Max(watermark, Math.Min(most, Math.Max(16L, 2L * _usedBy.Length)));
                Array.Resize(ref _usedBy, grown);
            }
            return ++_scopes;
        }

        /// <summary>
        /// As <see cref="TryGetOrAdd(string, out int, out bool)"/>, but <paramref name="repeated"/> answers for the sheet whose
        /// scope this is: true when the text is one this sheet already looked up or added, false when it is new to the
        /// sheet (added now, or an entry of an earlier sheet that this sheet uses for the first time).
        /// </summary>
        internal bool TryGetOrAdd(string value, out int id, out bool repeated, int watermark, int stamp)
        {
            if (!TryGetOrAdd(value, out id, out var existed))
            {
                repeated = false;
                return false;
            }
            if (!existed || id >= watermark)
            {
                repeated = existed;
            }
            else if (_usedBy[id] == stamp)
            {
                repeated = true;
            }
            else
            {
                _usedBy[id] = stamp;
                repeated = false;
            }
            return true;
        }

        /// <summary>Writes the <c>sharedStrings.xml</c> part.</summary>
        internal void WriteTo(XmlWriter w) => WriteTo(w, null);

        /// <summary>
        /// Writes the shared strings part. In template mode the template's <c>si</c> items are streamed from
        /// <paramref name="template"/> first, unchanged, so rich-text and phonetic strings survive.
        /// </summary>
        internal void WriteTo(XmlWriter w, XmlReader? template)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("sst", OpenXmlNamespaces.Main);
            // "count" is the number of cell references; the template's own count is unknown, so report the
            // unique count as a lower bound, which Excel accepts (it recomputes on save).
            var count = Math.Max(_references, UniqueCount);
            w.WriteAttributeString("count", count.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("uniqueCount", UniqueCount.ToString(CultureInfo.InvariantCulture));

            if (template != null)
            {
                XmlCopy.MoveToRoot(template);
                template.Read();
                while (!template.EOF && !(template.NodeType == XmlNodeType.EndElement && template.Depth == 0))
                {
                    if (template.NodeType == XmlNodeType.Element && template.LocalName == "si") XmlCopy.CopySubtree(template, w);
                    else if (template.NodeType == XmlNodeType.Element) template.Skip(); // extLst and the like
                    else template.Read();
                }
            }

            WriteItems(w);
            w.WriteEndElement();
            w.WriteEndDocument();
        }

        /// <summary>The characters composed before they are handed to the XML writer in one call (and the longest text composed in them).</summary>
        private const int ChunkChars = 8192;

        /// <summary>
        /// Writes the <c>si</c> items of the table's own strings. Each is composed as characters (the tags, then the text with
        /// <c>&amp;</c>, <c>&lt;</c> and <c>&gt;</c> escaped and each CR as <c>&amp;#xD;</c>, then the closing tags) in a chunk that is handed to the XML writer
        /// as raw text, so the writer does no per-element work: a table has one entry per distinct text, tens of thousands
        /// when a sheet has columns of unique text. The bytes are the ones <c>WriteStartElement</c>, <c>WriteAttributeString</c>,
        /// <c>WriteString</c> and <c>WriteEndElement</c> wrote (the table's texts have been through <see cref="XmlTextUtil.Sanitize"/>,
        /// so they hold nothing the writer would refuse, and it is set to leave line breaks as they are). Without
        /// <c>xml:space="preserve"</c> Excel trims leading and trailing white space on load.
        /// </summary>
        private void WriteItems(XmlWriter w)
        {
            if (_strings.Count == 0) return;
            var chunk = new char[ChunkChars];
            var n = 0;
            foreach (var s in _strings)
            {
                var open = XmlTextUtil.NeedsSpacePreserve(s) ? ItemOpenPreserve : ItemOpen;
                // Worst case: every character becomes a five-character entity.
                if (s.Length > (ChunkChars - ItemOpenPreserve.Length - ItemClose.Length) / 5)
                {
                    // A text longer than a chunk goes through the writer, which escapes it in pieces.
                    if (n > 0) { w.WriteRaw(chunk, 0, n); n = 0; }
                    w.WriteRaw(open);
                    XmlTextUtil.WriteText(w, s);
                    w.WriteRaw(ItemClose);
                    continue;
                }
                if (ChunkChars - n < open.Length + s.Length * 5 + ItemClose.Length)
                {
                    w.WriteRaw(chunk, 0, n);
                    n = 0;
                }

                open.CopyTo(0, chunk, n, open.Length);
                n += open.Length;
                var runStart = 0;
                for (var i = 0; i < s.Length; i++)
                {
                    var c = s[i];
                    // '&', '<', '>' and CR are all at or below '>', so almost every character is rejected by one compare.
                    if (c > '>') continue;
                    string entity;
                    if (c == '&') entity = "&amp;";
                    else if (c == '<') entity = "&lt;";
                    else if (c == '>') entity = "&gt;";
                    else if (c == '\r') entity = "&#xD;"; // A raw CR would be read back as LF (see XmlTextUtil.WriteText).
                    else continue;
                    s.CopyTo(runStart, chunk, n, i - runStart);
                    n += i - runStart;
                    entity.CopyTo(0, chunk, n, entity.Length);
                    n += entity.Length;
                    runStart = i + 1;
                }
                s.CopyTo(runStart, chunk, n, s.Length - runStart);
                n += s.Length - runStart;
                ItemClose.CopyTo(0, chunk, n, ItemClose.Length);
                n += ItemClose.Length;
            }
            if (n > 0) w.WriteRaw(chunk, 0, n);
        }
    }
}
