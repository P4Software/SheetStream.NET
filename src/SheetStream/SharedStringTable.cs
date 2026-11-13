using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>
    /// The workbook-wide shared string table (SST). Every distinct string is stored once and cells refer
    /// to it by index, which is both what the specification recommends and what keeps repeated values
    /// (status codes, category names, customer names) from bloating the file.
    /// </summary>
    /// <remarks>
    /// The table lives in memory until the package is saved because its index must be known while the
    /// sheets are being streamed. It stops growing at <see cref="PackageOptions.MaxSharedStringChars"/>;
    /// new strings beyond that are written inline, so an export dominated by unique strings cannot run the
    /// process out of memory. Set <see cref="PackageOptions.UseSharedStrings"/> to <c>false</c> to write
    /// every string inline from the start. When filling a template, the template's own strings keep their indexes and new
    /// strings are appended after them. Not thread-safe.
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
        internal bool TryGetOrAdd(string value, out int id)
        {
            if (_index.TryGetValue(value, out id))
            {
                _references++;
                return true;
            }
            var cost = value.Length + EntryOverhead;
            if (_used + cost > _budget) return false;
            _used += cost;
            _references++;
            id = _templateCount + _strings.Count;
            _strings.Add(value);
            _index[value] = id;
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

            foreach (var s in _strings)
            {
                w.WriteStartElement("si");
                w.WriteStartElement("t");
                // Without xml:space="preserve" Excel trims leading/trailing whitespace on load.
                if (XmlTextUtil.NeedsSpacePreserve(s)) w.WriteAttributeString("xml", "space", null, "preserve");
                w.WriteString(s);
                w.WriteEndElement();
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndDocument();
        }
    }
}
