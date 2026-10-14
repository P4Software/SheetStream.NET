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
    /// sheets are being streamed. For data sets dominated by unique strings, set
    /// <see cref="PackageOptions.UseSharedStrings"/> to <c>false</c> to write inline strings instead and
    /// keep memory flat. When filling a template, the template's own strings keep their indexes and new
    /// strings are appended after them. Not thread-safe.
    /// </remarks>
    public sealed class SharedStringTable
    {
        private readonly Dictionary<string, int> _index = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<string> _strings = new List<string>();
        private long _references;

        // Template mode: the template's items occupy indexes 0..._templateCount-1 and are copied verbatim.
        private int _templateCount;

        internal SharedStringTable()
        {
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
        /// Returns the index of <paramref name="value"/>, adding it on first sight. The value must already be
        /// sanitised with <see cref="XmlTextUtil.Sanitize"/>.
        /// </summary>
        internal int GetOrAdd(string value)
        {
            _references++;
            if (_index.TryGetValue(value, out var id)) return id;
            id = _templateCount + _strings.Count;
            _strings.Add(value);
            _index[value] = id;
            return id;
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
