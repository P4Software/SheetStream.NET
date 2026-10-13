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
    /// keep memory flat. Not thread-safe.
    /// </remarks>
    public sealed class SharedStringTable
    {
        private readonly Dictionary<string, int> _index = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<string> _strings = new List<string>();
        private long _references;

        internal SharedStringTable()
        {
        }

        /// <summary>Number of distinct strings stored.</summary>
        public int UniqueCount => _strings.Count;

        /// <summary>Total number of cells that reference the table.</summary>
        public long ReferenceCount => _references;

        /// <summary>
        /// Returns the index of <paramref name="value"/>, adding it on first sight. The value must already be
        /// sanitised with <see cref="XmlTextUtil.Sanitize"/>.
        /// </summary>
        internal int GetOrAdd(string value)
        {
            _references++;
            if (_index.TryGetValue(value, out var id)) return id;
            id = _strings.Count;
            _strings.Add(value);
            _index[value] = id;
            return id;
        }

        /// <summary>Writes the <c>sharedStrings.xml</c> part.</summary>
        internal void WriteTo(XmlWriter w)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("sst", OpenXmlNamespaces.Main);
            w.WriteAttributeString("count", _references.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("uniqueCount", _strings.Count.ToString(CultureInfo.InvariantCulture));
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
