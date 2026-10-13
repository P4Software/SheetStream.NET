using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>
    /// An Excel table (ListObject): a named, styled range with filter buttons and banded rows that formulas
    /// can address by name (<c>Sales[Revenue]</c>). Create with <see cref="Worksheet.AddTable(string, string, System.Collections.Generic.IReadOnlyList{string})"/>.
    /// </summary>
    /// <remarks>
    /// Excel requires each table column name to equal the text of its header cell. SheetStream captures that
    /// text while the header row is written, so add the table before writing its header row, or pass the
    /// column names explicitly.
    /// </remarks>
    public sealed class ExcelTable
    {
        private readonly string?[] _columnNames;

        internal ExcelTable(Worksheet sheet, int id, string name, RangeReference range, IReadOnlyList<string>? columnNames)
        {
            Sheet = sheet;
            Id = id;
            Name = name;
            Range = range;
            _columnNames = new string?[range.ColumnCount];
            if (columnNames != null)
            {
                if (columnNames.Count != range.ColumnCount)
                    throw new ArgumentException($"Table '{name}' spans {range.ColumnCount} columns but {columnNames.Count} column names were given.", nameof(columnNames));
                for (var i = 0; i < columnNames.Count; i++) _columnNames[i] = columnNames[i];
            }
        }

        internal Worksheet Sheet { get; }

        /// <summary>Workbook-wide table id (also the part number).</summary>
        internal int Id { get; }

        internal RangeReference Range { get; }

        /// <summary>Table name used in structured references, e.g. <c>Sales</c>.</summary>
        public string Name { get; }

        /// <summary>Built-in table style, e.g. "TableStyleMedium2" (default), "TableStyleLight9", "TableStyleDark1".</summary>
        public string StyleName { get; set; } = "TableStyleMedium2";

        /// <summary>Alternate row shading. Default true.</summary>
        public bool ShowRowStripes { get; set; } = true;

        /// <summary>Alternate column shading.</summary>
        public bool ShowColumnStripes { get; set; }

        /// <summary>Emphasise the first column.</summary>
        public bool ShowFirstColumn { get; set; }

        /// <summary>Emphasise the last column.</summary>
        public bool ShowLastColumn { get; set; }

        /// <summary>Show the filter drop-downs in the header. Default true.</summary>
        public bool ShowFilterButtons { get; set; } = true;

        /// <summary>True while header text still has to be captured from the row being written.</summary>
        internal bool NeedsHeaderCapture
        {
            get
            {
                foreach (var n in _columnNames) if (n == null) return true;
                return false;
            }
        }

        /// <summary>Records the text written into a header cell of this table.</summary>
        internal void CaptureHeader(int column, string text)
        {
            var i = column - Range.FromColumn;
            if (i >= 0 && i < _columnNames.Length && _columnNames[i] == null) _columnNames[i] = text;
        }

        /// <summary>Writes <c>xl/tables/tableN.xml</c>. CT_Table order: autoFilter, tableColumns, tableStyleInfo.</summary>
        internal void WriteTo(XmlWriter w)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < _columnNames.Length; i++)
            {
                var n = _columnNames[i];
                var cell = CellReference.ToA1(Range.FromRow, Range.FromColumn + i);
                if (string.IsNullOrEmpty(n))
                    throw new InvalidOperationException($"Table '{Name}' has no header text for column {cell}. Write a header cell there (after adding the table) or pass column names to AddTable.");
                if (!seen.Add(n!))
                    throw new InvalidOperationException($"Table '{Name}' has two columns named '{n}'. Excel requires unique header names.");
            }

            w.WriteStartDocument(true);
            w.WriteStartElement("table", OpenXmlNamespaces.Main);
            w.WriteAttributeString("id", Id.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("name", Name);
            w.WriteAttributeString("displayName", Name);
            w.WriteAttributeString("ref", Range.ToString());
            w.WriteAttributeString("totalsRowShown", "0");

            if (ShowFilterButtons)
            {
                w.WriteStartElement("autoFilter");
                w.WriteAttributeString("ref", Range.ToString());
                w.WriteEndElement();
            }

            w.WriteStartElement("tableColumns");
            w.WriteAttributeString("count", _columnNames.Length.ToString(CultureInfo.InvariantCulture));
            for (var i = 0; i < _columnNames.Length; i++)
            {
                w.WriteStartElement("tableColumn");
                w.WriteAttributeString("id", (i + 1).ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("name", XmlTextUtil.Sanitize(_columnNames[i]!));
                w.WriteEndElement();
            }
            w.WriteEndElement();

            w.WriteStartElement("tableStyleInfo");
            w.WriteAttributeString("name", StyleName);
            w.WriteAttributeString("showFirstColumn", ShowFirstColumn ? "1" : "0");
            w.WriteAttributeString("showLastColumn", ShowLastColumn ? "1" : "0");
            w.WriteAttributeString("showRowStripes", ShowRowStripes ? "1" : "0");
            w.WriteAttributeString("showColumnStripes", ShowColumnStripes ? "1" : "0");
            w.WriteEndElement();

            w.WriteEndElement();
            w.WriteEndDocument();
        }

        /// <summary>
        /// Validates a table name: starts with a letter, underscore or backslash; letters, digits, underscores
        /// and periods only; and does not look like a cell reference (Excel would confuse "AB12" or "R1C1").
        /// </summary>
        internal static void ValidateName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Table name cannot be empty.", nameof(name));
            if (name.Length > 255) throw new ArgumentException($"Table name '{name}' is longer than 255 characters.", nameof(name));
            var first = name[0];
            if (!(char.IsLetter(first) || first == '_' || first == '\\'))
                throw new ArgumentException($"Table name '{name}' must start with a letter, underscore or backslash.", nameof(name));
            foreach (var c in name)
                if (!(char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\\'))
                    throw new ArgumentException($"Table name '{name}' may only contain letters, digits, underscores and periods (no spaces).", nameof(name));
            if (LooksLikeCellReference(name) || string.Equals(name, "R", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "C", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException($"Table name '{name}' looks like a cell reference, which Excel does not allow.", nameof(name));
        }

        private static bool LooksLikeCellReference(string name)
        {
            try
            {
                RangeReference.ParseCell(name);
                return true;
            }
            catch (ArgumentException)
            {
            }
            // R1C1 style: R<digits>C<digits>
            var upper = name.ToUpperInvariant();
            if (upper.Length >= 2 && upper[0] == 'R')
            {
                var i = 1;
                while (i < upper.Length && char.IsDigit(upper[i])) i++;
                if (i < upper.Length && upper[i] == 'C')
                {
                    i++;
                    while (i < upper.Length && char.IsDigit(upper[i])) i++;
                    return i == upper.Length;
                }
            }
            return false;
        }
    }
}
