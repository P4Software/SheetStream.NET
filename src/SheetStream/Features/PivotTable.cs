using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>Aggregation applied by a pivot data field (<see cref="PivotTable.AddDataField"/>).</summary>
    public enum PivotFunction
    {
        /// <summary>Sum of the values (Excel's default for numeric columns).</summary>
        Sum,
        /// <summary>Number of non-empty values.</summary>
        Count,
        /// <summary>Arithmetic mean.</summary>
        Average,
        /// <summary>Largest value.</summary>
        Max,
        /// <summary>Smallest value.</summary>
        Min,
        /// <summary>Product of the values.</summary>
        Product,
        /// <summary>Number of numeric values.</summary>
        CountNumbers,
        /// <summary>Sample standard deviation.</summary>
        StdDev,
        /// <summary>Population standard deviation.</summary>
        StdDevP,
        /// <summary>Sample variance.</summary>
        Var,
        /// <summary>Population variance.</summary>
        VarP,
    }

    /// <summary>
    /// A pivot table summarising a range (or an Excel table) written earlier in the same workbook. Create it with
    /// <see cref="Worksheet.AddPivotTable(string, string, string, int, int, IReadOnlyList{string})"/> or
    /// <see cref="Worksheet.AddPivotTable(string, ExcelTable, int, int)"/>, then place source columns on the
    /// row, column, filter and data areas by their header text.
    /// </summary>
    /// <remarks>
    /// <para>
    /// SheetStream streams every sheet forward-only, so by the time the pivot is written the source data has
    /// already left memory. The file therefore carries the pivot's layout and an empty cache flagged
    /// <c>refreshOnLoad</c>: Excel reads the source cells and computes the totals the moment the workbook opens,
    /// exactly as if the user had pressed Refresh. Viewers that do not run the pivot engine (some web previews,
    /// LibreOffice in certain versions) show an empty pivot until the file has been opened and saved in Excel.
    /// </para>
    /// <para>
    /// Not supported: calculated fields, grouping, slicers, custom item sorting, reading or editing pivots of an
    /// existing workbook, and pivots on template or edited workbooks.
    /// </para>
    /// </remarks>
    /// <example>
    /// <code>
    /// var pivot = summary.AddPivotTable("ByCustomer", invoicesTable, row: 3, column: 1);
    /// pivot.AddRowField("Customer");
    /// pivot.AddColumnField("Status");
    /// pivot.AddDataField("Amount", PivotFunction.Sum, numberFormat: "$#,##0.00");
    /// pivot.AddDataField("Invoice #", PivotFunction.Count, "Invoices");
    /// </code>
    /// </example>
    public sealed class PivotTable
    {
        private readonly string[] _headers;
        private readonly List<int> _rowFields = new List<int>();
        private readonly List<int> _columnFields = new List<int>();
        private readonly List<int> _pageFields = new List<int>();
        private readonly List<DataField> _dataFields = new List<DataField>();

        /// <summary>One entry of the Values area.</summary>
        private readonly struct DataField
        {
            public DataField(int field, PivotFunction function, string name, string? numberFormat)
            {
                Field = field;
                Function = function;
                Name = name;
                NumberFormat = numberFormat;
            }

            public int Field { get; }
            public PivotFunction Function { get; }
            public string Name { get; }
            public string? NumberFormat { get; }
        }

        internal PivotTable(Worksheet sheet, int number, string name, string sourceSheet, RangeReference sourceRange, string? sourceTableName, IReadOnlyList<string> headers, int row, int column)
        {
            Sheet = sheet;
            Number = number;
            Name = name;
            SourceSheetName = sourceSheet;
            SourceRange = sourceRange;
            SourceTableName = sourceTableName;
            Row = row;
            Column = column;
            _headers = new string[headers.Count];
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < _headers.Length; i++)
            {
                var h = headers[i];
                if (string.IsNullOrWhiteSpace(h))
                    throw new ArgumentException($"Header {i + 1} of pivot table '{name}' is empty. Every source column needs a header text; Excel uses it as the field name.", nameof(headers));
                // Excel refuses a cache with two fields of the same name ("The PivotTable field name is not valid").
                if (!seen.Add(h))
                    throw new ArgumentException($"Pivot table '{name}' has two source columns named '{h}'. Excel requires unique header texts.", nameof(headers));
                _headers[i] = h;
            }
        }

        internal Worksheet Sheet { get; }

        /// <summary>Workbook-wide pivot number: the part number of the pivot table, its cache definition and its records.</summary>
        internal int Number { get; }

        internal RangeReference SourceRange { get; }

        /// <summary>Table name when the source is an Excel table; the cache then follows the table as it grows.</summary>
        internal string? SourceTableName { get; }

        /// <summary>Pivot table name shown in Excel's PivotTable Analyze ribbon. Unique within the workbook.</summary>
        public string Name { get; }

        /// <summary>Name of the sheet holding the source data.</summary>
        public string SourceSheetName { get; }

        /// <summary>1-based row of the pivot's top-left cell (the first filter cell when there are filter fields).</summary>
        public int Row { get; }

        /// <summary>1-based column of the pivot's top-left cell.</summary>
        public int Column { get; }

        /// <summary>The source header texts, in column order; these are the names <c>Add*Field</c> accepts.</summary>
        public IReadOnlyList<string> Headers => _headers;

        /// <summary>Built-in pivot style, e.g. "PivotStyleMedium9" (default), "PivotStyleLight16", "PivotStyleDark2".</summary>
        public string StyleName { get; set; } = "PivotStyleMedium9";

        /// <summary>Show the Grand Total column at the right (the total of each row). Default true.</summary>
        public bool ShowRowGrandTotals { get; set; } = true;

        /// <summary>Show the Grand Total row at the bottom (the total of each column). Default true.</summary>
        public bool ShowColumnGrandTotals { get; set; } = true;

        /// <summary>
        /// Compact form (default): all row fields share one indented column, as Excel creates pivots. Set to
        /// <c>false</c> for outline form, where each row field gets its own column.
        /// </summary>
        public bool Compact { get; set; } = true;

        /// <summary>Shade alternate rows with the style's banding.</summary>
        public bool ShowRowStripes { get; set; }

        /// <summary>Shade alternate columns with the style's banding.</summary>
        public bool ShowColumnStripes { get; set; }

        /// <summary>Puts a source column on the row axis. Fields nest in the order they are added.</summary>
        /// <param name="sourceColumnHeader">Header text of the source column, as passed in <c>headers</c> or captured from the table.</param>
        public PivotTable AddRowField(string sourceColumnHeader)
        {
            Sheet.EnsureWritable();
            _rowFields.Add(Resolve(sourceColumnHeader, nameof(sourceColumnHeader)));
            return this;
        }

        /// <summary>Puts a source column on the column axis.</summary>
        public PivotTable AddColumnField(string sourceColumnHeader)
        {
            Sheet.EnsureWritable();
            _columnFields.Add(Resolve(sourceColumnHeader, nameof(sourceColumnHeader)));
            return this;
        }

        /// <summary>
        /// Adds a report filter (page field) above the pivot. Each filter takes one row; the pivot body starts
        /// below them with one blank row in between, as Excel lays it out.
        /// </summary>
        public PivotTable AddFilterField(string sourceColumnHeader)
        {
            Sheet.EnsureWritable();
            _pageFields.Add(Resolve(sourceColumnHeader, nameof(sourceColumnHeader)));
            return this;
        }

        /// <summary>
        /// Adds a value to the data area. With more than one data field Excel shows them side by side under a
        /// "Values" column header.
        /// </summary>
        /// <param name="sourceColumnHeader">Header text of the source column to aggregate.</param>
        /// <param name="function">Aggregation. Default sum.</param>
        /// <param name="displayName">
        /// Caption shown in the pivot, default "Sum of Amount" style. Must differ from every source header:
        /// Excel reports "PivotTable field name already exists" otherwise, so that is rejected here.
        /// </param>
        /// <param name="numberFormat">Excel number format for the values, e.g. "#,##0.00", or null for General.</param>
        public PivotTable AddDataField(string sourceColumnHeader, PivotFunction function = PivotFunction.Sum, string? displayName = null, string? numberFormat = null)
        {
            Sheet.EnsureWritable();
            var field = Resolve(sourceColumnHeader, nameof(sourceColumnHeader));
            var name = displayName ?? Caption(function) + " of " + _headers[field];
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Data field display name cannot be empty.", nameof(displayName));
            foreach (var h in _headers)
                if (string.Equals(h, name, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException($"Data field name '{name}' equals the source column '{h}'. Excel refuses a data field named like a source field; pick another displayName (for example '{name} total').", nameof(displayName));
            foreach (var d in _dataFields)
                if (string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException($"Pivot table '{Name}' already has a data field named '{name}'. Give the second one a displayName.", nameof(displayName));
            _dataFields.Add(new DataField(field, function, name, string.IsNullOrEmpty(numberFormat) ? null : numberFormat));
            return this;
        }

        /// <summary>Maps a header text to its 0-based field index. Exact match first, then case-insensitive.</summary>
        private int Resolve(string header, string parameterName)
        {
            if (header == null) throw new ArgumentNullException(parameterName);
            for (var i = 0; i < _headers.Length; i++) if (_headers[i] == header) return i;
            for (var i = 0; i < _headers.Length; i++) if (string.Equals(_headers[i], header, StringComparison.OrdinalIgnoreCase)) return i;
            throw new ArgumentException($"Pivot table '{Name}' has no source column '{header}'. Available headers: {string.Join(", ", _headers)}.", parameterName);
        }

        private static string Caption(PivotFunction function)
        {
            switch (function)
            {
                case PivotFunction.Sum: return "Sum";
                case PivotFunction.Count: return "Count";
                case PivotFunction.Average: return "Average";
                case PivotFunction.Max: return "Max";
                case PivotFunction.Min: return "Min";
                case PivotFunction.Product: return "Product";
                case PivotFunction.CountNumbers: return "Count";
                case PivotFunction.StdDev: return "StdDev";
                case PivotFunction.StdDevP: return "StdDevp";
                case PivotFunction.Var: return "Var";
                case PivotFunction.VarP: return "Varp";
                default: throw new ArgumentOutOfRangeException(nameof(function));
            }
        }

        /// <summary>The ST_DataConsolidateFunction token of a function.</summary>
        private static string SubtotalToken(PivotFunction function)
        {
            switch (function)
            {
                case PivotFunction.Sum: return "sum";
                case PivotFunction.Count: return "count";
                case PivotFunction.Average: return "average";
                case PivotFunction.Max: return "max";
                case PivotFunction.Min: return "min";
                case PivotFunction.Product: return "product";
                case PivotFunction.CountNumbers: return "countNums";
                case PivotFunction.StdDev: return "stdDev";
                case PivotFunction.StdDevP: return "stdDevp";
                case PivotFunction.Var: return "var";
                case PivotFunction.VarP: return "varp";
                default: throw new ArgumentOutOfRangeException(nameof(function));
            }
        }

        /// <summary>
        /// The cells the pivot body occupies before Excel refreshes it, as Excel's <c>location</c> element
        /// describes it: the body (without the filter rows), the row offset of the header rows and the first data
        /// row, and the column offset of the first data column. The source data is unknown here, so the body is
        /// sized for one item per axis; Excel resizes it on refresh. Internal so tests can pin the arithmetic.
        /// </summary>
        internal (RangeReference Ref, int FirstHeaderRow, int FirstDataRow, int FirstDataCol) ComputeLocation()
        {
            var rowCount = _rowFields.Count;
            var colCount = _columnFields.Count;
            var dataCount = _dataFields.Count;
            var valueColumns = Math.Max(1, dataCount);

            // Filters sit above the body, one per row, followed by one blank row (Excel's layout).
            var bodyRow = Row + (_pageFields.Count > 0 ? _pageFields.Count + 1 : 0);

            // Header rows above the first data row. Without column fields Excel uses a single header row
            // ("Row Labels | Sum of X [| Count of Y]"). Column fields add a caption row ("Sum of X | Column Labels"),
            // and several data fields under column fields add a "Values" row as well.
            var firstDataRow = colCount == 0 ? 1 : dataCount > 1 ? 3 : 2;

            // Compact form stacks every row field in one column; outline form gives each its own.
            var firstDataCol = rowCount == 0 ? 0 : Compact ? 1 : rowCount;

            var rows = firstDataRow + 1 + (rowCount > 0 && ShowColumnGrandTotals ? 1 : 0);
            var cols = firstDataCol + valueColumns + (colCount > 0 && ShowRowGrandTotals ? valueColumns : 0);

            var lastRow = Math.Min(CellReference.MaxRows, bodyRow + rows - 1);
            var lastCol = Math.Min(CellReference.MaxColumns, Column + cols - 1);
            return (new RangeReference(bodyRow, Column, lastRow, lastCol), 1, firstDataRow, firstDataCol);
        }

        /// <summary>Throws when the pivot cannot be written; called before any part is opened so a failure leaves the archive consistent.</summary>
        internal void Validate()
        {
            if (_rowFields.Count + _columnFields.Count + _pageFields.Count + _dataFields.Count == 0)
                throw new InvalidOperationException($"Pivot table '{Name}' has no fields. Call AddRowField / AddColumnField / AddFilterField / AddDataField before the sheet is sealed.");
            ComputeLocation(); // Throws when the filter rows push the body past the last row of the sheet.
        }

        /// <summary>
        /// Writes <c>xl/pivotCache/pivotCacheDefinitionN.xml</c>. The cache is empty and flagged
        /// <c>refreshOnLoad</c>; <c>saveData="0"</c> tells Excel the records part carries nothing worth keeping.
        /// </summary>
        internal void WriteCacheDefinition(XmlWriter w, string recordsRelId)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("pivotCacheDefinition", OpenXmlNamespaces.Main);
            w.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);
            w.WriteAttributeString("id", OpenXmlNamespaces.Relationships, recordsRelId);
            w.WriteAttributeString("saveData", "0");
            w.WriteAttributeString("refreshOnLoad", "1");
            w.WriteAttributeString("refreshedBy", "SheetStream");
            // Version stamps as Excel 2016 writes them. A refreshedVersion below the pivot's updatedVersion
            // makes Excel treat the cache as stale and some builds then refuse to refresh it.
            w.WriteAttributeString("createdVersion", "6");
            w.WriteAttributeString("refreshedVersion", "6");
            w.WriteAttributeString("minRefreshableVersion", "3");
            w.WriteAttributeString("recordCount", "0");

            // CT_PivotCacheDefinition order: cacheSource, cacheFields.
            w.WriteStartElement("cacheSource");
            w.WriteAttributeString("type", "worksheet");
            w.WriteStartElement("worksheetSource");
            if (SourceTableName != null)
            {
                // A table source keeps the pivot in step with the table's current size on every refresh.
                w.WriteAttributeString("name", SourceTableName);
            }
            else
            {
                w.WriteAttributeString("ref", SourceRange.ToString());
                w.WriteAttributeString("sheet", SourceSheetName);
            }
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("cacheFields");
            w.WriteAttributeString("count", _headers.Length.ToString(CultureInfo.InvariantCulture));
            foreach (var header in _headers)
            {
                w.WriteStartElement("cacheField");
                w.WriteAttributeString("name", XmlTextUtil.Sanitize(header));
                w.WriteAttributeString("numFmtId", "0");
                // No items: Excel rebuilds the shared item list from the source on load.
                w.WriteStartElement("sharedItems");
                w.WriteEndElement();
                w.WriteEndElement();
            }
            w.WriteEndElement();

            w.WriteEndElement();
            w.WriteEndDocument();
        }

        /// <summary>
        /// Writes <c>xl/pivotCache/pivotCacheRecordsN.xml</c> with zero records. Excel expects the part to exist
        /// even when the definition says <c>saveData="0"</c>; without it some versions repair the file and drop the pivot.
        /// </summary>
        internal static void WriteCacheRecords(XmlWriter w)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("pivotCacheRecords", OpenXmlNamespaces.Main);
            w.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);
            w.WriteAttributeString("count", "0");
            w.WriteEndElement();
            w.WriteEndDocument();
        }

        /// <summary>
        /// Writes <c>xl/pivotTables/pivotTableN.xml</c>. CT_pivotTableDefinition order: location, pivotFields,
        /// rowFields, colFields, pageFields, dataFields, pivotTableStyleInfo (rowItems/colItems are omitted; Excel
        /// derives them when it refreshes the empty cache).
        /// </summary>
        internal void WriteDefinition(XmlWriter w, int cacheId)
        {
            var (location, firstHeaderRow, firstDataRow, firstDataCol) = ComputeLocation();

            w.WriteStartDocument(true);
            w.WriteStartElement("pivotTableDefinition", OpenXmlNamespaces.Main);
            w.WriteAttributeString("name", XmlTextUtil.Sanitize(Name));
            w.WriteAttributeString("cacheId", cacheId.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("applyNumberFormats", "0");
            w.WriteAttributeString("applyBorderFormats", "0");
            w.WriteAttributeString("applyFontFormats", "0");
            w.WriteAttributeString("applyPatternFormats", "0");
            w.WriteAttributeString("applyAlignmentFormats", "0");
            w.WriteAttributeString("applyWidthHeightFormats", "1");
            w.WriteAttributeString("dataCaption", "Values");
            w.WriteAttributeString("updatedVersion", "6");
            w.WriteAttributeString("minRefreshableVersion", "3");
            w.WriteAttributeString("useAutoFormatting", "1");
            if (!ShowRowGrandTotals) w.WriteAttributeString("rowGrandTotals", "0");
            if (!ShowColumnGrandTotals) w.WriteAttributeString("colGrandTotals", "0");
            w.WriteAttributeString("itemPrintTitles", "1");
            w.WriteAttributeString("createdVersion", "6");
            w.WriteAttributeString("indent", "0");
            if (!Compact) w.WriteAttributeString("compact", "0");
            w.WriteAttributeString("outline", "1");
            w.WriteAttributeString("outlineData", "1");
            if (!Compact) w.WriteAttributeString("compactData", "0");
            w.WriteAttributeString("multipleFieldFilters", "0");

            w.WriteStartElement("location");
            w.WriteAttributeString("ref", location.ToString());
            w.WriteAttributeString("firstHeaderRow", firstHeaderRow.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("firstDataRow", firstDataRow.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("firstDataCol", firstDataCol.ToString(CultureInfo.InvariantCulture));
            if (_pageFields.Count > 0)
            {
                w.WriteAttributeString("rowPageCount", _pageFields.Count.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("colPageCount", "1");
            }
            w.WriteEndElement();

            // One pivotField per cache field, in cache order; the axis and dataField flags say where it is used.
            w.WriteStartElement("pivotFields");
            w.WriteAttributeString("count", _headers.Length.ToString(CultureInfo.InvariantCulture));
            for (var i = 0; i < _headers.Length; i++)
            {
                var axis = _rowFields.Contains(i) ? "axisRow" : _columnFields.Contains(i) ? "axisCol" : _pageFields.Contains(i) ? "axisPage" : null;
                var isData = false;
                foreach (var d in _dataFields) isData |= d.Field == i;

                w.WriteStartElement("pivotField");
                if (axis != null) w.WriteAttributeString("axis", axis);
                if (isData) w.WriteAttributeString("dataField", "1");
                if (!Compact) w.WriteAttributeString("compact", "0");
                w.WriteAttributeString("showAll", "0");
                if (axis != null)
                {
                    // An axis field needs an item list; with an empty cache the only item is the "default" (subtotal) placeholder.
                    w.WriteStartElement("items");
                    w.WriteAttributeString("count", "1");
                    w.WriteStartElement("item");
                    w.WriteAttributeString("t", "default");
                    w.WriteEndElement();
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }
            w.WriteEndElement();

            WriteAxis(w, "rowFields", _rowFields, false);
            // Several data fields go on the column axis as the pseudo field -2 ("Values"), Excel's default placement.
            WriteAxis(w, "colFields", _columnFields, _dataFields.Count > 1);

            if (_pageFields.Count > 0)
            {
                w.WriteStartElement("pageFields");
                w.WriteAttributeString("count", _pageFields.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var f in _pageFields)
                {
                    w.WriteStartElement("pageField");
                    w.WriteAttributeString("fld", f.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("hier", "-1"); // Not an OLAP hierarchy.
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }

            if (_dataFields.Count > 0)
            {
                w.WriteStartElement("dataFields");
                w.WriteAttributeString("count", _dataFields.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var d in _dataFields)
                {
                    w.WriteStartElement("dataField");
                    w.WriteAttributeString("name", XmlTextUtil.Sanitize(d.Name));
                    w.WriteAttributeString("fld", d.Field.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("subtotal", SubtotalToken(d.Function));
                    w.WriteAttributeString("baseField", "0");
                    w.WriteAttributeString("baseItem", "0");
                    if (d.NumberFormat != null)
                        w.WriteAttributeString("numFmtId", Sheet.Workbook.Styles.RegisterNumberFormat(d.NumberFormat).ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }

            w.WriteStartElement("pivotTableStyleInfo");
            w.WriteAttributeString("name", StyleName);
            w.WriteAttributeString("showRowHeaders", "1");
            w.WriteAttributeString("showColHeaders", "1");
            w.WriteAttributeString("showRowStripes", ShowRowStripes ? "1" : "0");
            w.WriteAttributeString("showColStripes", ShowColumnStripes ? "1" : "0");
            w.WriteAttributeString("showLastColumn", "1");
            w.WriteEndElement();

            w.WriteEndElement();
            w.WriteEndDocument();
        }

        private static void WriteAxis(XmlWriter w, string element, List<int> fields, bool appendValuesField)
        {
            var count = fields.Count + (appendValuesField ? 1 : 0);
            if (count == 0) return;
            w.WriteStartElement(element);
            w.WriteAttributeString("count", count.ToString(CultureInfo.InvariantCulture));
            foreach (var f in fields)
            {
                w.WriteStartElement("field");
                w.WriteAttributeString("x", f.ToString(CultureInfo.InvariantCulture));
                w.WriteEndElement();
            }
            if (appendValuesField)
            {
                w.WriteStartElement("field");
                w.WriteAttributeString("x", "-2");
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }
    }

    /// <summary>
    /// Writes the three parts of one pivot table and their relationships: the cache definition (related to
    /// the records part), the empty records part, and the pivot table definition (related to the cache). The
    /// sheet's own <c>.rels</c> entry and the workbook's <c>pivotCaches</c> entry are added by the callers.
    /// </summary>
    internal static class PivotTableWriter
    {
        public static void Write(ExcelPackage package, Workbook workbook, PivotTable pivot)
        {
            pivot.Validate();
            var n = pivot.Number.ToString(CultureInfo.InvariantCulture);

            var cacheRels = new RelationshipSet();
            var recordsRelId = cacheRels.Add(OpenXmlNamespaces.RelPivotCacheRecords, "pivotCacheRecords" + n + ".xml");
            var cachePart = "xl/pivotCache/pivotCacheDefinition" + n + ".xml";
            using (var w = package.OpenPart(cachePart)) pivot.WriteCacheDefinition(w, recordsRelId);
            package.RegisterPart(cachePart, OpenXmlNamespaces.CtPivotCacheDefinition);
            using (var w = package.OpenPart("xl/pivotCache/_rels/pivotCacheDefinition" + n + ".xml.rels")) cacheRels.WriteTo(w);

            var recordsPart = "xl/pivotCache/pivotCacheRecords" + n + ".xml";
            using (var w = package.OpenPart(recordsPart)) PivotTable.WriteCacheRecords(w);
            package.RegisterPart(recordsPart, OpenXmlNamespaces.CtPivotCacheRecords);

            // The cache id is the part number; workbook.xml lists the same number against the cache relationship.
            workbook.PivotCaches.Add(pivot.Number);
            var tableRels = new RelationshipSet();
            tableRels.Add(OpenXmlNamespaces.RelPivotCacheDefinition, "../pivotCache/pivotCacheDefinition" + n + ".xml");
            var tablePart = "xl/pivotTables/pivotTable" + n + ".xml";
            using (var w = package.OpenPart(tablePart)) pivot.WriteDefinition(w, pivot.Number);
            package.RegisterPart(tablePart, OpenXmlNamespaces.CtPivotTable);
            using (var w = package.OpenPart("xl/pivotTables/_rels/pivotTable" + n + ".xml.rels")) tableRels.WriteTo(w);
        }
    }
}
