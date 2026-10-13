using System;
using System.Collections.Generic;
using SheetStream.Styles;

namespace SheetStream
{
    /// <summary>
    /// Options for <see cref="Worksheet.LoadFromCollection{T}(IEnumerable{T}, LoadOptions?)"/>,
    /// <see cref="Worksheet.LoadFromDataTable"/> and <see cref="Worksheet.LoadFromDataReader"/>.
    /// </summary>
    public sealed class LoadOptions
    {
        /// <summary>Write a header row with the column names. Default true.</summary>
        public bool PrintHeaders { get; set; } = true;

        /// <summary>Style of the header cells, or null.</summary>
        public CellStyle? HeaderStyle { get; set; }

        /// <summary>First column to write to (1-based). Default 1.</summary>
        public int StartColumn { get; set; } = 1;

        /// <summary>
        /// Creates an Excel table with this name over the loaded range (header included). Requires
        /// <see cref="PrintHeaders"/>.
        /// </summary>
        public string? TableName { get; set; }

        /// <summary>Table style when <see cref="TableName"/> is set. Default "TableStyleMedium2".</summary>
        public string TableStyle { get; set; } = "TableStyleMedium2";

        /// <summary>Adds AutoFilter drop-downs to the header row (ignored when a table is created, which has its own).</summary>
        public bool AutoFilter { get; set; }

        /// <summary>Per-column styles keyed by column name (property name, header text or data column name).</summary>
        public IDictionary<string, CellStyle> ColumnStyles { get; } = new Dictionary<string, CellStyle>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Controls how a property is exported by <see cref="Worksheet.LoadFromCollection{T}(IEnumerable{T}, LoadOptions?)"/>
    /// and matched by <c>ExcelReader.Read&lt;T&gt;</c>.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public sealed class SheetColumnAttribute : Attribute
    {
        /// <summary>Creates the attribute without a header override.</summary>
        public SheetColumnAttribute()
        {
        }

        /// <summary>Creates the attribute with a header text.</summary>
        public SheetColumnAttribute(string name)
        {
            Name = name;
        }

        /// <summary>Header text; defaults to the property name (or its [DisplayName]).</summary>
        public string? Name { get; set; }

        /// <summary>Sort key; columns without an order keep declaration order after the ordered ones.</summary>
        public int Order { get; set; } = int.MaxValue;

        /// <summary>Excel number format for the column, e.g. "#,##0.00".</summary>
        public string? NumberFormat { get; set; }

        /// <summary>Skip this member.</summary>
        public bool Ignore { get; set; }
    }

    /// <summary>The cells a load call wrote, for follow-up formatting such as charts or totals.</summary>
    public readonly struct LoadResult
    {
        internal LoadResult(int firstRow, int lastRow, int firstColumn, int lastColumn, int dataRows)
        {
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
            DataRows = dataRows;
        }

        /// <summary>First row written (the header row when headers are printed).</summary>
        public int FirstRow { get; }

        /// <summary>Last row written.</summary>
        public int LastRow { get; }

        /// <summary>First column written.</summary>
        public int FirstColumn { get; }

        /// <summary>Last column written.</summary>
        public int LastColumn { get; }

        /// <summary>Number of data rows (excluding the header).</summary>
        public int DataRows { get; }
    }
}
