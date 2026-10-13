using System;
using SheetStream.Styles;

namespace SheetStream
{
    /// <summary>
    /// Width, visibility and default style of one worksheet column. Obtain through
    /// <see cref="Worksheet.Column(int)"/>; settings must be made before the sheet's first row is written.
    /// </summary>
    public sealed class ColumnSettings
    {
        internal ColumnSettings(int index)
        {
            Index = index;
        }

        /// <summary>1-based column index.</summary>
        public int Index { get; }

        /// <summary>Width in Excel character units, or <c>null</c> for Excel's default (about 8.43).</summary>
        public double? Width { get; private set; }

        /// <summary>True when the column is hidden.</summary>
        public bool IsHidden { get; private set; }

        /// <summary>Default style for cells of this column that are not written explicitly.</summary>
        public CellStyle? Style { get; private set; }

        /// <summary>
        /// Sets the column width in Excel character units (the number shown when dragging a column edge
        /// in Excel). Excel accepts 0 to 255.
        /// </summary>
        public ColumnSettings SetWidth(double width)
        {
            if (width < 0 || width > 255) throw new ArgumentOutOfRangeException(nameof(width), width, "Column width must be between 0 and 255 characters.");
            Width = width;
            return this;
        }

        /// <summary>Hides the column.</summary>
        public ColumnSettings Hide()
        {
            IsHidden = true;
            return this;
        }

        /// <summary>Applies a default style to the whole column (cells written with their own style keep theirs).</summary>
        public ColumnSettings SetStyle(CellStyle style)
        {
            Style = style ?? throw new ArgumentNullException(nameof(style));
            return this;
        }

        /// <summary>True when the column needs a <c>&lt;col&gt;</c> element at all.</summary>
        internal bool HasSettings => Width.HasValue || IsHidden || (Style != null && !Style.IsDefault);

        /// <summary>True when two columns can share one <c>&lt;col min max&gt;</c> element.</summary>
        internal bool HasSameSettingsAs(ColumnSettings other)
            => Width == other.Width && IsHidden == other.IsHidden && (Style?.Id ?? 0) == (other.Style?.Id ?? 0);
    }

    /// <summary>
    /// Applies settings to a contiguous range of columns. Obtain through <see cref="Worksheet.Columns(int, int)"/>.
    /// Each call fans out to the individual <see cref="ColumnSettings"/>; the writer later coalesces
    /// identical neighbours back into ranges.
    /// </summary>
    public sealed class ColumnRangeSettings
    {
        private readonly Worksheet _sheet;

        internal ColumnRangeSettings(Worksheet sheet, int from, int to)
        {
            _sheet = sheet;
            From = from;
            To = to;
        }

        /// <summary>First column of the range (1-based, inclusive).</summary>
        public int From { get; }

        /// <summary>Last column of the range (1-based, inclusive).</summary>
        public int To { get; }

        /// <summary>Sets the width of every column in the range. See <see cref="ColumnSettings.SetWidth"/>.</summary>
        public ColumnRangeSettings SetWidth(double width)
        {
            if (width < 0 || width > 255) throw new ArgumentOutOfRangeException(nameof(width), width, "Column width must be between 0 and 255 characters.");
            for (var i = From; i <= To; i++) _sheet.GetOrCreateColumn(i).SetWidth(width);
            return this;
        }

        /// <summary>Hides every column in the range.</summary>
        public ColumnRangeSettings Hide()
        {
            for (var i = From; i <= To; i++) _sheet.GetOrCreateColumn(i).Hide();
            return this;
        }

        /// <summary>Applies a default style to every column in the range.</summary>
        public ColumnRangeSettings SetStyle(CellStyle style)
        {
            if (style == null) throw new ArgumentNullException(nameof(style));
            for (var i = From; i <= To; i++) _sheet.GetOrCreateColumn(i).SetStyle(style);
            return this;
        }
    }
}
