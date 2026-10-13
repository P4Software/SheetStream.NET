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
    }
}
