using System;
using SheetStream.Internal;

namespace SheetStream.Styles
{
    /// <summary>
    /// Describes the appearance of a cell. Configure an instance (usually through
    /// <see cref="StyleSheet.Create(Action{StyleOptions})"/>) and the style sheet turns it into a
    /// registered, de-duplicated <see cref="CellStyle"/>.
    /// </summary>
    public sealed class StyleOptions
    {
        /// <summary>Font settings (bold, italic, colour, size, name).</summary>
        public FontOptions Font { get; } = new FontOptions();

        /// <summary>Background fill settings.</summary>
        public FillOptions Fill { get; } = new FillOptions();

        /// <summary>Border settings for the four edges.</summary>
        public BorderOptions Border { get; } = new BorderOptions();

        /// <summary>
        /// Excel number format code, e.g. <c>"#,##0.00"</c>, <c>"yyyy-mm-dd"</c> or <c>"0%"</c>.
        /// <c>null</c> means "General". See <see cref="NumberFormats"/> for common codes.
        /// </summary>
        public string? NumberFormat { get; set; }

        /// <summary>Horizontal alignment. Defaults to <see cref="HorizontalAlignment.General"/>.</summary>
        public HorizontalAlignment HorizontalAlignment { get; set; } = HorizontalAlignment.General;

        /// <summary>Vertical alignment. Defaults to <see cref="VerticalAlignment.Bottom"/>.</summary>
        public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Bottom;

        /// <summary>Wraps long text onto multiple lines inside the cell.</summary>
        public bool WrapText { get; set; }

        /// <summary>Shrinks the font so the text fits the column width.</summary>
        public bool ShrinkToFit { get; set; }

        /// <summary>Indent level (each level is about three characters). 0 to 250.</summary>
        public int Indent { get; set; }

        /// <summary>
        /// Text rotation in degrees: 0 to 90 rotates counter-clockwise, 91 to 180 means 1 to 90 degrees
        /// clockwise (Excel's encoding), and 255 stacks the letters vertically.
        /// </summary>
        public int TextRotation { get; set; }

        /// <summary>Cell protection flags; they only take effect once the sheet is protected.</summary>
        public ProtectionOptions Protection { get; } = new ProtectionOptions();
    }

    /// <summary>Cell protection flags of a <see cref="StyleOptions"/>.</summary>
    public sealed class ProtectionOptions
    {
        /// <summary>
        /// True (Excel's default) keeps the cell read-only when the sheet is protected. Set to false for
        /// input cells users may still edit on a protected sheet.
        /// </summary>
        public bool Locked { get; set; } = true;

        /// <summary>Hides the cell's formula from the formula bar when the sheet is protected.</summary>
        public bool HideFormula { get; set; }
    }

    /// <summary>Font settings of a <see cref="StyleOptions"/>.</summary>
    public sealed class FontOptions
    {
        private string? _color;

        /// <summary>Bold text.</summary>
        public bool Bold { get; set; }

        /// <summary>Italic text.</summary>
        public bool Italic { get; set; }

        /// <summary>Single underline.</summary>
        public bool Underline { get; set; }

        /// <summary>Strikethrough text.</summary>
        public bool Strikethrough { get; set; }

        /// <summary>Font size in points. Defaults to 11 (Excel's default for Calibri).</summary>
        public double Size { get; set; } = 11;

        /// <summary>Font name. Defaults to Calibri, which is also what the workbook's "Normal" style uses.</summary>
        public string Name { get; set; } = "Calibri";

        /// <summary>
        /// Text colour as hex: <c>"#RRGGBB"</c>, <c>"RRGGBB"</c> or <c>"AARRGGBB"</c>.
        /// <c>null</c> keeps Excel's automatic (theme) colour. Validated on assignment.
        /// </summary>
        public string? Color
        {
            get => _color;
            set => _color = value == null ? null : XmlTextUtil.NormalizeColor(value, nameof(Color));
        }
    }

    /// <summary>Fill settings of a <see cref="StyleOptions"/>. Only solid fills are supported.</summary>
    public sealed class FillOptions
    {
        private string? _backgroundColor;

        /// <summary>
        /// Solid background colour as hex: <c>"#RRGGBB"</c>, <c>"RRGGBB"</c> or <c>"AARRGGBB"</c>.
        /// <c>null</c> means no fill. Validated on assignment.
        /// </summary>
        public string? BackgroundColor
        {
            get => _backgroundColor;
            set => _backgroundColor = value == null ? null : XmlTextUtil.NormalizeColor(value, nameof(BackgroundColor));
        }
    }

    /// <summary>Border settings of a <see cref="StyleOptions"/>.</summary>
    public sealed class BorderOptions
    {
        private string? _color;

        /// <summary>Left edge line style.</summary>
        public BorderStyle Left { get; set; }

        /// <summary>Right edge line style.</summary>
        public BorderStyle Right { get; set; }

        /// <summary>Top edge line style.</summary>
        public BorderStyle Top { get; set; }

        /// <summary>Bottom edge line style.</summary>
        public BorderStyle Bottom { get; set; }

        /// <summary>
        /// Convenience setter that applies one line style to all four edges. Reading it returns the
        /// style of the left edge, so a mixed border reads back as whatever its left edge is.
        /// </summary>
        public BorderStyle Style
        {
            get => Left;
            set { Left = value; Right = value; Top = value; Bottom = value; }
        }

        /// <summary>
        /// Line colour for all edges as hex. <c>null</c> means Excel's automatic colour (black).
        /// Validated on assignment.
        /// </summary>
        public string? Color
        {
            get => _color;
            set => _color = value == null ? null : XmlTextUtil.NormalizeColor(value, nameof(Color));
        }

        /// <summary>True when at least one edge has a visible line.</summary>
        public bool HasAnyLine => Left != BorderStyle.None || Right != BorderStyle.None || Top != BorderStyle.None || Bottom != BorderStyle.None;
    }

    /// <summary>Frequently used Excel number format codes.</summary>
    public static class NumberFormats
    {
        /// <summary>Excel's default format.</summary>
        public const string General = "General";
        /// <summary>Whole number with thousands separator.</summary>
        public const string Integer = "#,##0";
        /// <summary>Two decimals with thousands separator.</summary>
        public const string Decimal2 = "#,##0.00";
        /// <summary>Percentage with two decimals.</summary>
        public const string Percent = "0.00%";
        /// <summary>ISO-style date.</summary>
        public const string Date = "yyyy-mm-dd";
        /// <summary>ISO-style date and time.</summary>
        public const string DateTime = "yyyy-mm-dd hh:mm:ss";
        /// <summary>Elapsed time that can exceed 24 hours.</summary>
        public const string Duration = "[h]:mm:ss";
        /// <summary>Forces text display; numbers are shown as typed.</summary>
        public const string Text = "@";
    }

    /// <summary>Helpers for building colour strings accepted by the style options.</summary>
    public static class Colors
    {
        /// <summary>Builds an opaque <c>#RRGGBB</c> colour from byte components.</summary>
        public static string Rgb(byte red, byte green, byte blue)
            => "#" + red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");

        /// <summary>Pure white.</summary>
        public const string White = "#FFFFFF";
        /// <summary>Pure black.</summary>
        public const string Black = "#000000";
        /// <summary>Excel's default header blue (the "Accent 1" of the Office theme).</summary>
        public const string ExcelBlue = "#4F81BD";
        /// <summary>Light grey used for zebra striping.</summary>
        public const string LightGray = "#F2F2F2";
    }
}
