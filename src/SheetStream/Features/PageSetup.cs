using System;
using System.Globalization;
using System.Text;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>Printed page orientation.</summary>
    public enum PageOrientation
    {
        /// <summary>Taller than wide.</summary>
        Portrait,
        /// <summary>Wider than tall.</summary>
        Landscape,
    }

    /// <summary>Common paper sizes; values are the ECMA-376 paper size codes.</summary>
    public enum PaperSize
    {
        /// <summary>US Letter, 8.5 x 11 in.</summary>
        Letter = 1,
        /// <summary>US Legal, 8.5 x 14 in.</summary>
        Legal = 5,
        /// <summary>A3, 297 x 420 mm.</summary>
        A3 = 8,
        /// <summary>A4, 210 x 297 mm.</summary>
        A4 = 9,
        /// <summary>A5, 148 x 210 mm.</summary>
        A5 = 11,
        /// <summary>US Tabloid, 11 x 17 in.</summary>
        Tabloid = 3,
    }

    /// <summary>
    /// Print settings of a worksheet. Everything here can be set any time before the sheet is sealed, except
    /// <see cref="FitToWidth"/> and <see cref="FitToHeight"/>, which Excel stores in the sheet header and so
    /// must be set before the first row is written.
    /// </summary>
    public sealed class PageSetup
    {
        private readonly Worksheet _sheet;
        private int? _fitToWidth;
        private int? _fitToHeight;

        internal PageSetup(Worksheet sheet)
        {
            _sheet = sheet;
        }

        /// <summary>Page orientation; null keeps the printer default.</summary>
        public PageOrientation? Orientation { get; set; }

        /// <summary>Paper size; null keeps the printer default.</summary>
        public PaperSize? PaperSize { get; set; }

        /// <summary>Print scale in percent (10 to 400). Ignored when fit-to-page is used.</summary>
        public int? Scale { get; set; }

        /// <summary>
        /// Fit the printout to this many pages wide (0 = automatic). Must be set before the first row,
        /// because the "fit to page" switch lives in the sheet header.
        /// </summary>
        public int? FitToWidth
        {
            get => _fitToWidth;
            set { _sheet.EnsureHeaderNotWritten("Fit-to-page printing"); _fitToWidth = value; }
        }

        /// <summary>Fit the printout to this many pages tall (0 = automatic). Must be set before the first row.</summary>
        public int? FitToHeight
        {
            get => _fitToHeight;
            set { _sheet.EnsureHeaderNotWritten("Fit-to-page printing"); _fitToHeight = value; }
        }

        /// <summary>Print the cell gridlines.</summary>
        public bool PrintGridlines { get; set; }

        /// <summary>Center the printout horizontally on the page.</summary>
        public bool CenterHorizontally { get; set; }

        /// <summary>Center the printout vertically on the page.</summary>
        public bool CenterVertically { get; set; }

        /// <summary>Margins in inches: left, right, top, bottom, header, footer. Null keeps Excel's defaults.</summary>
        public (double Left, double Right, double Top, double Bottom, double Header, double Footer)? Margins { get; set; }

        /// <summary>
        /// Page header. Use Excel's codes: <c>&amp;L</c>, <c>&amp;C</c>, <c>&amp;R</c> for sections,
        /// <c>&amp;P</c> page, <c>&amp;N</c> pages, <c>&amp;D</c> date, <c>&amp;A</c> sheet name. A text
        /// without section codes is centered.
        /// </summary>
        public string? Header { get; set; }

        /// <summary>Page footer, same codes as <see cref="Header"/>. Example: <c>"Page &amp;P of &amp;N"</c>.</summary>
        public string? Footer { get; set; }

        /// <summary>Rows repeated at the top of every printed page, e.g. (1, 1) for a header row.</summary>
        public (int From, int To)? PrintTitleRows { get; set; }

        /// <summary>Columns repeated at the left of every printed page.</summary>
        public (int From, int To)? PrintTitleColumns { get; set; }

        /// <summary>Restricts printing to a range such as <c>"A1:F50"</c>.</summary>
        public string? PrintArea { get; set; }

        internal bool FitToPage => _fitToWidth.HasValue || _fitToHeight.HasValue;

        /// <summary>Value of the <c>_xlnm.Print_Titles</c> defined name, or null.</summary>
        internal string? PrintTitlesFormula
        {
            get
            {
                if (PrintTitleRows == null && PrintTitleColumns == null) return null;
                var sheet = RangeReference.QuoteSheetName(_sheet.Name);
                var sb = new StringBuilder();
                if (PrintTitleColumns is (int cf, int ct))
                    sb.Append(sheet).Append("!$").Append(CellReference.ColumnLetters(cf)).Append(":$").Append(CellReference.ColumnLetters(ct));
                if (PrintTitleRows is (int rf, int rt))
                {
                    if (sb.Length > 0) sb.Append(',');
                    sb.Append(sheet).Append("!$").Append(rf.ToString(CultureInfo.InvariantCulture)).Append(":$").Append(rt.ToString(CultureInfo.InvariantCulture));
                }
                return sb.ToString();
            }
        }

        /// <summary>Value of the <c>_xlnm.Print_Area</c> defined name, or null.</summary>
        internal string? PrintAreaFormula => PrintArea == null ? null : RangeReference.QualifyAbsolute(PrintArea, _sheet.Name);

        /// <summary>Writes printOptions, pageMargins, pageSetup and headerFooter, in that schema order.</summary>
        internal void WriteTo(XmlWriter w)
        {
            if (PrintGridlines || CenterHorizontally || CenterVertically)
            {
                w.WriteStartElement("printOptions");
                if (CenterHorizontally) w.WriteAttributeString("horizontalCentered", "1");
                if (CenterVertically) w.WriteAttributeString("verticalCentered", "1");
                if (PrintGridlines) w.WriteAttributeString("gridLines", "1");
                w.WriteEndElement();
            }

            // pageMargins is required whenever pageSetup is present; Excel's defaults are used if unset.
            var hasSetup = Orientation.HasValue || PaperSize.HasValue || Scale.HasValue || FitToPage;
            if (Margins.HasValue || hasSetup || Header != null || Footer != null)
            {
                var m = Margins ?? (0.7, 0.7, 0.75, 0.75, 0.3, 0.3);
                w.WriteStartElement("pageMargins");
                w.WriteAttributeString("left", m.Left.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("right", m.Right.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("top", m.Top.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("bottom", m.Bottom.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("header", m.Header.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("footer", m.Footer.ToString(CultureInfo.InvariantCulture));
                w.WriteEndElement();
            }

            if (hasSetup)
            {
                w.WriteStartElement("pageSetup");
                if (PaperSize.HasValue) w.WriteAttributeString("paperSize", ((int)PaperSize.Value).ToString(CultureInfo.InvariantCulture));
                if (Scale.HasValue)
                {
                    if (Scale < 10 || Scale > 400) throw new ArgumentOutOfRangeException(nameof(Scale), Scale, "Print scale must be between 10 and 400 percent.");
                    w.WriteAttributeString("scale", Scale.Value.ToString(CultureInfo.InvariantCulture));
                }
                if (FitToPage)
                {
                    w.WriteAttributeString("fitToWidth", (_fitToWidth ?? 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("fitToHeight", (_fitToHeight ?? 0).ToString(CultureInfo.InvariantCulture));
                }
                if (Orientation.HasValue) w.WriteAttributeString("orientation", Orientation == PageOrientation.Landscape ? "landscape" : "portrait");
                w.WriteEndElement();
            }

            if (Header != null || Footer != null)
            {
                w.WriteStartElement("headerFooter");
                if (Header != null) w.WriteElementString("oddHeader", Section(Header));
                if (Footer != null) w.WriteElementString("oddFooter", Section(Footer));
                w.WriteEndElement();
            }
        }

        /// <summary>Text without an explicit section code is centred, as in Excel's header dialog.</summary>
        private static string Section(string text)
        {
            var t = XmlTextUtil.Sanitize(text);
            return t.Contains("&L") || t.Contains("&C") || t.Contains("&R") ? t : "&C" + t;
        }
    }
}
