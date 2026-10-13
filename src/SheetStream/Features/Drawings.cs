using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>Chart kinds supported by <see cref="Worksheet.AddChart"/>.</summary>
    public enum ChartType
    {
        /// <summary>Vertical bars side by side.</summary>
        Column,
        /// <summary>Vertical bars stacked.</summary>
        ColumnStacked,
        /// <summary>Vertical bars stacked to 100%.</summary>
        ColumnStacked100,
        /// <summary>Horizontal bars side by side.</summary>
        Bar,
        /// <summary>Horizontal bars stacked.</summary>
        BarStacked,
        /// <summary>Horizontal bars stacked to 100%.</summary>
        BarStacked100,
        /// <summary>Lines without markers.</summary>
        Line,
        /// <summary>Lines with markers.</summary>
        LineMarkers,
        /// <summary>Stacked lines.</summary>
        LineStacked,
        /// <summary>Filled area.</summary>
        Area,
        /// <summary>Stacked filled areas.</summary>
        AreaStacked,
        /// <summary>Pie (first series only).</summary>
        Pie,
        /// <summary>Doughnut.</summary>
        Doughnut,
        /// <summary>XY scatter, markers only. Categories are the X values and must be numeric.</summary>
        Scatter,
        /// <summary>XY scatter with lines and markers.</summary>
        ScatterLines,
    }

    /// <summary>Where the chart legend sits.</summary>
    public enum LegendPosition
    {
        /// <summary>No legend.</summary>
        None,
        /// <summary>Right of the plot (Excel default).</summary>
        Right,
        /// <summary>Left of the plot.</summary>
        Left,
        /// <summary>Above the plot.</summary>
        Top,
        /// <summary>Below the plot.</summary>
        Bottom,
    }

    /// <summary>Labels drawn on data points.</summary>
    public enum ChartDataLabels
    {
        /// <summary>No labels.</summary>
        None,
        /// <summary>The value of each point.</summary>
        Value,
        /// <summary>The share of the total (pie and doughnut).</summary>
        Percent,
        /// <summary>Category name and value.</summary>
        CategoryAndValue,
    }

    /// <summary>One data series of a chart.</summary>
    public sealed class ChartSeries
    {
        internal ChartSeries(string values, string? categories, string? name, string? color)
        {
            Values = values;
            Categories = categories;
            Name = name;
            Color = color;
        }

        /// <summary>Absolute, sheet-qualified reference of the values.</summary>
        public string Values { get; }

        /// <summary>Absolute reference of the category (or X) values, or null.</summary>
        public string? Categories { get; }

        /// <summary>Series name: literal text, or a cell reference when it was given starting with "=".</summary>
        public string? Name { get; }

        /// <summary>Series colour as ARGB hex, or null for the theme colour.</summary>
        public string? Color { get; }
    }

    /// <summary>
    /// A chart placed on a worksheet, referencing data already on a sheet. Create with
    /// <see cref="Worksheet.AddChart"/>, then add series. The chart reads the cells when Excel opens the file.
    /// </summary>
    /// <example>
    /// <code>
    /// var chart = sheet.AddChart(ChartType.Column, "E2:L20");
    /// chart.Title = "Revenue by SKU";
    /// chart.AddSeries(values: "C2:C11", categories: "B2:B11", name: "=C1");
    /// </code>
    /// </example>
    public sealed class ExcelChart
    {
        private readonly List<ChartSeries> _series = new List<ChartSeries>();
        private readonly string _sheetName;

        internal ExcelChart(ChartType type, RangeReference anchor, string sheetName, int chartNumber)
        {
            Type = type;
            Anchor = anchor;
            _sheetName = sheetName;
            Number = chartNumber;
        }

        internal RangeReference Anchor { get; }
        internal int Number { get; }

        /// <summary>The chart kind.</summary>
        public ChartType Type { get; }

        /// <summary>Chart title, or null for none.</summary>
        public string? Title { get; set; }

        /// <summary>Category (horizontal) axis title.</summary>
        public string? XAxisTitle { get; set; }

        /// <summary>Value (vertical) axis title.</summary>
        public string? YAxisTitle { get; set; }

        /// <summary>Legend placement. Default right.</summary>
        public LegendPosition Legend { get; set; } = LegendPosition.Right;

        /// <summary>Data labels on points. Default none.</summary>
        public ChartDataLabels DataLabels { get; set; }

        /// <summary>Number format of the value axis, e.g. "$#,##0". Null follows the source cells.</summary>
        public string? ValueAxisNumberFormat { get; set; }

        /// <summary>The series added so far.</summary>
        public IReadOnlyList<ChartSeries> Series => _series;

        /// <summary>
        /// Adds a series. Ranges are A1 ranges on this sheet ("C2:C11") or sheet-qualified ("'Data'!C2:C11").
        /// <paramref name="name"/> is literal text, or a cell reference when it starts with "=" (e.g. "=C1").
        /// </summary>
        public ChartSeries AddSeries(string values, string? categories = null, string? name = null, string? color = null)
        {
            var v = RangeReference.QualifyAbsolute(values, _sheetName);
            var c = categories == null ? null : RangeReference.QualifyAbsolute(categories, _sheetName);
            string? n = null;
            if (name != null)
                n = name.StartsWith("=", StringComparison.Ordinal) ? "=" + RangeReference.QualifyAbsolute(name, _sheetName) : name;
            var col = color == null ? null : XmlTextUtil.NormalizeColor(color, nameof(color)).Substring(2);
            var s = new ChartSeries(v, c, n, col);
            _series.Add(s);
            return s;
        }

        private bool IsPieLike => Type == ChartType.Pie || Type == ChartType.Doughnut;
        private bool IsScatter => Type == ChartType.Scatter || Type == ChartType.ScatterLines;

        /// <summary>Writes the complete chart part (<c>xl/charts/chartN.xml</c>).</summary>
        internal void WriteTo(XmlWriter w)
        {
            if (_series.Count == 0) throw new InvalidOperationException($"Chart '{Title ?? "#" + Number}' has no series. Call AddSeries at least once.");

            w.WriteStartDocument(true);
            w.WriteStartElement("c", "chartSpace", OpenXmlNamespaces.Chart);
            w.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingMain);
            w.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);
            Val(w, "roundedCorners", "0");

            w.WriteStartElement("c", "chart", OpenXmlNamespaces.Chart);
            if (Title != null) WriteTitle(w, Title);
            Val(w, "autoTitleDeleted", Title == null ? "1" : "0");

            w.WriteStartElement("c", "plotArea", OpenXmlNamespaces.Chart);
            w.WriteStartElement("c", "layout", OpenXmlNamespaces.Chart);
            w.WriteEndElement();
            WritePlot(w);
            if (!IsPieLike) WriteAxes(w);
            w.WriteEndElement(); // plotArea

            if (Legend != LegendPosition.None)
            {
                w.WriteStartElement("c", "legend", OpenXmlNamespaces.Chart);
                Val(w, "legendPos", Legend switch { LegendPosition.Left => "l", LegendPosition.Top => "t", LegendPosition.Bottom => "b", _ => "r" });
                Val(w, "overlay", "0");
                w.WriteEndElement();
            }
            Val(w, "plotVisOnly", "1");
            Val(w, "dispBlanksAs", "gap");
            w.WriteEndElement(); // chart
            w.WriteEndElement(); // chartSpace
            w.WriteEndDocument();
        }

        private void WritePlot(XmlWriter w)
        {
            switch (Type)
            {
                case ChartType.Column:
                case ChartType.ColumnStacked:
                case ChartType.ColumnStacked100:
                case ChartType.Bar:
                case ChartType.BarStacked:
                case ChartType.BarStacked100:
                {
                    var bar = Type == ChartType.Bar || Type == ChartType.BarStacked || Type == ChartType.BarStacked100;
                    var grouping = Type == ChartType.ColumnStacked || Type == ChartType.BarStacked ? "stacked"
                        : Type == ChartType.ColumnStacked100 || Type == ChartType.BarStacked100 ? "percentStacked" : "clustered";
                    w.WriteStartElement("c", "barChart", OpenXmlNamespaces.Chart);
                    Val(w, "barDir", bar ? "bar" : "col");
                    Val(w, "grouping", grouping);
                    Val(w, "varyColors", "0");
                    for (var i = 0; i < _series.Count; i++) WriteSeries(w, i, SeriesKind.Bar);
                    WriteDataLabels(w);
                    Val(w, "gapWidth", "150");
                    if (grouping != "clustered") Val(w, "overlap", "100");
                    Val(w, "axId", "500000001");
                    Val(w, "axId", "500000002");
                    w.WriteEndElement();
                    break;
                }
                case ChartType.Line:
                case ChartType.LineMarkers:
                case ChartType.LineStacked:
                    w.WriteStartElement("c", "lineChart", OpenXmlNamespaces.Chart);
                    Val(w, "grouping", Type == ChartType.LineStacked ? "stacked" : "standard");
                    Val(w, "varyColors", "0");
                    for (var i = 0; i < _series.Count; i++) WriteSeries(w, i, Type == ChartType.LineMarkers ? SeriesKind.LineMarkers : SeriesKind.Line);
                    WriteDataLabels(w);
                    Val(w, "marker", "1");
                    Val(w, "axId", "500000001");
                    Val(w, "axId", "500000002");
                    w.WriteEndElement();
                    break;
                case ChartType.Area:
                case ChartType.AreaStacked:
                    w.WriteStartElement("c", "areaChart", OpenXmlNamespaces.Chart);
                    Val(w, "grouping", Type == ChartType.AreaStacked ? "stacked" : "standard");
                    Val(w, "varyColors", "0");
                    for (var i = 0; i < _series.Count; i++) WriteSeries(w, i, SeriesKind.Area);
                    WriteDataLabels(w);
                    Val(w, "axId", "500000001");
                    Val(w, "axId", "500000002");
                    w.WriteEndElement();
                    break;
                case ChartType.Pie:
                case ChartType.Doughnut:
                    w.WriteStartElement("c", Type == ChartType.Pie ? "pieChart" : "doughnutChart", OpenXmlNamespaces.Chart);
                    Val(w, "varyColors", "1");
                    // A pie shows one series; extra series would be silently ignored by Excel, so write only the first.
                    WriteSeries(w, 0, SeriesKind.Pie);
                    WriteDataLabels(w);
                    Val(w, "firstSliceAng", "0");
                    if (Type == ChartType.Doughnut) Val(w, "holeSize", "50");
                    w.WriteEndElement();
                    break;
                default: // scatter
                    w.WriteStartElement("c", "scatterChart", OpenXmlNamespaces.Chart);
                    Val(w, "scatterStyle", "lineMarker");
                    Val(w, "varyColors", "0");
                    for (var i = 0; i < _series.Count; i++) WriteSeries(w, i, Type == ChartType.Scatter ? SeriesKind.ScatterMarkers : SeriesKind.ScatterLines);
                    WriteDataLabels(w);
                    Val(w, "axId", "500000001");
                    Val(w, "axId", "500000002");
                    w.WriteEndElement();
                    break;
            }
        }

        private enum SeriesKind { Bar, Line, LineMarkers, Area, Pie, ScatterMarkers, ScatterLines }

        /// <summary>
        /// Writes one series. Child order differs per chart family (CT_BarSer, CT_LineSer, CT_PieSer,
        /// CT_AreaSer, CT_ScatterSer) but always starts idx, order, tx, spPr.
        /// </summary>
        private void WriteSeries(XmlWriter w, int index, SeriesKind kind)
        {
            var s = _series[index];
            var idx = index.ToString(CultureInfo.InvariantCulture);
            w.WriteStartElement("c", "ser", OpenXmlNamespaces.Chart);
            Val(w, "idx", idx);
            Val(w, "order", idx);

            if (s.Name != null)
            {
                w.WriteStartElement("c", "tx", OpenXmlNamespaces.Chart);
                if (s.Name.StartsWith("=", StringComparison.Ordinal))
                {
                    w.WriteStartElement("c", "strRef", OpenXmlNamespaces.Chart);
                    w.WriteElementString("c", "f", OpenXmlNamespaces.Chart, s.Name.Substring(1));
                    w.WriteEndElement();
                }
                else
                {
                    w.WriteElementString("c", "v", OpenXmlNamespaces.Chart, XmlTextUtil.Sanitize(s.Name));
                }
                w.WriteEndElement();
            }

            var lineLike = kind == SeriesKind.Line || kind == SeriesKind.LineMarkers || kind == SeriesKind.ScatterLines;
            if (kind == SeriesKind.ScatterMarkers)
            {
                // Markers only: hide the connecting line.
                w.WriteStartElement("c", "spPr", OpenXmlNamespaces.Chart);
                w.WriteStartElement("a", "ln", OpenXmlNamespaces.DrawingMain);
                w.WriteAttributeString("w", "19050");
                w.WriteStartElement("a", "noFill", OpenXmlNamespaces.DrawingMain);
                w.WriteEndElement();
                w.WriteEndElement();
                w.WriteEndElement();
            }
            else if (s.Color != null && kind != SeriesKind.Pie)
            {
                w.WriteStartElement("c", "spPr", OpenXmlNamespaces.Chart);
                if (lineLike)
                {
                    w.WriteStartElement("a", "ln", OpenXmlNamespaces.DrawingMain);
                    w.WriteAttributeString("w", "28575");
                    SolidFill(w, s.Color);
                    w.WriteEndElement();
                }
                else
                {
                    SolidFill(w, s.Color);
                }
                w.WriteEndElement();
            }

            if (kind == SeriesKind.Bar) Val(w, "invertIfNegative", "0");
            if (kind == SeriesKind.Line || kind == SeriesKind.LineMarkers || kind == SeriesKind.ScatterMarkers || kind == SeriesKind.ScatterLines)
            {
                w.WriteStartElement("c", "marker", OpenXmlNamespaces.Chart);
                Val(w, "symbol", kind == SeriesKind.Line ? "none" : "circle");
                if (kind != SeriesKind.Line)
                {
                    Val(w, "size", "5");
                    // CT_Marker: symbol, size, spPr. The series colour must reach the markers too, otherwise a
                    // marker-only scatter series ignores it (its line is hidden).
                    if (s.Color != null)
                    {
                        w.WriteStartElement("c", "spPr", OpenXmlNamespaces.Chart);
                        SolidFill(w, s.Color);
                        w.WriteStartElement("a", "ln", OpenXmlNamespaces.DrawingMain);
                        SolidFill(w, s.Color);
                        w.WriteEndElement();
                        w.WriteEndElement();
                    }
                }
                w.WriteEndElement();
            }

            if (IsScatter)
            {
                if (s.Categories != null) Ref(w, "xVal", "numRef", s.Categories);
                Ref(w, "yVal", "numRef", s.Values);
                Val(w, "smooth", "0");
            }
            else
            {
                if (s.Categories != null) Ref(w, "cat", "strRef", s.Categories);
                Ref(w, "val", "numRef", s.Values);
                if (lineLike) Val(w, "smooth", "0");
            }
            w.WriteEndElement();
        }

        private void WriteDataLabels(XmlWriter w)
        {
            if (DataLabels == ChartDataLabels.None) return;
            w.WriteStartElement("c", "dLbls", OpenXmlNamespaces.Chart);
            Val(w, "showLegendKey", "0");
            Val(w, "showVal", DataLabels == ChartDataLabels.Value || DataLabels == ChartDataLabels.CategoryAndValue ? "1" : "0");
            Val(w, "showCatName", DataLabels == ChartDataLabels.CategoryAndValue ? "1" : "0");
            Val(w, "showSerName", "0");
            Val(w, "showPercent", DataLabels == ChartDataLabels.Percent ? "1" : "0");
            Val(w, "showBubbleSize", "0");
            w.WriteEndElement();
        }

        /// <summary>
        /// Writes the two axes. Bar charts put categories on the left (axPos "l") and values at the bottom;
        /// scatter charts use two value axes.
        /// </summary>
        private void WriteAxes(XmlWriter w)
        {
            var horizontalBars = Type == ChartType.Bar || Type == ChartType.BarStacked || Type == ChartType.BarStacked100;
            if (IsScatter) WriteValueAxis(w, "500000001", "500000002", "b", XAxisTitle, null, false);
            else WriteCategoryAxis(w, horizontalBars ? "l" : "b");
            WriteValueAxis(w, "500000002", "500000001", horizontalBars ? "b" : "l", YAxisTitle, ValueAxisNumberFormat, true);
        }

        private void WriteCategoryAxis(XmlWriter w, string position)
        {
            w.WriteStartElement("c", "catAx", OpenXmlNamespaces.Chart);
            Val(w, "axId", "500000001");
            Scaling(w);
            Val(w, "delete", "0");
            Val(w, "axPos", position);
            if (XAxisTitle != null) WriteTitle(w, XAxisTitle);
            Val(w, "majorTickMark", "out");
            Val(w, "minorTickMark", "none");
            Val(w, "tickLblPos", "nextTo");
            Val(w, "crossAx", "500000002");
            Val(w, "crosses", "autoZero");
            Val(w, "auto", "1");
            Val(w, "lblAlgn", "ctr");
            Val(w, "lblOffset", "100");
            Val(w, "noMultiLvlLbl", "0");
            w.WriteEndElement();
        }

        private static void WriteValueAxis(XmlWriter w, string id, string crossId, string position, string? title, string? format, bool gridlines)
        {
            w.WriteStartElement("c", "valAx", OpenXmlNamespaces.Chart);
            Val(w, "axId", id);
            Scaling(w);
            Val(w, "delete", "0");
            Val(w, "axPos", position);
            if (gridlines)
            {
                w.WriteStartElement("c", "majorGridlines", OpenXmlNamespaces.Chart);
                w.WriteEndElement();
            }
            if (title != null) WriteTitle(w, title);
            w.WriteStartElement("c", "numFmt", OpenXmlNamespaces.Chart);
            w.WriteAttributeString("formatCode", format ?? "General");
            w.WriteAttributeString("sourceLinked", format == null ? "1" : "0");
            w.WriteEndElement();
            Val(w, "majorTickMark", "out");
            Val(w, "minorTickMark", "none");
            Val(w, "tickLblPos", "nextTo");
            Val(w, "crossAx", crossId);
            Val(w, "crosses", "autoZero");
            Val(w, "crossBetween", "between");
            w.WriteEndElement();
        }

        private static void Scaling(XmlWriter w)
        {
            w.WriteStartElement("c", "scaling", OpenXmlNamespaces.Chart);
            Val(w, "orientation", "minMax");
            w.WriteEndElement();
        }

        private static void WriteTitle(XmlWriter w, string text)
        {
            w.WriteStartElement("c", "title", OpenXmlNamespaces.Chart);
            w.WriteStartElement("c", "tx", OpenXmlNamespaces.Chart);
            w.WriteStartElement("c", "rich", OpenXmlNamespaces.Chart);
            w.WriteStartElement("a", "bodyPr", OpenXmlNamespaces.DrawingMain);
            w.WriteEndElement();
            w.WriteStartElement("a", "lstStyle", OpenXmlNamespaces.DrawingMain);
            w.WriteEndElement();
            w.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingMain);
            w.WriteStartElement("a", "r", OpenXmlNamespaces.DrawingMain);
            w.WriteElementString("a", "t", OpenXmlNamespaces.DrawingMain, XmlTextUtil.Sanitize(text));
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            Val(w, "overlay", "0");
            w.WriteEndElement();
        }

        private static void Ref(XmlWriter w, string element, string refKind, string formula)
        {
            w.WriteStartElement("c", element, OpenXmlNamespaces.Chart);
            w.WriteStartElement("c", refKind, OpenXmlNamespaces.Chart);
            w.WriteElementString("c", "f", OpenXmlNamespaces.Chart, formula);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void SolidFill(XmlWriter w, string rgb)
        {
            w.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingMain);
            w.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingMain);
            w.WriteAttributeString("val", rgb);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void Val(XmlWriter w, string element, string value)
        {
            w.WriteStartElement("c", element, OpenXmlNamespaces.Chart);
            w.WriteAttributeString("val", value);
            w.WriteEndElement();
        }
    }

    /// <summary>A picture placed on a worksheet. Create with <see cref="Worksheet.AddPicture(byte[], int, int, int?, int?, string?)"/>.</summary>
    public sealed class ExcelPicture
    {
        internal ExcelPicture(byte[] data, string extension, int widthPx, int heightPx, int row, int column, string name)
        {
            Data = data;
            Extension = extension;
            WidthPixels = widthPx;
            HeightPixels = heightPx;
            Row = row;
            Column = column;
            Name = name;
        }

        internal byte[] Data { get; }
        internal string Extension { get; }

        /// <summary>Displayed width in pixels.</summary>
        public int WidthPixels { get; }

        /// <summary>Displayed height in pixels.</summary>
        public int HeightPixels { get; }

        /// <summary>1-based row of the top-left corner.</summary>
        public int Row { get; }

        /// <summary>1-based column of the top-left corner.</summary>
        public int Column { get; }

        /// <summary>Name shown in Excel's selection pane.</summary>
        public string Name { get; }

        /// <summary>
        /// Reads the pixel size and format from the image header. PNG, JPEG and GIF are supported; reading the
        /// header ourselves avoids System.Drawing, which is not available on every platform netstandard2.0 runs on.
        /// </summary>
        internal static (string Extension, int Width, int Height) Inspect(byte[] data)
        {
            if (data.Length >= 24 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
            {
                var wpx = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
                var hpx = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
                return ("png", wpx, hpx);
            }
            if (data.Length >= 10 && data[0] == 'G' && data[1] == 'I' && data[2] == 'F')
                return ("gif", data[6] | (data[7] << 8), data[8] | (data[9] << 8));
            if (data.Length >= 4 && data[0] == 0xFF && data[1] == 0xD8)
            {
                // Walk the JPEG segments until a start-of-frame marker, which carries height then width.
                var i = 2;
                while (i + 9 < data.Length)
                {
                    if (data[i] != 0xFF) { i++; continue; }
                    var marker = data[i + 1];
                    if (marker == 0xD8 || marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7)) { i += 2; continue; }
                    var length = (data[i + 2] << 8) | data[i + 3];
                    var isSof = marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;
                    if (isSof)
                    {
                        var h = (data[i + 5] << 8) | data[i + 6];
                        var wd = (data[i + 7] << 8) | data[i + 8];
                        return ("jpeg", wd, h);
                    }
                    i += 2 + length;
                }
            }
            throw new ArgumentException("Unsupported image format. SheetStream embeds PNG, JPEG and GIF images.", nameof(data));
        }
    }

    /// <summary>
    /// Writes a sheet's drawing part: one anchor per picture and chart, plus the drawing's relationships, the
    /// media files and the chart parts.
    /// </summary>
    internal static class DrawingWriter
    {
        /// <summary>EMUs (English Metric Units) per pixel at 96 DPI.</summary>
        private const long EmuPerPixel = 9525;

        public static void Write(ExcelPackage package, Workbook workbook, int drawingNumber, IReadOnlyList<ExcelPicture> pictures, IReadOnlyList<ExcelChart> charts)
        {
            var rels = new RelationshipSet();
            var drawingPart = "xl/drawings/drawing" + drawingNumber.ToString(CultureInfo.InvariantCulture) + ".xml";
            var mediaTargets = new List<(string Part, byte[] Data)>();

            using (var w = package.OpenPart(drawingPart))
            {
                w.WriteStartDocument(true);
                w.WriteStartElement("xdr", "wsDr", OpenXmlNamespaces.SpreadsheetDrawing);
                w.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingMain);
                w.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);
                w.WriteAttributeString("xmlns", "c", null, OpenXmlNamespaces.Chart);

                var shapeId = 2;
                foreach (var pic in pictures)
                {
                    var imageNumber = ++workbook.ImageCount;
                    var media = "xl/media/image" + imageNumber.ToString(CultureInfo.InvariantCulture) + "." + pic.Extension;
                    mediaTargets.Add((media, pic.Data));
                    var rid = rels.Add(OpenXmlNamespaces.RelImage, "../media/image" + imageNumber.ToString(CultureInfo.InvariantCulture) + "." + pic.Extension);
                    WritePicture(w, pic, rid, shapeId++);
                }
                foreach (var chart in charts)
                {
                    var rid = rels.Add(OpenXmlNamespaces.RelChart, "../charts/chart" + chart.Number.ToString(CultureInfo.InvariantCulture) + ".xml");
                    WriteChartFrame(w, chart, rid, shapeId++);
                }
                w.WriteEndElement();
                w.WriteEndDocument();
            }
            package.RegisterPart(drawingPart, OpenXmlNamespaces.CtDrawing);

            using (var w = package.OpenPart("xl/drawings/_rels/drawing" + drawingNumber.ToString(CultureInfo.InvariantCulture) + ".xml.rels"))
                rels.WriteTo(w);

            foreach (var (part, data) in mediaTargets)
            {
                package.WriteBinaryPart(part, data);
                var ext = part.Substring(part.LastIndexOf('.') + 1);
                package.RegisterDefault(ext, ext == "png" ? "image/png" : ext == "gif" ? "image/gif" : "image/jpeg");
            }

            foreach (var chart in charts)
            {
                var chartPart = "xl/charts/chart" + chart.Number.ToString(CultureInfo.InvariantCulture) + ".xml";
                using (var w = package.OpenPart(chartPart)) chart.WriteTo(w);
                package.RegisterPart(chartPart, OpenXmlNamespaces.CtChart);
            }
        }

        private static void WritePicture(XmlWriter w, ExcelPicture pic, string rid, int shapeId)
        {
            var cx = (pic.WidthPixels * EmuPerPixel).ToString(CultureInfo.InvariantCulture);
            var cy = (pic.HeightPixels * EmuPerPixel).ToString(CultureInfo.InvariantCulture);

            w.WriteStartElement("xdr", "oneCellAnchor", OpenXmlNamespaces.SpreadsheetDrawing);
            Marker(w, "from", pic.Row - 1, pic.Column - 1);
            w.WriteStartElement("xdr", "ext", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteAttributeString("cx", cx);
            w.WriteAttributeString("cy", cy);
            w.WriteEndElement();

            w.WriteStartElement("xdr", "pic", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteStartElement("xdr", "nvPicPr", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteStartElement("xdr", "cNvPr", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteAttributeString("id", shapeId.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("name", XmlTextUtil.Sanitize(pic.Name));
            w.WriteEndElement();
            w.WriteStartElement("xdr", "cNvPicPr", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteStartElement("a", "picLocks", OpenXmlNamespaces.DrawingMain);
            w.WriteAttributeString("noChangeAspect", "1");
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("xdr", "blipFill", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingMain);
            w.WriteAttributeString("r", "embed", OpenXmlNamespaces.Relationships, rid);
            w.WriteEndElement();
            w.WriteStartElement("a", "stretch", OpenXmlNamespaces.DrawingMain);
            w.WriteStartElement("a", "fillRect", OpenXmlNamespaces.DrawingMain);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("xdr", "spPr", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingMain);
            w.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingMain);
            w.WriteAttributeString("x", "0");
            w.WriteAttributeString("y", "0");
            w.WriteEndElement();
            w.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingMain);
            w.WriteAttributeString("cx", cx);
            w.WriteAttributeString("cy", cy);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingMain);
            w.WriteAttributeString("prst", "rect");
            w.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingMain);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement(); // pic

            w.WriteStartElement("xdr", "clientData", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteChartFrame(XmlWriter w, ExcelChart chart, string rid, int shapeId)
        {
            w.WriteStartElement("xdr", "twoCellAnchor", OpenXmlNamespaces.SpreadsheetDrawing);
            Marker(w, "from", chart.Anchor.FromRow - 1, chart.Anchor.FromColumn - 1);
            // "to" is exclusive in practice: the frame ends at the start of the cell after the anchor range.
            Marker(w, "to", chart.Anchor.ToRow, chart.Anchor.ToColumn);

            w.WriteStartElement("xdr", "graphicFrame", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteAttributeString("macro", "");
            w.WriteStartElement("xdr", "nvGraphicFramePr", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteStartElement("xdr", "cNvPr", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteAttributeString("id", shapeId.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("name", "Chart " + chart.Number.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
            w.WriteStartElement("xdr", "cNvGraphicFramePr", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("xdr", "xfrm", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingMain);
            w.WriteAttributeString("x", "0");
            w.WriteAttributeString("y", "0");
            w.WriteEndElement();
            w.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingMain);
            w.WriteAttributeString("cx", "0");
            w.WriteAttributeString("cy", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("a", "graphic", OpenXmlNamespaces.DrawingMain);
            w.WriteStartElement("a", "graphicData", OpenXmlNamespaces.DrawingMain);
            w.WriteAttributeString("uri", OpenXmlNamespaces.Chart);
            w.WriteStartElement("c", "chart", OpenXmlNamespaces.Chart);
            w.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, rid);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement(); // graphicFrame

            w.WriteStartElement("xdr", "clientData", OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void Marker(XmlWriter w, string name, int row0, int col0)
        {
            w.WriteStartElement("xdr", name, OpenXmlNamespaces.SpreadsheetDrawing);
            w.WriteElementString("xdr", "col", OpenXmlNamespaces.SpreadsheetDrawing, col0.ToString(CultureInfo.InvariantCulture));
            w.WriteElementString("xdr", "colOff", OpenXmlNamespaces.SpreadsheetDrawing, "0");
            w.WriteElementString("xdr", "row", OpenXmlNamespaces.SpreadsheetDrawing, row0.ToString(CultureInfo.InvariantCulture));
            w.WriteElementString("xdr", "rowOff", OpenXmlNamespaces.SpreadsheetDrawing, "0");
            w.WriteEndElement();
        }
    }
}
