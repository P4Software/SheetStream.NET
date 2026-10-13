using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream.Styles
{
    /// <summary>
    /// The workbook's style registry. Translates <see cref="StyleOptions"/> into the five interlocking
    /// tables SpreadsheetML needs (number formats, fonts, fills, borders and cell formats), de-duplicating
    /// each so a workbook with a thousand "bold blue header" cells carries a single font record.
    /// </summary>
    /// <remarks>
    /// Styles can be created at any time before the package is saved; the <c>styles.xml</c> part is
    /// written last, after every sheet has been streamed. Not thread-safe.
    /// </remarks>
    public sealed class StyleSheet
    {
        /// <summary>Custom number formats start here; ids below 164 are reserved for Excel's built-ins.</summary>
        private const int FirstCustomNumberFormatId = 164;

        // Built-in number formats (ECMA-376 §18.8.30). A user format matching one of these is mapped to the
        // built-in id instead of a custom entry, which keeps styles.xml small and renders locale-aware.
        private static readonly Dictionary<string, int> BuiltInNumberFormats = new Dictionary<string, int>(StringComparer.Ordinal)
        {
            ["General"] = 0, ["0"] = 1, ["0.00"] = 2, ["#,##0"] = 3, ["#,##0.00"] = 4,
            ["0%"] = 9, ["0.00%"] = 10, ["0.00E+00"] = 11, ["# ?/?"] = 12, ["# ??/??"] = 13,
            ["mm-dd-yy"] = 14, ["d-mmm-yy"] = 15, ["d-mmm"] = 16, ["mmm-yy"] = 17,
            ["h:mm AM/PM"] = 18, ["h:mm:ss AM/PM"] = 19, ["h:mm"] = 20, ["h:mm:ss"] = 21, ["m/d/yy h:mm"] = 22,
            ["#,##0 ;(#,##0)"] = 37, ["#,##0 ;[Red](#,##0)"] = 38, ["#,##0.00;(#,##0.00)"] = 39, ["#,##0.00;[Red](#,##0.00)"] = 40,
            ["mm:ss"] = 45, ["[h]:mm:ss"] = 46, ["mmss.0"] = 47, ["##0.0E+0"] = 48, ["@"] = 49,
        };

        private readonly List<FontSpec> _fonts = new List<FontSpec>();
        private readonly Dictionary<FontSpec, int> _fontIndex = new Dictionary<FontSpec, int>();

        private readonly List<string?> _fills = new List<string?>();           // null = none / gray125 placeholders
        private readonly Dictionary<string, int> _fillIndex = new Dictionary<string, int>(StringComparer.Ordinal);

        private readonly List<BorderSpec> _borders = new List<BorderSpec>();
        private readonly Dictionary<BorderSpec, int> _borderIndex = new Dictionary<BorderSpec, int>();

        private readonly List<KeyValuePair<int, string>> _customNumberFormats = new List<KeyValuePair<int, string>>();
        private readonly Dictionary<string, int> _numberFormatIndex = new Dictionary<string, int>(StringComparer.Ordinal);

        private readonly List<XfSpec> _xfs = new List<XfSpec>();
        private readonly Dictionary<XfSpec, int> _xfIndex = new Dictionary<XfSpec, int>();

        private CellStyle? _dateStyle;
        private CellStyle? _dateTimeStyle;
        private CellStyle? _durationStyle;

        internal StyleSheet()
        {
            // Index 0 of every table is the workbook default and must exist even if never referenced.
            _fonts.Add(FontSpec.Default);
            _fontIndex[FontSpec.Default] = 0;

            // Excel requires fills[0] = none and fills[1] = gray125; custom fills start at index 2.
            _fills.Add(null);
            _fills.Add(null);

            _borders.Add(BorderSpec.Empty);
            _borderIndex[BorderSpec.Empty] = 0;

            var defaultXf = new XfSpec(0, 0, 0, 0, HorizontalAlignment.General, VerticalAlignment.Bottom, false);
            _xfs.Add(defaultXf);
            _xfIndex[defaultXf] = 0;
            Default = new CellStyle(0);
        }

        /// <summary>The workbook default style (Calibri 11, no fill, no border, General format).</summary>
        public CellStyle Default { get; }

        /// <summary>Number of registered cell styles, including the default.</summary>
        public int Count => _xfs.Count;

        /// <summary>
        /// Creates (or reuses) a style from a configuration callback.
        /// </summary>
        /// <example>
        /// <code>
        /// var header = workbook.Styles.Create(o =>
        /// {
        ///     o.Font.Bold = true;
        ///     o.Font.Color = "#FFFFFF";
        ///     o.Fill.BackgroundColor = "#4F81BD";
        /// });
        /// </code>
        /// </example>
        public CellStyle Create(Action<StyleOptions> configure)
        {
            if (configure == null) throw new ArgumentNullException(nameof(configure));
            var options = new StyleOptions();
            configure(options);
            return Create(options);
        }

        /// <summary>Creates (or reuses) a style from a fully populated <see cref="StyleOptions"/>.</summary>
        public CellStyle Create(StyleOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            var fontId = RegisterFont(FontSpec.From(options.Font));
            var fillId = RegisterFill(options.Fill.BackgroundColor);
            var borderId = RegisterBorder(BorderSpec.From(options.Border));
            var numFmtId = RegisterNumberFormat(options.NumberFormat);

            var xf = new XfSpec(numFmtId, fontId, fillId, borderId, options.HorizontalAlignment, options.VerticalAlignment, options.WrapText);
            return new CellStyle(RegisterXf(xf));
        }

        /// <summary>Style applied to <see cref="DateTime"/> cells written without an explicit style (built-in format 14).</summary>
        internal CellStyle DateStyle => _dateStyle ??= Create(o => o.NumberFormat = "mm-dd-yy");

        /// <summary>Style applied to date-with-time cells written without an explicit style (built-in format 22).</summary>
        internal CellStyle DateTimeStyle => _dateTimeStyle ??= Create(o => o.NumberFormat = "m/d/yy h:mm");

        /// <summary>Style applied to <see cref="TimeSpan"/> cells written without an explicit style (built-in format 46).</summary>
        internal CellStyle DurationStyle => _durationStyle ??= Create(o => o.NumberFormat = NumberFormats.Duration);

        private int RegisterFont(FontSpec font)
        {
            if (_fontIndex.TryGetValue(font, out var id)) return id;
            id = _fonts.Count;
            _fonts.Add(font);
            _fontIndex[font] = id;
            return id;
        }

        private int RegisterFill(string? color)
        {
            if (color == null) return 0;
            if (_fillIndex.TryGetValue(color, out var id)) return id;
            id = _fills.Count;
            _fills.Add(color);
            _fillIndex[color] = id;
            return id;
        }

        private int RegisterBorder(BorderSpec border)
        {
            if (_borderIndex.TryGetValue(border, out var id)) return id;
            id = _borders.Count;
            _borders.Add(border);
            _borderIndex[border] = id;
            return id;
        }

        private int RegisterNumberFormat(string? formatCode)
        {
            if (string.IsNullOrEmpty(formatCode)) return 0;
            if (BuiltInNumberFormats.TryGetValue(formatCode!, out var builtIn)) return builtIn;
            if (_numberFormatIndex.TryGetValue(formatCode!, out var id)) return id;
            id = FirstCustomNumberFormatId + _customNumberFormats.Count;
            _customNumberFormats.Add(new KeyValuePair<int, string>(id, formatCode!));
            _numberFormatIndex[formatCode!] = id;
            return id;
        }

        private int RegisterXf(XfSpec xf)
        {
            if (_xfIndex.TryGetValue(xf, out var id)) return id;
            id = _xfs.Count;
            _xfs.Add(xf);
            _xfIndex[xf] = id;
            return id;
        }

        /// <summary>
        /// Writes the complete <c>styles.xml</c> part. Element order follows CT_Stylesheet exactly:
        /// numFmts, fonts, fills, borders, cellStyleXfs, cellXfs, cellStyles. Excel rejects any other order.
        /// </summary>
        internal void WriteTo(XmlWriter w)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("styleSheet", OpenXmlNamespaces.Main);

            if (_customNumberFormats.Count > 0)
            {
                w.WriteStartElement("numFmts");
                WriteCount(w, _customNumberFormats.Count);
                foreach (var pair in _customNumberFormats)
                {
                    w.WriteStartElement("numFmt");
                    w.WriteAttributeString("numFmtId", pair.Key.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("formatCode", XmlTextUtil.Sanitize(pair.Value));
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }

            w.WriteStartElement("fonts");
            WriteCount(w, _fonts.Count);
            foreach (var font in _fonts) font.WriteTo(w);
            w.WriteEndElement();

            w.WriteStartElement("fills");
            WriteCount(w, _fills.Count);
            WritePatternFill(w, "none", null);
            WritePatternFill(w, "gray125", null);
            for (var i = 2; i < _fills.Count; i++) WritePatternFill(w, "solid", _fills[i]);
            w.WriteEndElement();

            w.WriteStartElement("borders");
            WriteCount(w, _borders.Count);
            foreach (var border in _borders) border.WriteTo(w);
            w.WriteEndElement();

            // A single master "Normal" cell style; every cellXf points at it through xfId="0".
            w.WriteStartElement("cellStyleXfs");
            WriteCount(w, 1);
            w.WriteStartElement("xf");
            w.WriteAttributeString("numFmtId", "0");
            w.WriteAttributeString("fontId", "0");
            w.WriteAttributeString("fillId", "0");
            w.WriteAttributeString("borderId", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("cellXfs");
            WriteCount(w, _xfs.Count);
            foreach (var xf in _xfs) xf.WriteTo(w);
            w.WriteEndElement();

            w.WriteStartElement("cellStyles");
            WriteCount(w, 1);
            w.WriteStartElement("cellStyle");
            w.WriteAttributeString("name", "Normal");
            w.WriteAttributeString("xfId", "0");
            w.WriteAttributeString("builtinId", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteEndElement(); // styleSheet
            w.WriteEndDocument();
        }

        private static void WriteCount(XmlWriter w, int count)
            => w.WriteAttributeString("count", count.ToString(CultureInfo.InvariantCulture));

        /// <summary>
        /// Writes one fill. For solid fills the colour goes in <c>fgColor</c>, not <c>bgColor</c>: in
        /// SpreadsheetML the "foreground" of a pattern is the pattern ink, and a solid pattern is all ink.
        /// </summary>
        private static void WritePatternFill(XmlWriter w, string patternType, string? argb)
        {
            w.WriteStartElement("fill");
            w.WriteStartElement("patternFill");
            w.WriteAttributeString("patternType", patternType);
            if (argb != null)
            {
                w.WriteStartElement("fgColor");
                w.WriteAttributeString("rgb", argb);
                w.WriteEndElement();
                w.WriteStartElement("bgColor");
                w.WriteAttributeString("indexed", "64");
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static string ToXml(BorderStyle style) => style switch
        {
            BorderStyle.Hair => "hair",
            BorderStyle.Thin => "thin",
            BorderStyle.Medium => "medium",
            BorderStyle.Thick => "thick",
            BorderStyle.Dashed => "dashed",
            BorderStyle.Dotted => "dotted",
            BorderStyle.Double => "double",
            _ => "none",
        };

        private static string ToXml(HorizontalAlignment alignment) => alignment switch
        {
            HorizontalAlignment.Left => "left",
            HorizontalAlignment.Center => "center",
            HorizontalAlignment.Right => "right",
            HorizontalAlignment.Fill => "fill",
            HorizontalAlignment.Justify => "justify",
            _ => "general",
        };

        private static string ToXml(VerticalAlignment alignment) => alignment switch
        {
            VerticalAlignment.Top => "top",
            VerticalAlignment.Center => "center",
            VerticalAlignment.Justify => "justify",
            _ => "bottom",
        };

        /// <summary>Immutable, hashable snapshot of a font. Used as the de-duplication key.</summary>
        private readonly struct FontSpec : IEquatable<FontSpec>
        {
            public static readonly FontSpec Default = new FontSpec(false, false, false, false, 11, "Calibri", null);

            public readonly bool Bold, Italic, Underline, Strike;
            public readonly double Size;
            public readonly string Name;
            public readonly string? Color;

            public FontSpec(bool bold, bool italic, bool underline, bool strike, double size, string name, string? color)
            {
                Bold = bold; Italic = italic; Underline = underline; Strike = strike;
                Size = size; Name = name; Color = color;
            }

            public static FontSpec From(FontOptions f)
            {
                if (f.Size <= 0) throw new ArgumentOutOfRangeException(nameof(FontOptions.Size), f.Size, "Font size must be positive.");
                if (string.IsNullOrWhiteSpace(f.Name)) throw new ArgumentException("Font name cannot be empty.", nameof(FontOptions.Name));
                return new FontSpec(f.Bold, f.Italic, f.Underline, f.Strikethrough, f.Size, f.Name, f.Color);
            }

            private (bool, bool, bool, bool, double, string, string?) Key => (Bold, Italic, Underline, Strike, Size, Name, Color);
            public bool Equals(FontSpec other) => Key.Equals(other.Key);
            public override bool Equals(object? obj) => obj is FontSpec other && Equals(other);
            public override int GetHashCode() => Key.GetHashCode();

            public void WriteTo(XmlWriter w)
            {
                w.WriteStartElement("font");
                if (Bold) WriteEmpty(w, "b");
                if (Italic) WriteEmpty(w, "i");
                if (Strike) WriteEmpty(w, "strike");
                if (Underline) WriteEmpty(w, "u");
                w.WriteStartElement("sz");
                w.WriteAttributeString("val", Size.ToString(CultureInfo.InvariantCulture));
                w.WriteEndElement();
                if (Color != null)
                {
                    w.WriteStartElement("color");
                    w.WriteAttributeString("rgb", Color);
                    w.WriteEndElement();
                }
                w.WriteStartElement("name");
                w.WriteAttributeString("val", Name);
                w.WriteEndElement();
                w.WriteStartElement("family");
                w.WriteAttributeString("val", "2");
                w.WriteEndElement();
                // Only the theme font may carry a scheme; declaring one on a custom font would let the
                // theme override the user's chosen name.
                if (string.Equals(Name, "Calibri", StringComparison.OrdinalIgnoreCase))
                {
                    w.WriteStartElement("scheme");
                    w.WriteAttributeString("val", "minor");
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }

            private static void WriteEmpty(XmlWriter w, string name)
            {
                w.WriteStartElement(name);
                w.WriteEndElement();
            }
        }

        /// <summary>Immutable, hashable snapshot of a border. Used as the de-duplication key.</summary>
        private readonly struct BorderSpec : IEquatable<BorderSpec>
        {
            public static readonly BorderSpec Empty = new BorderSpec(BorderStyle.None, BorderStyle.None, BorderStyle.None, BorderStyle.None, null);

            public readonly BorderStyle Left, Right, Top, Bottom;
            public readonly string? Color;

            public BorderSpec(BorderStyle left, BorderStyle right, BorderStyle top, BorderStyle bottom, string? color)
            {
                Left = left; Right = right; Top = top; Bottom = bottom;
                // A colour on an invisible border is meaningless; dropping it improves de-duplication.
                Color = (left == BorderStyle.None && right == BorderStyle.None && top == BorderStyle.None && bottom == BorderStyle.None) ? null : color;
            }

            public static BorderSpec From(BorderOptions b) => new BorderSpec(b.Left, b.Right, b.Top, b.Bottom, b.Color);

            private (BorderStyle, BorderStyle, BorderStyle, BorderStyle, string?) Key => (Left, Right, Top, Bottom, Color);
            public bool Equals(BorderSpec other) => Key.Equals(other.Key);
            public override bool Equals(object? obj) => obj is BorderSpec other && Equals(other);
            public override int GetHashCode() => Key.GetHashCode();

            /// <summary>Edge order is fixed by CT_Border: left, right, top, bottom, diagonal.</summary>
            public void WriteTo(XmlWriter w)
            {
                w.WriteStartElement("border");
                WriteEdge(w, "left", Left);
                WriteEdge(w, "right", Right);
                WriteEdge(w, "top", Top);
                WriteEdge(w, "bottom", Bottom);
                w.WriteStartElement("diagonal");
                w.WriteEndElement();
                w.WriteEndElement();
            }

            private void WriteEdge(XmlWriter w, string edge, BorderStyle style)
            {
                w.WriteStartElement(edge);
                if (style != BorderStyle.None)
                {
                    w.WriteAttributeString("style", ToXml(style));
                    w.WriteStartElement("color");
                    if (Color != null) w.WriteAttributeString("rgb", Color);
                    else w.WriteAttributeString("auto", "1");
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }
        }

        /// <summary>Immutable, hashable cell format (one <c>cellXfs/xf</c> entry).</summary>
        private readonly struct XfSpec : IEquatable<XfSpec>
        {
            public readonly int NumFmtId, FontId, FillId, BorderId;
            public readonly HorizontalAlignment HAlign;
            public readonly VerticalAlignment VAlign;
            public readonly bool Wrap;

            public XfSpec(int numFmtId, int fontId, int fillId, int borderId, HorizontalAlignment hAlign, VerticalAlignment vAlign, bool wrap)
            {
                NumFmtId = numFmtId; FontId = fontId; FillId = fillId; BorderId = borderId;
                HAlign = hAlign; VAlign = vAlign; Wrap = wrap;
            }

            private (int, int, int, int, HorizontalAlignment, VerticalAlignment, bool) Key => (NumFmtId, FontId, FillId, BorderId, HAlign, VAlign, Wrap);
            public bool Equals(XfSpec other) => Key.Equals(other.Key);
            public override bool Equals(object? obj) => obj is XfSpec other && Equals(other);
            public override int GetHashCode() => Key.GetHashCode();

            public void WriteTo(XmlWriter w)
            {
                var hasAlignment = HAlign != HorizontalAlignment.General || VAlign != VerticalAlignment.Bottom || Wrap;

                w.WriteStartElement("xf");
                w.WriteAttributeString("numFmtId", NumFmtId.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("fontId", FontId.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("fillId", FillId.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("borderId", BorderId.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("xfId", "0");
                // The apply* flags tell Excel the xf overrides the parent cell style; without them
                // a fill or font can be silently ignored.
                if (NumFmtId != 0) w.WriteAttributeString("applyNumberFormat", "1");
                if (FontId != 0) w.WriteAttributeString("applyFont", "1");
                if (FillId != 0) w.WriteAttributeString("applyFill", "1");
                if (BorderId != 0) w.WriteAttributeString("applyBorder", "1");
                if (hasAlignment)
                {
                    w.WriteAttributeString("applyAlignment", "1");
                    w.WriteStartElement("alignment");
                    if (HAlign != HorizontalAlignment.General) w.WriteAttributeString("horizontal", ToXml(HAlign));
                    if (VAlign != VerticalAlignment.Bottom) w.WriteAttributeString("vertical", ToXml(VAlign));
                    if (Wrap) w.WriteAttributeString("wrapText", "1");
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }
        }
    }
}
