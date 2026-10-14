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

        private readonly List<DxfSpec> _dxfs = new List<DxfSpec>();
        private readonly Dictionary<DxfSpec, int> _dxfIndex = new Dictionary<DxfSpec, int>();

        private CellStyle? _dateStyle;
        private CellStyle? _dateTimeStyle;
        private CellStyle? _durationStyle;
        private CellStyle? _hyperlinkStyle;
        private int _hyperlinkFontId = -1;

        // Template ("append") mode: the template's own fonts, fills, borders, cell formats and differential
        // formats keep their indexes 0..base-1 and new ones are numbered after them. In normal mode every base
        // is 0 and the lists start with the mandatory defaults.
        private bool _appendMode;
        private int _fontBase, _fillBase, _borderBase, _xfBase, _dxfBase;
        private int _nextNumberFormatId = FirstCustomNumberFormatId;
        private readonly List<XfSpec> _templateXfs = new List<XfSpec>();
        private readonly Dictionary<string, int> _templateCounts = new Dictionary<string, int>(StringComparer.Ordinal);

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

            var defaultXf = XfSpec.Default;
            _xfs.Add(defaultXf);
            _xfIndex[defaultXf] = 0;
            Default = new CellStyle(0);
        }

        /// <summary>The workbook default style (Calibri 11, no fill, no border, General format).</summary>
        public CellStyle Default { get; }

        /// <summary>Number of registered cell styles, including the default (and the template's, when filling one).</summary>
        public int Count => _xfBase + _xfs.Count;

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

            if (options.Indent < 0 || options.Indent > 250) throw new ArgumentOutOfRangeException(nameof(StyleOptions.Indent), options.Indent, "Indent must be between 0 and 250.");
            if ((options.TextRotation < 0 || options.TextRotation > 180) && options.TextRotation != 255)
                throw new ArgumentOutOfRangeException(nameof(StyleOptions.TextRotation), options.TextRotation, "Text rotation must be 0 to 180, or 255 for vertical text.");

            var xf = new XfSpec(numFmtId, fontId, fillId, borderId, options.HorizontalAlignment, options.VerticalAlignment, options.WrapText,
                options.ShrinkToFit, options.Indent, options.TextRotation, options.Protection.Locked, options.Protection.HideFormula, 0);
            return new CellStyle(RegisterXf(xf));
        }

        /// <summary>
        /// Creates (or reuses) a differential format for conditional formatting. Only the font colour,
        /// bold / italic / underline / strikethrough, fill colour, borders and number format are used; font
        /// name and size are ignored because Excel does not apply them conditionally.
        /// </summary>
        /// <example>
        /// <code>
        /// var bad = workbook.Styles.CreateDifferential(o => { o.Font.Color = "#9C0006"; o.Fill.BackgroundColor = "#FFC7CE"; });
        /// </code>
        /// </example>
        public DifferentialStyle CreateDifferential(Action<StyleOptions> configure)
        {
            if (configure == null) throw new ArgumentNullException(nameof(configure));
            var o = new StyleOptions();
            configure(o);
            string? numFmtCode = string.IsNullOrEmpty(o.NumberFormat) ? null : o.NumberFormat;
            var numFmtId = numFmtCode == null ? -1 : RegisterNumberFormat(numFmtCode);
            var spec = new DxfSpec(o.Font.Bold, o.Font.Italic, o.Font.Underline, o.Font.Strikethrough, o.Font.Color,
                o.Fill.BackgroundColor, BorderSpec.From(o.Border), numFmtId, numFmtCode);
            if (!_dxfIndex.TryGetValue(spec, out var id))
            {
                id = _dxfBase + _dxfs.Count;
                _dxfs.Add(spec);
                _dxfIndex[spec] = id;
            }
            return new DifferentialStyle(id);
        }

        /// <summary>
        /// Excel's built-in "Hyperlink" cell style (blue, underlined). Hyperlink cells use it unless given
        /// another style, so they look and behave like links typed in Excel.
        /// </summary>
        public CellStyle Hyperlink
        {
            get
            {
                if (_hyperlinkStyle != null) return _hyperlinkStyle;
                _hyperlinkFontId = RegisterFont(new FontSpec(false, false, true, false, 11, "Calibri", "FF0563C1"));
                // The built-in Hyperlink master style (cellStyleXfs[1]) only exists in workbooks we write from
                // scratch; in a template the index could point at an unrelated style, so parent to Normal.
                var xf = new XfSpec(0, _hyperlinkFontId, 0, 0, HorizontalAlignment.General, VerticalAlignment.Bottom, false,
                    false, 0, 0, true, false, _appendMode ? 0 : 1);
                _hyperlinkStyle = new CellStyle(RegisterXf(xf));
                return _hyperlinkStyle;
            }
        }

        /// <summary>
        /// Approximate width multiplier of a style's font relative to Calibri 11 regular, used by AutoFit.
        /// Bold text is about 10% wider; size scales linearly.
        /// </summary>
        internal double WidthFactor(int styleId)
        {
            if (styleId <= 0 || !TryGetXf(styleId, out var xf)) return 1.0;
            var local = xf.FontId - _fontBase;
            if (local < 0 || local >= _fonts.Count) return 1.0; // A template font: metrics unknown.
            var font = _fonts[local];
            return font.Size / 11.0 * (font.Bold ? 1.1 : 1.0);
        }

        /// <summary>Number format id of a style, used by AutoFit to estimate the displayed width of numbers.</summary>
        internal int NumberFormatId(int styleId) => styleId > 0 && TryGetXf(styleId, out var xf) ? xf.NumFmtId : 0;

        /// <summary>Looks up a cell format by its global index, whether it came from the template or was added here.</summary>
        private bool TryGetXf(int id, out XfSpec xf)
        {
            if (id >= 0 && id < _xfBase && id < _templateXfs.Count) { xf = _templateXfs[id]; return true; }
            var local = id - _xfBase;
            if (local >= 0 && local < _xfs.Count) { xf = _xfs[local]; return true; }
            xf = default;
            return false;
        }

        /// <summary>True when <paramref name="id"/> is a valid cell format index for this workbook.</summary>
        internal bool IsValidStyleId(int id) => id >= 0 && id < _xfBase + _xfs.Count;

        /// <summary>Style applied to <see cref="DateTime"/> cells written without an explicit style (built-in format 14).</summary>
        internal CellStyle DateStyle => _dateStyle ??= Create(o => o.NumberFormat = "mm-dd-yy");

        /// <summary>Style applied to date-with-time cells written without an explicit style (built-in format 22).</summary>
        internal CellStyle DateTimeStyle => _dateTimeStyle ??= Create(o => o.NumberFormat = "m/d/yy h:mm");

        /// <summary>Style applied to <see cref="TimeSpan"/> cells written without an explicit style (built-in format 46).</summary>
        internal CellStyle DurationStyle => _durationStyle ??= Create(o => o.NumberFormat = NumberFormats.Duration);

        /// <summary>
        /// Returns a style identical to <paramref name="baseStyle"/> but with <paramref name="formatCode"/> as its
        /// number format, unless the base already has a non-General format, in which case the base is returned.
        /// Used so a date written on a row that has a default style (zebra striping, say) keeps the fill but
        /// still displays as a date instead of a serial number.
        /// </summary>
        internal CellStyle WithNumberFormat(CellStyle baseStyle, string formatCode)
        {
            if (!TryGetXf(baseStyle.Id, out var xf) || xf.NumFmtId != 0) return baseStyle;
            var derived = xf.WithNumberFormat(RegisterNumberFormat(formatCode));
            return new CellStyle(RegisterXf(derived));
        }

        private int RegisterFont(FontSpec font)
        {
            if (_fontIndex.TryGetValue(font, out var id)) return id;
            id = _fontBase + _fonts.Count;
            _fonts.Add(font);
            _fontIndex[font] = id;
            return id;
        }

        private int RegisterFill(string? color)
        {
            if (color == null) return 0;
            if (_fillIndex.TryGetValue(color, out var id)) return id;
            id = _fillBase + _fills.Count;
            _fills.Add(color);
            _fillIndex[color] = id;
            return id;
        }

        private int RegisterBorder(BorderSpec border)
        {
            if (_borderIndex.TryGetValue(border, out var id)) return id;
            id = _borderBase + _borders.Count;
            _borders.Add(border);
            _borderIndex[border] = id;
            return id;
        }

        private int RegisterNumberFormat(string? formatCode)
        {
            if (string.IsNullOrEmpty(formatCode)) return 0;
            if (BuiltInNumberFormats.TryGetValue(formatCode!, out var builtIn)) return builtIn;
            if (_numberFormatIndex.TryGetValue(formatCode!, out var id)) return id;
            id = _nextNumberFormatId++;
            _customNumberFormats.Add(new KeyValuePair<int, string>(id, formatCode!));
            _numberFormatIndex[formatCode!] = id;
            return id;
        }

        private int RegisterXf(XfSpec xf)
        {
            if (_xfIndex.TryGetValue(xf, out var id)) return id;
            id = _xfBase + _xfs.Count;
            _xfs.Add(xf);
            _xfIndex[xf] = id;
            return id;
        }

        /// <summary>
        /// Switches to template mode by reading the template's <c>styles.xml</c>: records how many fonts, fills,
        /// borders, cell formats and differential formats it has, remembers its custom number formats (so equal
        /// codes are reused) and parses its cell formats (so a template style can be combined with a date format).
        /// </summary>
        internal void AttachTemplate(XmlReader r)
        {
            _appendMode = true;
            _fonts.Clear(); _fills.Clear(); _borders.Clear(); _xfs.Clear();
            _fontIndex.Clear(); _fillIndex.Clear(); _borderIndex.Clear(); _xfIndex.Clear();

            var maxNumFmt = FirstCustomNumberFormatId - 1;
            string? collection = null;
            var collectionDepth = -1;
            XfSpec? pending = null;
            var pendingDepth = -1;
            while (r.Read())
            {
                if (r.NodeType == XmlNodeType.EndElement)
                {
                    if (pending.HasValue && r.Depth == pendingDepth) { _templateXfs.Add(pending.Value); pending = null; }
                    if (r.Depth == collectionDepth) { collection = null; collectionDepth = -1; }
                    continue;
                }
                if (r.NodeType != XmlNodeType.Element) continue;

                if (r.Depth == 1)
                {
                    collection = r.LocalName;
                    collectionDepth = 1;
                    _templateCounts[collection] = 0;
                    if (r.IsEmptyElement) { collection = null; collectionDepth = -1; }
                    continue;
                }
                if (collection != null && r.Depth == 2)
                {
                    _templateCounts[collection]++;
                    if (collection == "numFmts" && r.LocalName == "numFmt")
                    {
                        var id = TemplateSource.Int(r.GetAttribute("numFmtId")) ?? 0;
                        if (id > maxNumFmt) maxNumFmt = id;
                        var code = r.GetAttribute("formatCode");
                        if (code != null && !_numberFormatIndex.ContainsKey(code)) _numberFormatIndex[code] = id;
                    }
                    else if (collection == "cellXfs" && r.LocalName == "xf")
                    {
                        var xf = XfSpec.FromTemplate(r);
                        if (r.IsEmptyElement) _templateXfs.Add(xf);
                        else { pending = xf; pendingDepth = r.Depth; }
                    }
                    continue;
                }
                if (pending.HasValue && r.Depth == pendingDepth + 1)
                {
                    if (r.LocalName == "alignment") pending = pending.Value.WithTemplateAlignment(r);
                    else if (r.LocalName == "protection") pending = pending.Value.WithTemplateProtection(r);
                }
            }

            _fontBase = Count("fonts");
            _fillBase = Count("fills");
            _borderBase = Count("borders");
            _xfBase = Count("cellXfs");
            _dxfBase = Count("dxfs");
            _nextNumberFormatId = maxNumFmt + 1;

            // Our defaults map onto the template's defaults, so plain cells inherit the template's look.
            _fontIndex[FontSpec.Default] = 0;
            _borderIndex[BorderSpec.Empty] = 0;
            _xfIndex[XfSpec.Default] = 0;

            int Count(string name) => _templateCounts.TryGetValue(name, out var n) ? n : 0;
        }

        /// <summary>
        /// Writes the template's <c>styles.xml</c> with the styles created here appended to each collection and
        /// the counts corrected. Collections the template lacks (number formats, differential formats) are
        /// inserted at their schema position.
        /// </summary>
        internal void WriteMerged(XmlReader r, XmlWriter w)
        {
            w.WriteStartDocument(true);
            XmlCopy.MoveToRoot(r);
            XmlCopy.CopyStartElement(r, w);
            r.Read();
            var wroteNumFmts = false;
            var wroteDxfs = false;
            while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == 0))
            {
                if (r.NodeType != XmlNodeType.Element || r.Depth != 1)
                {
                    XmlCopy.CopyNode(r, w);
                    r.Read();
                    continue;
                }
                var name = r.LocalName;
                if (!wroteNumFmts && _customNumberFormats.Count > 0 && name != "numFmts")
                {
                    WriteNumberFormats(w, _customNumberFormats.Count, null);
                    wroteNumFmts = true;
                }
                if (!wroteDxfs && _dxfs.Count > 0 && (name == "tableStyles" || name == "colors" || name == "extLst"))
                {
                    WriteDifferentials(w, _dxfs.Count, null);
                    wroteDxfs = true;
                }
                switch (name)
                {
                    case "numFmts":
                        MergeCollection(r, w, _customNumberFormats.Count, () => WriteNumberFormatItems(w));
                        wroteNumFmts = true;
                        break;
                    case "fonts":
                        MergeCollection(r, w, _fonts.Count, () => { foreach (var f in _fonts) f.WriteTo(w); });
                        break;
                    case "fills":
                        MergeCollection(r, w, _fills.Count, () => { foreach (var f in _fills) WritePatternFill(w, "solid", f); });
                        break;
                    case "borders":
                        MergeCollection(r, w, _borders.Count, () => { foreach (var b in _borders) b.WriteTo(w); });
                        break;
                    case "cellXfs":
                        MergeCollection(r, w, _xfs.Count, () => { foreach (var x in _xfs) x.WriteTo(w); });
                        break;
                    case "dxfs":
                        MergeCollection(r, w, _dxfs.Count, () => { foreach (var d in _dxfs) d.WriteTo(w); });
                        wroteDxfs = true;
                        break;
                    default:
                        XmlCopy.CopySubtree(r, w);
                        break;
                }
            }
            if (!wroteDxfs && _dxfs.Count > 0) WriteDifferentials(w, _dxfs.Count, null);
            w.WriteEndElement();
            w.WriteEndDocument();
        }

        /// <summary>Copies a collection element, fixing its count and appending <paramref name="extra"/> new items.</summary>
        private void MergeCollection(XmlReader r, XmlWriter w, int extra, Action writeExtra)
        {
            var name = r.LocalName;
            var total = (_templateCounts.TryGetValue(name, out var n) ? n : 0) + extra;
            XmlCopy.AttributeRewriter fixCount = (element, attribute, value) =>
                element == name && attribute == "count" ? total.ToString(CultureInfo.InvariantCulture) : value;
            var depth = r.Depth;
            var empty = r.IsEmptyElement;
            // Write the start tag manually so an empty template collection can still receive items.
            w.WriteStartElement(r.Prefix, r.LocalName, r.NamespaceURI);
            var hadCount = false;
            if (r.MoveToFirstAttribute())
            {
                do
                {
                    if (r.LocalName == "count") { hadCount = true; w.WriteAttributeString("count", total.ToString(CultureInfo.InvariantCulture)); }
                    else w.WriteAttributeString(r.Prefix, r.LocalName, r.NamespaceURI, r.Value);
                }
                while (r.MoveToNextAttribute());
                r.MoveToElement();
            }
            if (!hadCount) w.WriteAttributeString("count", total.ToString(CultureInfo.InvariantCulture));
            r.Read();
            if (!empty)
            {
                while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
                {
                    if (r.NodeType == XmlNodeType.Element) XmlCopy.CopySubtree(r, w, fixCount);
                    else { XmlCopy.CopyNode(r, w); r.Read(); }
                }
                r.Read(); // past the collection's end tag
            }
            writeExtra();
            w.WriteEndElement();
        }

        private void WriteNumberFormats(XmlWriter w, int count, int? templateCount)
        {
            w.WriteStartElement("numFmts", OpenXmlNamespaces.Main);
            WriteCount(w, count + (templateCount ?? 0));
            WriteNumberFormatItems(w);
            w.WriteEndElement();
        }

        private void WriteNumberFormatItems(XmlWriter w)
        {
            foreach (var pair in _customNumberFormats)
            {
                w.WriteStartElement("numFmt", OpenXmlNamespaces.Main);
                w.WriteAttributeString("numFmtId", pair.Key.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("formatCode", XmlTextUtil.Sanitize(pair.Value));
                w.WriteEndElement();
            }
        }

        private void WriteDifferentials(XmlWriter w, int count, int? templateCount)
        {
            w.WriteStartElement("dxfs", OpenXmlNamespaces.Main);
            WriteCount(w, count + (templateCount ?? 0));
            foreach (var d in _dxfs) d.WriteTo(w);
            w.WriteEndElement();
        }

        /// <summary>
        /// Writes the complete <c>styles.xml</c> part. Element order follows CT_Stylesheet exactly:
        /// numFmts, fonts, fills, borders, cellStyleXfs, cellXfs, cellStyles, dxfs. Excel rejects any other order.
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

            // Master cell styles: "Normal" (xfId 0) always, plus the built-in "Hyperlink" (xfId 1) when used.
            var hasHyperlink = _hyperlinkFontId >= 0 && !_appendMode;
            w.WriteStartElement("cellStyleXfs");
            WriteCount(w, hasHyperlink ? 2 : 1);
            w.WriteStartElement("xf");
            w.WriteAttributeString("numFmtId", "0");
            w.WriteAttributeString("fontId", "0");
            w.WriteAttributeString("fillId", "0");
            w.WriteAttributeString("borderId", "0");
            w.WriteEndElement();
            if (hasHyperlink)
            {
                w.WriteStartElement("xf");
                w.WriteAttributeString("numFmtId", "0");
                w.WriteAttributeString("fontId", _hyperlinkFontId.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("fillId", "0");
                w.WriteAttributeString("borderId", "0");
                w.WriteAttributeString("applyNumberFormat", "0");
                w.WriteAttributeString("applyFill", "0");
                w.WriteAttributeString("applyBorder", "0");
                w.WriteAttributeString("applyAlignment", "0");
                w.WriteAttributeString("applyProtection", "0");
                w.WriteEndElement();
            }
            w.WriteEndElement();

            w.WriteStartElement("cellXfs");
            WriteCount(w, _xfs.Count);
            foreach (var xf in _xfs) xf.WriteTo(w);
            w.WriteEndElement();

            w.WriteStartElement("cellStyles");
            WriteCount(w, hasHyperlink ? 2 : 1);
            if (hasHyperlink)
            {
                w.WriteStartElement("cellStyle");
                w.WriteAttributeString("name", "Hyperlink");
                w.WriteAttributeString("xfId", "1");
                w.WriteAttributeString("builtinId", "8");
                w.WriteEndElement();
            }
            w.WriteStartElement("cellStyle");
            w.WriteAttributeString("name", "Normal");
            w.WriteAttributeString("xfId", "0");
            w.WriteAttributeString("builtinId", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            // dxfs follow cellStyles in CT_Stylesheet; conditional formatting rules point here by index.
            if (_dxfs.Count > 0)
            {
                w.WriteStartElement("dxfs");
                WriteCount(w, _dxfs.Count);
                foreach (var dxf in _dxfs) dxf.WriteTo(w);
                w.WriteEndElement();
            }

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
            public static readonly XfSpec Default = new XfSpec(0, 0, 0, 0, HorizontalAlignment.General, VerticalAlignment.Bottom, false, false, 0, 0, true, false, 0);

            public readonly int NumFmtId, FontId, FillId, BorderId;
            public readonly HorizontalAlignment HAlign;
            public readonly VerticalAlignment VAlign;
            public readonly bool Wrap, Shrink;
            public readonly int Indent, Rotation;
            public readonly bool Locked, HideFormula;
            /// <summary>Index into cellStyleXfs: 0 = Normal, 1 = Hyperlink.</summary>
            public readonly int ParentXf;

            public XfSpec(int numFmtId, int fontId, int fillId, int borderId, HorizontalAlignment hAlign, VerticalAlignment vAlign, bool wrap,
                bool shrink, int indent, int rotation, bool locked, bool hideFormula, int parentXf)
            {
                NumFmtId = numFmtId; FontId = fontId; FillId = fillId; BorderId = borderId;
                HAlign = hAlign; VAlign = vAlign; Wrap = wrap; Shrink = shrink; Indent = indent; Rotation = rotation;
                Locked = locked; HideFormula = hideFormula; ParentXf = parentXf;
            }

            public XfSpec WithNumberFormat(int numFmtId)
                => new XfSpec(numFmtId, FontId, FillId, BorderId, HAlign, VAlign, Wrap, Shrink, Indent, Rotation, Locked, HideFormula, ParentXf);

            /// <summary>Reads the attributes of a template <c>xf</c> element (children are applied separately).</summary>
            public static XfSpec FromTemplate(XmlReader r)
                => new XfSpec(TemplateSource.Int(r.GetAttribute("numFmtId")) ?? 0, TemplateSource.Int(r.GetAttribute("fontId")) ?? 0,
                    TemplateSource.Int(r.GetAttribute("fillId")) ?? 0, TemplateSource.Int(r.GetAttribute("borderId")) ?? 0,
                    HorizontalAlignment.General, VerticalAlignment.Bottom, false, false, 0, 0, true, false,
                    TemplateSource.Int(r.GetAttribute("xfId")) ?? 0);

            /// <summary>Applies a template <c>alignment</c> element.</summary>
            public XfSpec WithTemplateAlignment(XmlReader r)
            {
                var h = r.GetAttribute("horizontal") switch
                {
                    "left" => HorizontalAlignment.Left, "center" => HorizontalAlignment.Center, "right" => HorizontalAlignment.Right,
                    "fill" => HorizontalAlignment.Fill, "justify" => HorizontalAlignment.Justify, _ => HorizontalAlignment.General,
                };
                var v = r.GetAttribute("vertical") switch
                {
                    "top" => VerticalAlignment.Top, "center" => VerticalAlignment.Center, "justify" => VerticalAlignment.Justify, _ => VerticalAlignment.Bottom,
                };
                return new XfSpec(NumFmtId, FontId, FillId, BorderId, h, v, Flag(r.GetAttribute("wrapText")), Flag(r.GetAttribute("shrinkToFit")),
                    TemplateSource.Int(r.GetAttribute("indent")) ?? 0, TemplateSource.Int(r.GetAttribute("textRotation")) ?? 0, Locked, HideFormula, ParentXf);
            }

            /// <summary>Applies a template <c>protection</c> element.</summary>
            public XfSpec WithTemplateProtection(XmlReader r)
            {
                var locked = r.GetAttribute("locked");
                return new XfSpec(NumFmtId, FontId, FillId, BorderId, HAlign, VAlign, Wrap, Shrink, Indent, Rotation,
                    locked == null || Flag(locked), Flag(r.GetAttribute("hidden")), ParentXf);
            }

            private static bool Flag(string? v) => v == "1" || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase);

            private ((int, int, int, int, int), (HorizontalAlignment, VerticalAlignment, bool, bool, int, int), (bool, bool)) Key
                => ((NumFmtId, FontId, FillId, BorderId, ParentXf), (HAlign, VAlign, Wrap, Shrink, Indent, Rotation), (Locked, HideFormula));
            public bool Equals(XfSpec other) => Key.Equals(other.Key);
            public override bool Equals(object? obj) => obj is XfSpec other && Equals(other);
            public override int GetHashCode() => Key.GetHashCode();

            public void WriteTo(XmlWriter w)
            {
                var hasAlignment = HAlign != HorizontalAlignment.General || VAlign != VerticalAlignment.Bottom || Wrap || Shrink || Indent != 0 || Rotation != 0;
                var hasProtection = !Locked || HideFormula;

                w.WriteStartElement("xf");
                w.WriteAttributeString("numFmtId", NumFmtId.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("fontId", FontId.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("fillId", FillId.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("borderId", BorderId.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("xfId", ParentXf.ToString(CultureInfo.InvariantCulture));
                // The apply* flags tell Excel the xf overrides the parent cell style; without them
                // a fill or font can be silently ignored.
                if (NumFmtId != 0) w.WriteAttributeString("applyNumberFormat", "1");
                if (FontId != 0) w.WriteAttributeString("applyFont", "1");
                if (FillId != 0) w.WriteAttributeString("applyFill", "1");
                if (BorderId != 0) w.WriteAttributeString("applyBorder", "1");
                if (hasAlignment) w.WriteAttributeString("applyAlignment", "1");
                if (hasProtection) w.WriteAttributeString("applyProtection", "1");
                // CT_Xf children: alignment, then protection.
                if (hasAlignment)
                {
                    w.WriteStartElement("alignment");
                    if (HAlign != HorizontalAlignment.General) w.WriteAttributeString("horizontal", ToXml(HAlign));
                    if (VAlign != VerticalAlignment.Bottom) w.WriteAttributeString("vertical", ToXml(VAlign));
                    if (Rotation != 0) w.WriteAttributeString("textRotation", Rotation.ToString(CultureInfo.InvariantCulture));
                    if (Wrap) w.WriteAttributeString("wrapText", "1");
                    if (Indent != 0) w.WriteAttributeString("indent", Indent.ToString(CultureInfo.InvariantCulture));
                    if (Shrink) w.WriteAttributeString("shrinkToFit", "1");
                    w.WriteEndElement();
                }
                if (hasProtection)
                {
                    w.WriteStartElement("protection");
                    if (!Locked) w.WriteAttributeString("locked", "0");
                    if (HideFormula) w.WriteAttributeString("hidden", "1");
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }
        }

        /// <summary>Immutable, hashable differential format (one <c>dxfs/dxf</c> entry).</summary>
        private readonly struct DxfSpec : IEquatable<DxfSpec>
        {
            public readonly bool Bold, Italic, Underline, Strike;
            public readonly string? FontColor, FillColor;
            public readonly BorderSpec Border;
            public readonly int NumFmtId;
            public readonly string? NumFmtCode;

            public DxfSpec(bool bold, bool italic, bool underline, bool strike, string? fontColor, string? fillColor, BorderSpec border, int numFmtId, string? numFmtCode)
            {
                Bold = bold; Italic = italic; Underline = underline; Strike = strike;
                FontColor = fontColor; FillColor = fillColor; Border = border; NumFmtId = numFmtId; NumFmtCode = numFmtCode;
            }

            private (bool, bool, bool, bool, string?, string?, BorderSpec, string?) Key => (Bold, Italic, Underline, Strike, FontColor, FillColor, Border, NumFmtCode);
            public bool Equals(DxfSpec other) => Key.Equals(other.Key);
            public override bool Equals(object? obj) => obj is DxfSpec other && Equals(other);
            public override int GetHashCode() => Key.GetHashCode();

            /// <summary>CT_Dxf order: font, numFmt, fill, alignment, protection, border.</summary>
            public void WriteTo(XmlWriter w)
            {
                w.WriteStartElement("dxf");
                if (Bold || Italic || Underline || Strike || FontColor != null)
                {
                    w.WriteStartElement("font");
                    if (Bold) { w.WriteStartElement("b"); w.WriteEndElement(); }
                    if (Italic) { w.WriteStartElement("i"); w.WriteEndElement(); }
                    if (Strike) { w.WriteStartElement("strike"); w.WriteEndElement(); }
                    if (FontColor != null) { w.WriteStartElement("color"); w.WriteAttributeString("rgb", FontColor); w.WriteEndElement(); }
                    if (Underline) { w.WriteStartElement("u"); w.WriteEndElement(); }
                    w.WriteEndElement();
                }
                if (NumFmtCode != null)
                {
                    w.WriteStartElement("numFmt");
                    w.WriteAttributeString("numFmtId", NumFmtId.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("formatCode", XmlTextUtil.Sanitize(NumFmtCode));
                    w.WriteEndElement();
                }
                if (FillColor != null)
                {
                    // Unlike cellXfs fills, a differential solid fill carries its colour in bgColor.
                    w.WriteStartElement("fill");
                    w.WriteStartElement("patternFill");
                    w.WriteStartElement("bgColor");
                    w.WriteAttributeString("rgb", FillColor);
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                }
                if (!Border.Equals(BorderSpec.Empty)) Border.WriteTo(w);
                w.WriteEndElement();
            }
        }
    }
}
