using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using SheetStream.Internal;
using SheetStream.Styles;

namespace SheetStream
{
    /// <summary>Comparison of a "cell value" conditional formatting rule.</summary>
    public enum ConditionalOperator
    {
        /// <summary>Between two values (inclusive).</summary>
        Between,
        /// <summary>Outside two values.</summary>
        NotBetween,
        /// <summary>Equal to the value.</summary>
        Equal,
        /// <summary>Not equal to the value.</summary>
        NotEqual,
        /// <summary>Greater than the value.</summary>
        GreaterThan,
        /// <summary>Less than the value.</summary>
        LessThan,
        /// <summary>Greater than or equal to the value.</summary>
        GreaterThanOrEqual,
        /// <summary>Less than or equal to the value.</summary>
        LessThanOrEqual,
    }

    /// <summary>Built-in icon sets for <see cref="ConditionalFormatting.IconSet"/>.</summary>
    public enum IconSetType
    {
        /// <summary>Three coloured arrows.</summary>
        ThreeArrows,
        /// <summary>Three grey arrows.</summary>
        ThreeArrowsGray,
        /// <summary>Three flags.</summary>
        ThreeFlags,
        /// <summary>Three traffic lights (Excel's default).</summary>
        ThreeTrafficLights,
        /// <summary>Three road signs.</summary>
        ThreeSigns,
        /// <summary>Three circled symbols.</summary>
        ThreeSymbols,
        /// <summary>Four coloured arrows.</summary>
        FourArrows,
        /// <summary>Four rating bars.</summary>
        FourRating,
        /// <summary>Four traffic lights.</summary>
        FourTrafficLights,
        /// <summary>Five coloured arrows.</summary>
        FiveArrows,
        /// <summary>Five rating bars.</summary>
        FiveRating,
        /// <summary>Five quarters.</summary>
        FiveQuarters,
    }

    /// <summary>
    /// Conditional formatting rules over one range. Create through
    /// <see cref="Worksheet.AddConditionalFormatting(string)"/> and chain any number of rules; earlier rules win
    /// when several match (Excel priority order).
    /// </summary>
    /// <example>
    /// <code>
    /// var bad = workbook.Styles.CreateDifferential(o => { o.Font.Color = "#9C0006"; o.Fill.BackgroundColor = "#FFC7CE"; });
    /// sheet.AddConditionalFormatting("C2:C101")
    ///      .CellValue(ConditionalOperator.LessThan, "0", bad)
    ///      .DataBar("#638EC6");
    /// </code>
    /// </example>
    public sealed class ConditionalFormatting
    {
        private readonly Worksheet _sheet;
        private readonly List<Action<XmlWriter, int>> _rules = new List<Action<XmlWriter, int>>();

        internal ConditionalFormatting(Worksheet sheet, RangeReference range)
        {
            _sheet = sheet;
            Range = range;
        }

        internal RangeReference Range { get; }

        internal int RuleCount => _rules.Count;

        /// <summary>Formats cells whose value compares to one or two values or formulas (written without "=").</summary>
        public ConditionalFormatting CellValue(ConditionalOperator op, string value1, DifferentialStyle style, string? value2 = null)
        {
            if (value1 == null) throw new ArgumentNullException(nameof(value1));
            Require(style);
            var needsTwo = op == ConditionalOperator.Between || op == ConditionalOperator.NotBetween;
            if (needsTwo && value2 == null) throw new ArgumentException("Between and NotBetween need a second value.", nameof(value2));
            return Add((w, p) =>
            {
                Start(w, "cellIs", p, style);
                w.WriteAttributeString("operator", ToXml(op));
                Formula(w, value1);
                if (needsTwo) Formula(w, value2!);
                w.WriteEndElement();
            });
        }

        /// <summary>
        /// Formats cells where the formula is TRUE. Write it relative to the range's top-left cell, e.g.
        /// <c>"$C2&gt;$D2"</c> for a range starting in row 2, to colour whole rows.
        /// </summary>
        public ConditionalFormatting Expression(string formula, DifferentialStyle style)
        {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentException("Formula cannot be empty.", nameof(formula));
            Require(style);
            return Add((w, p) =>
            {
                Start(w, "expression", p, style);
                Formula(w, formula);
                w.WriteEndElement();
            });
        }

        /// <summary>Formats cells whose text contains <paramref name="text"/> (case-insensitive, like Excel).</summary>
        public ConditionalFormatting ContainsText(string text, DifferentialStyle style)
        {
            if (string.IsNullOrEmpty(text)) throw new ArgumentException("Text cannot be empty.", nameof(text));
            Require(style);
            var topLeft = CellReference.ToA1(Range.FromRow, Range.FromColumn);
            return Add((w, p) =>
            {
                Start(w, "containsText", p, style);
                w.WriteAttributeString("operator", "containsText");
                w.WriteAttributeString("text", XmlTextUtil.Sanitize(text));
                // Excel evaluates the stored formula; the text attribute only drives the UI.
                Formula(w, "NOT(ISERROR(SEARCH(\"" + text.Replace("\"", "\"\"") + "\"," + topLeft + ")))");
                w.WriteEndElement();
            });
        }

        /// <summary>Formats values that appear more than once in the range.</summary>
        public ConditionalFormatting DuplicateValues(DifferentialStyle style)
        {
            Require(style);
            return Add((w, p) => { Start(w, "duplicateValues", p, style); w.WriteEndElement(); });
        }

        /// <summary>Formats values that appear exactly once in the range.</summary>
        public ConditionalFormatting UniqueValues(DifferentialStyle style)
        {
            Require(style);
            return Add((w, p) => { Start(w, "uniqueValues", p, style); w.WriteEndElement(); });
        }

        /// <summary>Formats the top <paramref name="rank"/> values (or top percent).</summary>
        public ConditionalFormatting Top(int rank, DifferentialStyle style, bool percent = false) => TopBottom(rank, style, percent, false);

        /// <summary>Formats the bottom <paramref name="rank"/> values (or bottom percent).</summary>
        public ConditionalFormatting Bottom(int rank, DifferentialStyle style, bool percent = false) => TopBottom(rank, style, percent, true);

        /// <summary>Formats values above the range average.</summary>
        public ConditionalFormatting AboveAverage(DifferentialStyle style) => Average(style, true);

        /// <summary>Formats values below the range average.</summary>
        public ConditionalFormatting BelowAverage(DifferentialStyle style) => Average(style, false);

        /// <summary>Two-colour scale from the lowest to the highest value.</summary>
        public ConditionalFormatting ColorScale(string minColor, string maxColor)
        {
            var min = XmlTextUtil.NormalizeColor(minColor, nameof(minColor));
            var max = XmlTextUtil.NormalizeColor(maxColor, nameof(maxColor));
            return Add((w, p) =>
            {
                Start(w, "colorScale", p, null);
                w.WriteStartElement("colorScale");
                Cfvo(w, "min", null);
                Cfvo(w, "max", null);
                Color(w, min);
                Color(w, max);
                w.WriteEndElement();
                w.WriteEndElement();
            });
        }

        /// <summary>Three-colour scale: lowest, 50th percentile, highest (e.g. red, yellow, green).</summary>
        public ConditionalFormatting ColorScale(string minColor, string midColor, string maxColor)
        {
            var min = XmlTextUtil.NormalizeColor(minColor, nameof(minColor));
            var mid = XmlTextUtil.NormalizeColor(midColor, nameof(midColor));
            var max = XmlTextUtil.NormalizeColor(maxColor, nameof(maxColor));
            return Add((w, p) =>
            {
                Start(w, "colorScale", p, null);
                w.WriteStartElement("colorScale");
                Cfvo(w, "min", null);
                Cfvo(w, "percentile", "50");
                Cfvo(w, "max", null);
                Color(w, min);
                Color(w, mid);
                Color(w, max);
                w.WriteEndElement();
                w.WriteEndElement();
            });
        }

        /// <summary>In-cell data bars scaled from the lowest to the highest value.</summary>
        public ConditionalFormatting DataBar(string color)
        {
            var argb = XmlTextUtil.NormalizeColor(color, nameof(color));
            return Add((w, p) =>
            {
                Start(w, "dataBar", p, null);
                w.WriteStartElement("dataBar");
                Cfvo(w, "min", null);
                Cfvo(w, "max", null);
                Color(w, argb);
                w.WriteEndElement();
                w.WriteEndElement();
            });
        }

        /// <summary>Icon set with thresholds at equal percentage steps.</summary>
        public ConditionalFormatting IconSet(IconSetType type, bool reverse = false, bool showValue = true)
        {
            var (name, count) = IconSetName(type);
            return Add((w, p) =>
            {
                Start(w, "iconSet", p, null);
                w.WriteStartElement("iconSet");
                w.WriteAttributeString("iconSet", name);
                if (!showValue) w.WriteAttributeString("showValue", "0");
                if (reverse) w.WriteAttributeString("reverse", "1");
                for (var i = 0; i < count; i++)
                    Cfvo(w, "percent", (100 * i / count).ToString(CultureInfo.InvariantCulture));
                w.WriteEndElement();
                w.WriteEndElement();
            });
        }

        private ConditionalFormatting TopBottom(int rank, DifferentialStyle style, bool percent, bool bottom)
        {
            if (rank < 1 || (percent && rank > 100)) throw new ArgumentOutOfRangeException(nameof(rank), rank, "Rank must be at least 1 (and at most 100 for percentages).");
            Require(style);
            return Add((w, p) =>
            {
                Start(w, "top10", p, style);
                if (percent) w.WriteAttributeString("percent", "1");
                if (bottom) w.WriteAttributeString("bottom", "1");
                w.WriteAttributeString("rank", rank.ToString(CultureInfo.InvariantCulture));
                w.WriteEndElement();
            });
        }

        private ConditionalFormatting Average(DifferentialStyle style, bool above)
        {
            Require(style);
            return Add((w, p) =>
            {
                Start(w, "aboveAverage", p, style);
                if (!above) w.WriteAttributeString("aboveAverage", "0");
                w.WriteEndElement();
            });
        }

        private ConditionalFormatting Add(Action<XmlWriter, int> rule)
        {
            _sheet.EnsureWritable();
            _rules.Add(rule);
            return this;
        }

        private static void Require(DifferentialStyle style)
        {
            if (style == null) throw new ArgumentNullException(nameof(style), "Create one with workbook.Styles.CreateDifferential(...).");
        }

        private static void Start(XmlWriter w, string type, int priority, DifferentialStyle? style)
        {
            w.WriteStartElement("cfRule");
            w.WriteAttributeString("type", type);
            if (style != null) w.WriteAttributeString("dxfId", style.Id.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("priority", priority.ToString(CultureInfo.InvariantCulture));
        }

        private static void Formula(XmlWriter w, string formula)
            => w.WriteElementString("formula", XmlTextUtil.Sanitize(formula.Trim().TrimStart('=')));

        private static void Cfvo(XmlWriter w, string type, string? val)
        {
            w.WriteStartElement("cfvo");
            w.WriteAttributeString("type", type);
            if (val != null) w.WriteAttributeString("val", val);
            w.WriteEndElement();
        }

        private static void Color(XmlWriter w, string argb)
        {
            w.WriteStartElement("color");
            w.WriteAttributeString("rgb", argb);
            w.WriteEndElement();
        }

        private static string ToXml(ConditionalOperator op) => op switch
        {
            ConditionalOperator.Between => "between",
            ConditionalOperator.NotBetween => "notBetween",
            ConditionalOperator.Equal => "equal",
            ConditionalOperator.NotEqual => "notEqual",
            ConditionalOperator.GreaterThan => "greaterThan",
            ConditionalOperator.LessThan => "lessThan",
            ConditionalOperator.GreaterThanOrEqual => "greaterThanOrEqual",
            _ => "lessThanOrEqual",
        };

        private static (string Name, int Count) IconSetName(IconSetType type) => type switch
        {
            IconSetType.ThreeArrows => ("3Arrows", 3),
            IconSetType.ThreeArrowsGray => ("3ArrowsGray", 3),
            IconSetType.ThreeFlags => ("3Flags", 3),
            IconSetType.ThreeSigns => ("3Signs", 3),
            IconSetType.ThreeSymbols => ("3Symbols", 3),
            IconSetType.FourArrows => ("4Arrows", 4),
            IconSetType.FourRating => ("4Rating", 4),
            IconSetType.FourTrafficLights => ("4TrafficLights", 4),
            IconSetType.FiveArrows => ("5Arrows", 5),
            IconSetType.FiveRating => ("5Rating", 5),
            IconSetType.FiveQuarters => ("5Quarters", 5),
            _ => ("3TrafficLights1", 3),
        };

        /// <summary>
        /// Writes the <c>conditionalFormatting</c> element. Priorities are unique per sheet, so the caller passes
        /// the next free priority and receives the one after the last rule written.
        /// </summary>
        internal int WriteTo(XmlWriter w, int nextPriority)
        {
            if (_rules.Count == 0) return nextPriority;
            w.WriteStartElement("conditionalFormatting");
            w.WriteAttributeString("sqref", Range.ToString());
            foreach (var rule in _rules) rule(w, nextPriority++);
            w.WriteEndElement();
            return nextPriority;
        }
    }
}
