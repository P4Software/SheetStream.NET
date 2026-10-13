using System;
using System.Globalization;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>
    /// One formatted run of text inside a rich-text cell. Unset properties inherit the cell's font.
    /// </summary>
    /// <example>
    /// <code>
    /// sheet.Row().WriteRichText(new RichText("Status: "), new RichText("LATE") { Bold = true, Color = "#C00000" });
    /// </code>
    /// </example>
    public sealed class RichText
    {
        private string? _color;

        /// <summary>Creates a run with the given text.</summary>
        public RichText(string text)
        {
            Text = text ?? throw new ArgumentNullException(nameof(text));
        }

        /// <summary>The run's text.</summary>
        public string Text { get; }

        /// <summary>Bold run.</summary>
        public bool Bold { get; set; }

        /// <summary>Italic run.</summary>
        public bool Italic { get; set; }

        /// <summary>Underlined run.</summary>
        public bool Underline { get; set; }

        /// <summary>Struck-through run.</summary>
        public bool Strikethrough { get; set; }

        /// <summary>Font size in points, or null to inherit.</summary>
        public double? Size { get; set; }

        /// <summary>Font name, or null to inherit.</summary>
        public string? FontName { get; set; }

        /// <summary>Text colour as hex, or null to inherit. Validated on assignment.</summary>
        public string? Color
        {
            get => _color;
            set => _color = value == null ? null : XmlTextUtil.NormalizeColor(value, nameof(Color));
        }

        /// <summary>Implicit conversion so plain strings can be mixed with formatted runs.</summary>
        public static implicit operator RichText(string text) => new RichText(text);

        /// <summary>Writes one <c>r</c> element (CT_RElt: rPr, t).</summary>
        internal void WriteTo(XmlWriter w)
        {
            w.WriteStartElement("r");
            if (Bold || Italic || Underline || Strikethrough || Size.HasValue || FontName != null || _color != null)
            {
                w.WriteStartElement("rPr");
                if (FontName != null) { w.WriteStartElement("rFont"); w.WriteAttributeString("val", FontName); w.WriteEndElement(); }
                if (Bold) { w.WriteStartElement("b"); w.WriteEndElement(); }
                if (Italic) { w.WriteStartElement("i"); w.WriteEndElement(); }
                if (Strikethrough) { w.WriteStartElement("strike"); w.WriteEndElement(); }
                if (_color != null) { w.WriteStartElement("color"); w.WriteAttributeString("rgb", _color); w.WriteEndElement(); }
                if (Size.HasValue) { w.WriteStartElement("sz"); w.WriteAttributeString("val", Size.Value.ToString(CultureInfo.InvariantCulture)); w.WriteEndElement(); }
                if (Underline) { w.WriteStartElement("u"); w.WriteEndElement(); }
                w.WriteEndElement();
            }
            w.WriteStartElement("t");
            var text = XmlTextUtil.Sanitize(Text);
            // Runs routinely start or end with a space ("Status: "), which Excel would trim without this.
            if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
                w.WriteAttributeString("xml", "space", null, "preserve");
            w.WriteString(text);
            w.WriteEndElement();
            w.WriteEndElement();
        }
    }
}
