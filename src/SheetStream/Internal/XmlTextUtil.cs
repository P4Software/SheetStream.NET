using System;
using System.Globalization;
using System.Text;

namespace SheetStream.Internal
{
    /// <summary>
    /// Text helpers that keep user-supplied strings legal in XML 1.0 and faithful to Excel's conventions.
    /// </summary>
    internal static class XmlTextUtil
    {
        /// <summary>
        /// Makes a string safe to hand to <see cref="System.Xml.XmlWriter"/> and faithful to Excel's own
        /// escaping convention for characters XML cannot carry.
        /// </summary>
        /// <remarks>
        /// XML 1.0 forbids control characters other than tab, LF and CR. Database exports contain them
        /// surprisingly often (form feeds, NULs from fixed-width fields). Rather than throwing, or writing a
        /// file Excel then declares corrupt, we apply the same <c>_xHHHH_</c> escape Excel itself uses, so
        /// Excel decodes the original character on load. A literal <c>_xHHHH_</c> sequence typed by a user
        /// is protected by escaping its underscore as <c>_x005F_</c>, again mirroring Excel. Lone surrogates
        /// (invalid UTF-16) are replaced with U+FFFD because they cannot be encoded at all.
        /// The fast path returns the original instance when nothing needs changing, which is the common case.
        /// </remarks>
        public static string Sanitize(string value)
        {
            var needsWork = false;
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (IsForbidden(c) || char.IsSurrogate(c) || (c == '_' && LooksLikeEscape(value, i)))
                {
                    needsWork = true;
                    break;
                }
            }
            if (!needsWork) return value;

            var sb = new StringBuilder(value.Length + 16);
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (IsForbidden(c))
                {
                    AppendEscape(sb, c);
                }
                else if (char.IsHighSurrogate(c) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
                {
                    sb.Append(c).Append(value[i + 1]);
                    i++;
                }
                else if (char.IsSurrogate(c))
                {
                    sb.Append('�');
                }
                else if (c == '_' && LooksLikeEscape(value, i))
                {
                    sb.Append("_x005F_");
                }
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        /// <summary>Characters XML 1.0 cannot represent even when escaped as entities.</summary>
        private static bool IsForbidden(char c)
            => (c < 0x20 && c != '\t' && c != '\n' && c != '\r') || c == 0xFFFE || c == 0xFFFF;

        /// <summary>True when the text at <paramref name="i"/> reads <c>_xHHHH_</c> with four hex digits.</summary>
        private static bool LooksLikeEscape(string s, int i)
        {
            if (i + 6 >= s.Length) return false;
            if (s[i + 1] != 'x' && s[i + 1] != 'X') return false;
            for (var k = 2; k < 6; k++)
                if (!Uri.IsHexDigit(s[i + k])) return false;
            return s[i + 6] == '_';
        }

        private static void AppendEscape(StringBuilder sb, char c)
            => sb.Append("_x").Append(((int)c).ToString("X4", CultureInfo.InvariantCulture)).Append('_');

        /// <summary>True when a shared string needs <c>xml:space="preserve"</c> to survive a round trip.</summary>
        public static bool NeedsSpacePreserve(string value)
        {
            if (value.Length == 0) return false;
            return char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]);
        }

        /// <summary>
        /// Normalises a user colour ("#4F81BD", "4F81BD", "#FF4F81BD", "FF4F81BD") into the 8-digit
        /// <c>AARRGGBB</c> form SpreadsheetML expects. Six-digit input is treated as fully opaque.
        /// </summary>
        public static string NormalizeColor(string color, string parameterName)
        {
            if (color == null) throw new ArgumentNullException(parameterName);
            var hex = color.Trim();
            if (hex.StartsWith("#", StringComparison.Ordinal)) hex = hex.Substring(1);
            if (hex.Length != 6 && hex.Length != 8)
                throw new ArgumentException($"Colour '{color}' must be RRGGBB or AARRGGBB hex, optionally prefixed with '#'.", parameterName);
            for (var i = 0; i < hex.Length; i++)
                if (!Uri.IsHexDigit(hex[i]))
                    throw new ArgumentException($"Colour '{color}' contains a non-hex character.", parameterName);
            hex = hex.ToUpperInvariant();
            return hex.Length == 6 ? "FF" + hex : hex;
        }
    }
}
