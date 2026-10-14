using System;
using System.Globalization;
using System.Text;

namespace SheetStream.Internal
{
    /// <summary>
    /// Approximates the text Excel displays for a number under a number format, so AutoFit can measure what the
    /// user will see rather than the raw value. One instance is parsed per number format id and cached by the
    /// <see cref="Styles.StyleSheet"/>; rendering a value then costs one <c>ToString</c>, or nothing for date
    /// and time formats, whose display width does not depend on the value.
    /// </summary>
    /// <remarks>
    /// Only the first section of the format code (positive numbers) is read. Literal text, escaped characters,
    /// <c>_x</c> padding (Excel reserves the width of <c>x</c>), currency in <c>[$€-407]</c> blocks, percent
    /// scaling, thousands separators, fixed decimals and scientific notation are honoured; colours, conditions
    /// and <c>*</c> fill characters are ignored. Exact rendering would need Excel's formatter and locale.
    /// </remarks>
    internal sealed class NumberDisplayFormat
    {
        /// <summary>Excel's General format: up to 11 characters, scientific beyond.</summary>
        public static readonly NumberDisplayFormat General = new NumberDisplayFormat(null, true, "", "", 1);

        // Placeholders in a date template for the parts whose width depends on the value: month and day names
        // (Excel shows them in English for the en-US locale assumed throughout) and elapsed hours/minutes/seconds,
        // which grow with the value. Control characters never occur in a format code, so they cannot collide.
        private const char MonthName = '\u0001';
        private const char MonthAbbreviation = '\u0002';
        private const char DayName = '\u0003';
        private const char DayAbbreviation = '\u0004';
        private const char Elapsed = '\u0005';

        private readonly string? _dateTemplate;
        private readonly bool _dateHasPlaceholders;
        private readonly int _elapsedDigits;
        private readonly double _elapsedPerDay;
        private readonly string? _netFormat;
        private readonly bool _general;
        private readonly string _prefix;
        private readonly string _suffix;
        private readonly double _scale;

        private NumberDisplayFormat(string? netFormat, bool general, string prefix, string suffix, double scale)
        {
            _netFormat = netFormat;
            _general = general;
            _prefix = prefix;
            _suffix = suffix;
            _scale = scale;
        }

        private NumberDisplayFormat(string dateTemplate, int elapsedDigits, double elapsedPerDay)
        {
            _dateTemplate = dateTemplate;
            _dateHasPlaceholders = dateTemplate.IndexOfAny(new[] { MonthName, MonthAbbreviation, DayName, DayAbbreviation, Elapsed }) >= 0;
            _elapsedDigits = elapsedDigits;
            _elapsedPerDay = elapsedPerDay;
            _prefix = _suffix = "";
            _scale = 1;
        }

        /// <summary>The text a cell holding <paramref name="value"/> displays under this format, approximately.</summary>
        public string Render(double value)
        {
            if (_dateTemplate != null) return _dateHasPlaceholders ? ExpandDate(value) : _dateTemplate;
            var digits = _netFormat == null || _general
                ? GeneralText(value * _scale)
                : (value * _scale).ToString(_netFormat, CultureInfo.InvariantCulture);
            if (_prefix.Length == 0 && _suffix.Length == 0) return digits;
            return _prefix + digits + _suffix;
        }

        /// <summary>Fills the month/day names and elapsed digits of a date template from the serial date value.</summary>
        private string ExpandDate(double value)
        {
            // Out-of-range serials (Excel shows ####) fall back to the widest common names.
            var inRange = value >= -657435 && value < 2958466;
            var date = inRange ? DateTime.FromOADate(value) : new DateTime(2000, 9, 6); // A Wednesday in September.
            var sb = new StringBuilder(_dateTemplate!.Length + 16);
            foreach (var c in _dateTemplate)
            {
                switch (c)
                {
                    case MonthName: sb.Append(CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(date.Month)); break;
                    case MonthAbbreviation: sb.Append(CultureInfo.InvariantCulture.DateTimeFormat.GetAbbreviatedMonthName(date.Month)); break;
                    case DayName: sb.Append(CultureInfo.InvariantCulture.DateTimeFormat.GetDayName(date.DayOfWeek)); break;
                    case DayAbbreviation: sb.Append(CultureInfo.InvariantCulture.DateTimeFormat.GetAbbreviatedDayName(date.DayOfWeek)); break;
                    case Elapsed:
                    {
                        var total = Math.Floor(Math.Abs(value) * _elapsedPerDay);
                        var digits = total < 1 ? 1 : (int)Math.Floor(Math.Log10(total)) + 1;
                        sb.Append('0', Math.Max(digits, _elapsedDigits));
                        break;
                    }
                    default: sb.Append(c); break;
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// What Excel's General format shows in a column wide enough for the value: whole numbers in full up to
        /// 11 digits, decimals to 10 significant digits, and six-significant-digit scientific notation beyond
        /// (<c>1.23457E+12</c>), which is also how Excel abbreviates very small values.
        /// </summary>
        public static string GeneralText(double value)
        {
            var abs = Math.Abs(value);
            if (abs == 0) return "0";
            if (abs >= 1e11 || abs < 1e-4) return value.ToString("0.#####E+00", CultureInfo.InvariantCulture);
            if (value == Math.Floor(value)) return value.ToString("0", CultureInfo.InvariantCulture);
            return value.ToString("G10", CultureInfo.InvariantCulture);
        }

        /// <summary>Parses an Excel number format code. <c>null</c>, empty and <c>General</c> give <see cref="General"/>.</summary>
        public static NumberDisplayFormat Parse(string? code)
        {
            if (string.IsNullOrEmpty(code) || code == "General") return General;
            var s = FirstSection(code!);
            if (s.Length == 0) return General;

            // Literals are collected twice: in display order for date formats (whose text is value-independent)
            // and split into prefix / suffix around the digits for numeric formats.
            var dateText = new StringBuilder();
            var prefix = new StringBuilder();
            var suffix = new StringBuilder();
            var isDate = false;
            var isText = false;
            var general = false;
            var hasPlaceholder = false;
            var thousands = false;
            var percent = false;
            var scientific = false;
            var inDecimals = false;
            var decimals = 0;
            var elapsedDigits = 0;
            var elapsedPerDay = 0.0;

            void Literal(string text)
            {
                dateText.Append(text);
                (hasPlaceholder ? suffix : prefix).Append(text);
            }

            for (var i = 0; i < s.Length; i++)
            {
                var c = s[i];
                switch (c)
                {
                    case '"':
                    {
                        var end = s.IndexOf('"', i + 1);
                        if (end < 0) end = s.Length;
                        Literal(s.Substring(i + 1, end - i - 1));
                        i = end;
                        break;
                    }
                    case '\\': // Escaped literal character.
                    case '_':  // Padding: Excel leaves the width of the next character blank.
                        if (i + 1 < s.Length) Literal(s[++i].ToString());
                        break;
                    case '*': // Repeat-to-fill: no intrinsic width.
                        if (i + 1 < s.Length) i++;
                        break;
                    case '[':
                    {
                        var close = s.IndexOf(']', i);
                        if (close < 0) close = s.Length - 1;
                        var inner = s.Substring(i + 1, close - i - 1);
                        if (inner.Length > 0 && inner[0] == '$')
                        {
                            // [$€-407]: a currency symbol with a locale tag; [$-409] is a bare locale tag.
                            var dash = inner.IndexOf('-');
                            Literal(dash < 0 ? inner.Substring(1) : inner.Substring(1, dash - 1));
                        }
                        else if (inner.Length > 0 && IsElapsedToken(inner))
                        {
                            // [h], [mm], [ss]: elapsed time that can exceed its unit; the digit count follows the value.
                            isDate = true;
                            elapsedDigits = inner.Length;
                            var unit = char.ToLowerInvariant(inner[0]);
                            elapsedPerDay = unit == 'h' ? 24 : unit == 'm' ? 1440 : 86400;
                            dateText.Append(Elapsed);
                        }
                        // Colours ([Red]) and conditions ([>=100]) do not display.
                        i = close;
                        break;
                    }
                    case '0':
                    case '#':
                    case '?':
                        if (isDate) dateText.Append('0'); // Fractional seconds: ss.00
                        else
                        {
                            hasPlaceholder = true;
                            if (inDecimals) decimals++;
                        }
                        break;
                    case '.':
                        if (isDate) dateText.Append('.');
                        else if (!inDecimals && !scientific) inDecimals = true;
                        break;
                    case ',':
                        if (isDate) dateText.Append(',');
                        else if (hasPlaceholder && !inDecimals) thousands = true; // Trailing commas (scaling) are rare; ignored.
                        break;
                    case '%':
                        percent = true;
                        Literal("%");
                        break;
                    case 'E':
                    case 'e':
                        if (!isDate && i + 1 < s.Length && (s[i + 1] == '+' || s[i + 1] == '-'))
                        {
                            scientific = true;
                            inDecimals = false;
                            i++;
                            // The exponent digits are placeholders too; skip them so they do not count as decimals.
                            while (i + 1 < s.Length && s[i + 1] == '0') i++;
                        }
                        else Literal(c.ToString());
                        break;
                    case '@':
                        isText = true;
                        break;
                    case 'y': case 'Y':
                    case 'm': case 'M':
                    case 'd': case 'D':
                    case 'h': case 'H':
                    case 's': case 'S':
                    {
                        isDate = true;
                        var run = 1;
                        while (i + run < s.Length && char.ToLowerInvariant(s[i + run]) == char.ToLowerInvariant(c)) run++;
                        dateText.Append(DateToken(char.ToLowerInvariant(c), run));
                        i += run - 1;
                        break;
                    }
                    case 'A':
                    case 'a':
                        if (StartsWith(s, i, "AM/PM")) { isDate = true; dateText.Append("AM"); i += 4; }
                        else if (StartsWith(s, i, "A/P")) { isDate = true; dateText.Append('A'); i += 2; }
                        else Literal(c.ToString());
                        break;
                    case 'G':
                        if (StartsWith(s, i, "General")) { hasPlaceholder = true; general = true; i += 6; }
                        else Literal(c.ToString());
                        break;
                    default:
                        Literal(c.ToString());
                        break;
                }
            }

            if (isDate) return new NumberDisplayFormat(dateText.ToString(), elapsedDigits, elapsedPerDay);
            if (isText || !hasPlaceholder) return General;

            string netFormat;
            if (scientific) netFormat = decimals > 0 ? "0." + new string('0', decimals) + "E+00" : "0E+00";
            else
            {
                netFormat = thousands ? "#,##0" : "0";
                if (decimals > 0) netFormat += "." + new string('0', decimals);
            }
            return new NumberDisplayFormat(netFormat, general, prefix.ToString(), suffix.ToString(), percent ? 100 : 1);
        }

        /// <summary>The positive-number section: everything before the first <c>;</c> outside quotes and brackets.</summary>
        private static string FirstSection(string code)
        {
            var inQuote = false;
            var depth = 0;
            for (var i = 0; i < code.Length; i++)
            {
                var c = code[i];
                if (c == '"') inQuote = !inQuote;
                else if (inQuote) continue;
                else if (c == '\\' || c == '_' || c == '*') i++;
                else if (c == '[') depth++;
                else if (c == ']') depth--;
                else if (c == ';' && depth <= 0) return code.Substring(0, i);
            }
            return code;
        }

        /// <summary>
        /// Stand-in text for a date/time token: zeros for numeric parts (every digit is the same width) and a
        /// placeholder for month and day names, filled from the value when rendering.
        /// </summary>
        private static string DateToken(char token, int run)
        {
            switch (token)
            {
                case 'y': return run >= 3 ? "0000" : "00";
                case 'm': return run >= 5 ? "S" : run == 4 ? MonthName.ToString() : run == 3 ? MonthAbbreviation.ToString() : "00";
                case 'd': return run >= 4 ? DayName.ToString() : run == 3 ? DayAbbreviation.ToString() : "00";
                default: return "00"; // h, s (and minutes, which share 'm').
            }
        }

        /// <summary>True for <c>h</c>, <c>hh</c>, <c>mm</c>, <c>ss</c>... as the whole bracket content.</summary>
        private static bool IsElapsedToken(string inner)
        {
            var first = char.ToLowerInvariant(inner[0]);
            if (first != 'h' && first != 'm' && first != 's') return false;
            for (var i = 1; i < inner.Length; i++)
                if (char.ToLowerInvariant(inner[i]) != first) return false;
            return true;
        }

        private static bool StartsWith(string s, int index, string token)
            => string.Compare(s, index, token, 0, token.Length, StringComparison.OrdinalIgnoreCase) == 0
               && index + token.Length <= s.Length;
    }
}
