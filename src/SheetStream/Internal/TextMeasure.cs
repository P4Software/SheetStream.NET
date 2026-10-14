using System;

namespace SheetStream.Internal
{
    /// <summary>
    /// Estimates how wide a text renders in Excel's default font (Calibri 11, the "Normal" style) without a
    /// graphics library. Widths are returned in Excel's column unit: the width of one digit of the default
    /// font, 7 px at 96 dpi, so <c>"0000000000"</c> measures exactly 10.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The tables hold the hinted advance width of every glyph in whole pixels at 11 pt / 96 dpi, which is how
    /// Excel itself lays cell text out (it rounds each glyph to the pixel grid; 'i' is 4 px, 'W' 13 px, digits
    /// 7 px). The values were calibrated by auto-fitting every printable ASCII character, regular and bold, in
    /// Excel 2016 and reading the resulting column widths back; the Latin-1 block reuses the width of the base
    /// letter, which is what Excel shows for accented text (é renders as wide as e).
    /// </para>
    /// <para>
    /// Calibri Bold at 11 pt is barely wider than regular on the pixel grid: only nine ASCII glyphs gain a
    /// pixel (" &amp; ' G M U W Y `), so bold is handled as per-glyph overrides instead of a blanket factor.
    /// Characters outside the tables are assumed to be one digit wide (two for CJK and other full-width
    /// scripts); applications with real font metrics plug them in through
    /// <see cref="PackageOptions.TextMeasurer"/> or <c>Worksheet.AutoFitColumns(measure)</c>.
    /// </para>
    /// </remarks>
    internal static class TextMeasure
    {
        /// <summary>Width of a digit in Calibri 11 at 96 dpi, Excel's "maximum digit width" and column unit.</summary>
        public const int DigitWidthPixels = 7;

        /// <summary>Width of a glyph outside the tables: one digit.</summary>
        private const int DefaultWidth = DigitWidthPixels;

        /// <summary>Width of a full-width (CJK, Hangul) glyph: two digits.</summary>
        private const int WideWidth = 2 * DigitWidthPixels;

        // Hinted Calibri 11 advance widths in pixels for U+0020..U+007E, as measured in Excel.
        private static readonly byte[] Ascii =
        {
            3,  // ' '
            5,  // !
            6,  // "
            7,  // #
            7,  // $
            11, // %
            10, // &
            3,  // '
            5,  // (
            5,  // )
            7,  // *
            7,  // +
            4,  // ,
            5,  // -
            4,  // .
            6,  // /
            7, 7, 7, 7, 7, 7, 7, 7, 7, 7, // 0-9
            4,  // :
            4,  // ;
            7,  // <
            7,  // =
            7,  // >
            7,  // ?
            13, // @
            9, 8, 8, 9, 7, 7, 9, 9, 4, 5, 8, 6, 12, // A-M
            10, 10, 8, 10, 8, 7, 7, 9, 9, 13, 8, 7, 7, // N-Z
            5,  // [
            6,  // \
            5,  // ]
            7,  // ^
            7,  // _
            4,  // `
            7, 8, 6, 8, 8, 5, 7, 8, 4, 4, 7, 4, 12, // a-m
            8, 8, 8, 8, 5, 6, 5, 8, 7, 11, 7, 7, 6, // n-z
            5,  // {
            7,  // |
            5,  // }
            7,  // ~
        };

        // Latin-1 Supplement U+00A0..U+00FF. Accented letters take the width of their base letter; the symbols
        // are the nearest ASCII equivalent (© and ® like &, superscripts like a narrow digit).
        private static readonly byte[] Latin1 =
        {
            3, 5, 7, 7, 7, 7, 7, 7, 4, 10, 5, 7, 7, 5, 10, 7,   // nbsp ¡ ¢ £ ¤ ¥ ¦ § ¨ © ª « ¬ shy ® ¯
            5, 7, 5, 5, 4, 8, 8, 4, 4, 5, 5, 7, 10, 10, 10, 7,  // ° ± ² ³ ´ µ ¶ · ¸ ¹ º » ¼ ½ ¾ ¿
            9, 9, 9, 9, 9, 9, 13, 8, 7, 7, 7, 7, 4, 4, 4, 4,    // À Á Â Ã Ä Å Æ Ç È É Ê Ë Ì Í Î Ï
            9, 10, 10, 10, 10, 10, 10, 7, 10, 9, 9, 9, 9, 7, 8, 8, // Ð Ñ Ò Ó Ô Õ Ö × Ø Ù Ú Û Ü Ý Þ ß
            7, 7, 7, 7, 7, 7, 12, 6, 8, 8, 8, 8, 4, 4, 4, 4,    // à á â ã ä å æ ç è é ê ë ì í î ï
            8, 8, 8, 8, 8, 8, 8, 7, 8, 8, 8, 8, 8, 7, 8, 7,     // ð ñ ò ó ô õ ö ÷ ø ù ú û ü ý þ ÿ
        };

        /// <summary>
        /// Width of <paramref name="text"/> in digit units. A multi-line cell (wrap text, Alt+Enter) is as wide
        /// as its widest line. Allocation-free.
        /// </summary>
        public static double Measure(string text, bool bold)
        {
            var longest = 0;
            var current = 0;
            for (var i = 0; i < text.Length; i++)
            {
                var c = text[i];
                if (c == '\n')
                {
                    if (current > longest) longest = current;
                    current = 0;
                }
                else current += PixelWidth(c, bold);
            }
            return (current > longest ? current : longest) / (double)DigitWidthPixels;
        }

        /// <summary>Hinted advance width of one UTF-16 code unit in pixels at Calibri 11 / 96 dpi.</summary>
        public static int PixelWidth(char c, bool bold)
        {
            if (c < 0x20) return 0; // Control characters (tab, CR) take no room in a cell.
            if (c < 0x7F)
            {
                if (bold)
                {
                    // The only ASCII glyphs Calibri Bold widens on the 11 pt pixel grid.
                    switch (c)
                    {
                        case '"': return 7;
                        case '&': return 11;
                        case '\'': return 4;
                        case 'G': return 10;
                        case 'M': return 13;
                        case 'U': return 10;
                        case 'W': return 14;
                        case 'Y': return 8;
                        case '`': return 5;
                    }
                }
                return Ascii[c - 0x20];
            }
            if (c < 0xA0) return 0; // C1 controls.
            if (c <= 0xFF) return Latin1[c - 0xA0];
            if (c >= 0x0300 && c <= 0x036F) return 0; // Combining diacritics sit on the previous glyph.
            if (c >= 0x200B && c <= 0x200F) return 0; // Zero-width space, joiners, direction marks.
            if (c == 0xFEFF) return 0;
            if (char.IsLowSurrogate(c)) return 0; // Counted with the high surrogate.
            if (char.IsHighSurrogate(c)) return WideWidth; // Supplementary planes: emoji and CJK extensions, all wide.
            if (IsWide(c)) return bold ? WideWidth + 1 : WideWidth;
            return bold ? DefaultWidth + 1 : DefaultWidth;
        }

        /// <summary>True for East Asian full-width ranges (Hangul, Kana, CJK ideographs, full-width forms).</summary>
        private static bool IsWide(char c)
            => (c >= 0x1100 && c <= 0x115F)
            || (c >= 0x2E80 && c <= 0x303E)
            || (c >= 0x3041 && c <= 0x33FF)
            || (c >= 0x3400 && c <= 0x4DBF)
            || (c >= 0x4E00 && c <= 0x9FFF)
            || (c >= 0xA000 && c <= 0xA4CF)
            || (c >= 0xAC00 && c <= 0xD7A3)
            || (c >= 0xF900 && c <= 0xFAFF)
            || (c >= 0xFE30 && c <= 0xFE4F)
            || (c >= 0xFF00 && c <= 0xFF60)
            || (c >= 0xFFE0 && c <= 0xFFE6);
    }
}
