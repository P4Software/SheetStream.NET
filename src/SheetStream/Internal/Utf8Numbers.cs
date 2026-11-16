using System;
using System.Globalization;
using System.Xml;

namespace SheetStream.Internal
{
    /// <summary>
    /// Formats numbers straight into a UTF-8 byte buffer, which is where the row writer composes the sheet XML.
    /// The text is the text <see cref="XmlConvert"/> produces (it is what <c>XmlWriter.WriteValue</c> used): invariant
    /// culture, integers in plain digits, doubles in the shortest form that reads back as the same number
    /// (<c>0.1</c>, <c>1E-05</c>). A test pins that against <c>double.ToString("R")</c> over millions of values.
    /// </summary>
    internal static class Utf8Numbers
    {
        /// <summary>
        /// A buffer with this much room after the offset holds any number written here: the longest is a decimal such
        /// as <c>-79228162514264337593543950335</c> (30 characters), and a double in round-trip form stays under 26.
        /// </summary>
        public const int MaxLength = 40;

        // "00" "01" ... "99": two digits per division halves the work of printing an integer.
        private static readonly byte[] Pairs = BuildPairs();

        private static byte[] BuildPairs()
        {
            var pairs = new byte[200];
            for (var i = 0; i < 100; i++)
            {
                pairs[2 * i] = (byte)('0' + i / 10);
                pairs[2 * i + 1] = (byte)('0' + i % 10);
            }
            return pairs;
        }

        // The fast path below produces the shortest text that reads back as the same number, which is what
        // double.ToString("R") gives on .NET Core 3.0 and later. .NET Framework and .NET Core 1 and 2 do not: they try
        // 15 digits, then 17, and their parser is not always exact, so for some values they write 17 digits where a
        // shorter text exists. A file must not change with the runtime that wrote it, so there the runtime's own
        // formatter does all the work, as it always did. The net8.0 build only ever runs on a runtime that is right.
#if NET8_0_OR_GREATER
        private const bool RuntimeFormatsShortest = true;
#else
        private static readonly bool RuntimeFormatsShortest = DetectShortestFormatting();

        private static bool DetectShortestFormatting()
        {
            // ".NET Core 3.1.x", ".NET 5.0.x", ".NET 8.0.x" are right; ".NET Framework 4.8.x", ".NET Core 2.1.x" and
            // anything unknown (Mono, for one) keep the old behaviour.
            var description = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription;
            if (description.StartsWith(".NET Core ", StringComparison.Ordinal))
                return description.Length > 10 && description[10] >= '3' && description[10] <= '9';
            return description.StartsWith(".NET ", StringComparison.Ordinal) && description.Length > 5 && char.IsDigit(description[5]);
        }
#endif

        // Powers of ten that a double holds exactly (10^22 is the last one). The fast double path needs up to 10^6.
        private static readonly double[] PowersOfTen = { 1, 10, 100, 1e3, 1e4, 1e5, 1e6 };
        private static readonly long[] LongPowersOfTen = { 1, 10, 100, 1000, 10000, 100000, 1000000 };

        /// <summary>
        /// Writes <paramref name="value"/> as plain digits (with a leading minus sign when negative) at
        /// <paramref name="offset"/> and returns how many bytes were used. The caller guarantees the room.
        /// </summary>
        public static int Format(long value, byte[] buffer, int offset)
        {
            if (value >= 0) return FormatUnsigned((ulong)value, buffer, offset);
            buffer[offset] = (byte)'-';
            // The magnitude comes from the bitwise complement so that long.MinValue, which has no positive
            // counterpart in a long, still comes out right.
            return 1 + FormatUnsigned((ulong)~value + 1UL, buffer, offset + 1);
        }

        private static int FormatUnsigned(ulong value, byte[] buffer, int offset)
        {
            // Most cells are small integers (ids, quantities), so the one-digit case is first.
            if (value < 10)
            {
                buffer[offset] = (byte)('0' + (int)value);
                return 1;
            }

            var digits = CountDigits(value);
            var position = offset + digits;
            var pairs = Pairs;
            while (value >= 100)
            {
                var quotient = value / 100;
                var pair = (int)(value - quotient * 100) * 2;
                value = quotient;
                position -= 2;
                buffer[position] = pairs[pair];
                buffer[position + 1] = pairs[pair + 1];
            }
            if (value >= 10)
            {
                var pair = (int)value * 2;
                buffer[position - 2] = pairs[pair];
                buffer[position - 1] = pairs[pair + 1];
            }
            else
            {
                buffer[position - 1] = (byte)('0' + (int)value);
            }
            return digits;
        }

        private static int CountDigits(ulong value)
        {
            if (value < 100) return value < 10 ? 1 : 2;
            if (value < 10000) return value < 1000 ? 3 : 4;
            if (value < 1000000) return value < 100000 ? 5 : 6;
            if (value < 100000000) return value < 10000000 ? 7 : 8;
            var digits = 9;
            for (var limit = 1000000000UL; digits < 20 && value >= limit; limit *= 10) digits++;
            return digits;
        }

        /// <summary>
        /// Writes <paramref name="value"/> in the shortest form that reads back as the same number and returns the
        /// length. NaN and the infinities are not handled; the row writer turns them into an error cell before it
        /// gets here.
        /// </summary>
        /// <remarks>
        /// Business numbers are overwhelmingly short decimals (<c>1234.5</c>, <c>0.07</c>, <c>19.99</c>). For those the
        /// shortest form is found with a handful of multiplications: if rounding <c>value * 10^d</c> to a whole number
        /// <c>m</c> and dividing back (<c>m / 10^d</c>, one correctly rounded division of two exact numbers) gives
        /// <c>value</c> again, then the decimal <c>m / 10^d</c> reads back as <c>value</c>. With fewer than 16
        /// significant digits at most one decimal can do that (the spacing of such decimals is wider than the gap
        /// between two doubles), so it is the shortest one, and trying fewer digits first never skips a shorter one.
        /// Anything else (long noise like <c>0.30000000000000004</c>, huge or tiny magnitudes, negative zero) goes
        /// to the runtime's own formatter, which gives the text this method has to match.
        /// </remarks>
        public static int Format(double value, byte[] buffer, int offset)
        {
            var magnitude = Math.Abs(value);
            // 1e-4 and 1e15 are where double.ToString("R") switches to exponent notation (1E-05, 1E+15).
            if (RuntimeFormatsShortest && magnitude >= 1e-4 && magnitude < 1e15)
            {
                // Two, four, then six decimals; a value that needs more is noise the fast path cannot help with.
                for (var d = 2; d <= 6; d += 2)
                {
                    var scaled = magnitude * PowersOfTen[d];
                    if (scaled >= 1e15) break; // 16 or more significant digits: not unique, leave it to the runtime.
                    var m = (long)(scaled + 0.5);
                    if ((double)m / PowersOfTen[d] != magnitude) continue;

                    // m / 10^d is the number; take off the zeros it does not need (12.50 is 12.5, 12.00 is 12).
                    var decimals = d;
                    while (decimals > 0 && m % 10 == 0)
                    {
                        m /= 10;
                        decimals--;
                    }
                    return WriteDecimal(value < 0, m, decimals, buffer, offset);
                }
            }
            else if (RuntimeFormatsShortest && value == 0 && !IsNegativeZero(value))
            {
                buffer[offset] = (byte)'0';
                return 1;
            }
            return FormatWithRuntime(value, buffer, offset);
        }

        private static bool IsNegativeZero(double value) => BitConverter.DoubleToInt64Bits(value) < 0;

        /// <summary>Writes <c>m / 10^decimals</c> in plain digits, with a leading zero before the point when m is smaller.</summary>
        private static int WriteDecimal(bool negative, long m, int decimals, byte[] buffer, int offset)
        {
            var start = offset;
            if (negative) buffer[offset++] = (byte)'-';
            if (decimals == 0)
                return offset - start + FormatUnsigned((ulong)m, buffer, offset);

            var scale = LongPowersOfTen[decimals];
            var whole = m / scale;
            var fraction = m - whole * scale;
            offset += FormatUnsigned((ulong)whole, buffer, offset);
            buffer[offset++] = (byte)'.';
            // The fraction keeps its leading zeros (0.05 has the fraction 5 over two decimals), so it is written
            // right to left into exactly 'decimals' places.
            for (var i = offset + decimals - 1; i >= offset; i--)
            {
                var next = fraction / 10;
                buffer[i] = (byte)('0' + (int)(fraction - next * 10));
                fraction = next;
            }
            return offset + decimals - start;
        }

        private static int FormatWithRuntime(double value, byte[] buffer, int offset)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(buffer.AsSpan(offset), out var written, "R", CultureInfo.InvariantCulture)) return written;
#endif
            // netstandard2.0 cannot format into a buffer. This is the same text XmlWriter.WriteValue(double) wrote.
            return CopyAscii(XmlConvert.ToString(value), buffer, offset);
        }

        /// <summary>Writes <paramref name="value"/> in invariant culture, as <see cref="XmlConvert.ToString(decimal)"/> does.</summary>
        public static int Format(decimal value, byte[] buffer, int offset)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(buffer.AsSpan(offset), out var written, default, CultureInfo.InvariantCulture)) return written;
#endif
            return CopyAscii(XmlConvert.ToString(value), buffer, offset);
        }

        /// <summary>Copies a string known to hold only ASCII (digits, signs, 'E', '.') into the buffer.</summary>
        private static int CopyAscii(string text, byte[] buffer, int offset)
        {
            for (var i = 0; i < text.Length; i++) buffer[offset + i] = (byte)text[i];
            return text.Length;
        }
    }
}
