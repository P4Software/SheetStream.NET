using System;
using System.Globalization;

namespace SheetStream.Internal
{
    /// <summary>
    /// Formats numbers into a caller-owned character buffer, so the row writer can hand them to the XML writer
    /// without creating a string for every cell. <c>XmlWriter.WriteValue(double)</c> and friends build a new
    /// string each time; across a million cells that is tens of megabytes of garbage for no benefit.
    /// </summary>
    /// <remarks>
    /// The text is identical to what <see cref="System.Xml.XmlConvert"/> produces (which is what
    /// <c>WriteValue</c> uses): invariant culture, integers in plain digits and doubles in the shortest
    /// round-trip form (<c>0.1</c>, <c>1E-05</c>). A test pins that.
    /// </remarks>
    internal static class CharNumbers
    {
        /// <summary>
        /// A buffer of this length holds any number written here: the longest is a decimal such as
        /// <c>-79228162514264337593543950335</c> (30 characters), and a double in round-trip form stays under 26.
        /// </summary>
        public const int BufferLength = 40;

        /// <summary>
        /// Writes <paramref name="value"/> as plain digits (with a leading minus sign when negative) at the start
        /// of <paramref name="buffer"/> and returns how many characters were used.
        /// </summary>
        public static int Format(long value, char[] buffer)
        {
            // The magnitude is taken from the bitwise complement so that long.MinValue, which has no positive
            // counterpart in a long, still comes out right.
            var negative = value < 0;
            var magnitude = negative ? (ulong)~value + 1UL : (ulong)value;

            var digits = 1;
            for (var rest = magnitude; rest >= 10; rest /= 10) digits++;

            var length = negative ? digits + 1 : digits;
            var position = length;
            do
            {
                buffer[--position] = (char)('0' + (int)(magnitude % 10));
                magnitude /= 10;
            }
            while (magnitude != 0);
            if (negative) buffer[0] = '-';
            return length;
        }

        /// <summary>
        /// Writes <paramref name="value"/> in the shortest form that reads back as the same number and returns the
        /// length, or -1 when the runtime cannot format into a buffer (<c>netstandard2.0</c>), in which case the
        /// caller falls back to a string. NaN and the infinities are not handled; the row writer turns them into
        /// an error cell before it gets here.
        /// </summary>
        public static int Format(double value, char[] buffer)
        {
#if NET8_0_OR_GREATER
            return value.TryFormat(buffer, out var written, "R", CultureInfo.InvariantCulture) ? written : -1;
#else
            return -1;
#endif
        }

        /// <summary>Writes <paramref name="value"/> in invariant culture, or returns -1 as <see cref="Format(double, char[])"/> does.</summary>
        public static int Format(decimal value, char[] buffer)
        {
#if NET8_0_OR_GREATER
            return value.TryFormat(buffer, out var written, default, CultureInfo.InvariantCulture) ? written : -1;
#else
            return -1;
#endif
        }
    }
}
