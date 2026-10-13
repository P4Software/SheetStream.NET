using System;
using System.Globalization;
using System.Text;

namespace SheetStream.Internal
{
    /// <summary>
    /// Converts 1-based column indexes to Excel column letters ("A", "Z", "AA", "XFD") and builds
    /// A1-style references.
    /// </summary>
    /// <remarks>
    /// Letters are cached per column index because the same handful of columns is referenced on every
    /// row of a sheet. After the first row no allocation is needed to build a cell reference, which keeps
    /// the per-cell cost of the streaming writer close to zero.
    /// </remarks>
    internal static class CellReference
    {
        /// <summary>Highest column Excel supports (XFD).</summary>
        public const int MaxColumns = 16384;

        /// <summary>Highest row Excel supports.</summary>
        public const int MaxRows = 1048576;

        private static readonly string?[] Cache = new string?[MaxColumns + 1];

        /// <summary>Returns the column letters for a 1-based column index, e.g. 1 → "A", 28 → "AB".</summary>
        public static string ColumnLetters(int column)
        {
            if (column < 1 || column > MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(column), column, "Column index must be between 1 and 16384.");

            var cached = Cache[column];
            if (cached != null) return cached;

            // Build right-to-left: Excel columns are a bijective base-26 numeral system (there is no zero digit).
            var buffer = new char[3];
            var pos = 3;
            var n = column;
            while (n > 0)
            {
                n--;
                buffer[--pos] = (char)('A' + n % 26);
                n /= 26;
            }
            var letters = new string(buffer, pos, 3 - pos);
            Cache[column] = letters; // Benign race: two threads may compute the same immutable string.
            return letters;
        }

        /// <summary>Returns an A1-style reference such as "C7".</summary>
        public static string ToA1(int row, int column)
            => ColumnLetters(column) + row.ToString(CultureInfo.InvariantCulture);

        /// <summary>Returns a range reference such as "A1:C1".</summary>
        public static string ToRange(int fromRow, int fromColumn, int toRow, int toColumn)
            => ToA1(fromRow, fromColumn) + ":" + ToA1(toRow, toColumn);

        /// <summary>Returns an absolute range such as "$A$1:$C$1" (used in defined names).</summary>
        public static string ToAbsoluteRange(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            var sb = new StringBuilder(16);
            sb.Append('$').Append(ColumnLetters(fromColumn)).Append('$').Append(fromRow.ToString(CultureInfo.InvariantCulture));
            sb.Append(':');
            sb.Append('$').Append(ColumnLetters(toColumn)).Append('$').Append(toRow.ToString(CultureInfo.InvariantCulture));
            return sb.ToString();
        }
    }
}
