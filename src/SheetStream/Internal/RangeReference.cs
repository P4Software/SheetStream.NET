using System;
using System.Globalization;

namespace SheetStream.Internal
{
    /// <summary>
    /// A parsed rectangular cell range. Accepts "B2", "B2:D10" and "$B$2:$D$10". Whole-row and whole-column
    /// forms are deliberately not supported, because every feature that takes a range here (validation,
    /// conditional formatting, tables, charts) needs concrete corners.
    /// </summary>
    internal readonly struct RangeReference
    {
        public RangeReference(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            if (fromRow < 1 || toRow > CellReference.MaxRows || fromRow > toRow)
                throw new ArgumentOutOfRangeException(nameof(fromRow), "Range rows must satisfy 1 <= fromRow <= toRow <= 1048576.");
            if (fromColumn < 1 || toColumn > CellReference.MaxColumns || fromColumn > toColumn)
                throw new ArgumentOutOfRangeException(nameof(fromColumn), "Range columns must satisfy 1 <= fromColumn <= toColumn <= 16384.");
            FromRow = fromRow;
            FromColumn = fromColumn;
            ToRow = toRow;
            ToColumn = toColumn;
        }

        public int FromRow { get; }
        public int FromColumn { get; }
        public int ToRow { get; }
        public int ToColumn { get; }

        public int RowCount => ToRow - FromRow + 1;
        public int ColumnCount => ToColumn - FromColumn + 1;

        /// <summary>Relative A1 form, "B2:D10", or "B2" for a single cell (what sqref and ref attributes expect).</summary>
        public override string ToString()
            => FromRow == ToRow && FromColumn == ToColumn
                ? CellReference.ToA1(FromRow, FromColumn)
                : CellReference.ToRange(FromRow, FromColumn, ToRow, ToColumn);

        /// <summary>Absolute form, "$B$2:$D$10" (or "$B$2" for one cell), used in formulas and defined names.</summary>
        public string ToAbsolute()
        {
            var range = CellReference.ToAbsoluteRange(FromRow, FromColumn, ToRow, ToColumn);
            return FromRow == ToRow && FromColumn == ToColumn ? range.Substring(0, range.IndexOf(':')) : range;
        }

        /// <summary>Absolute form qualified with a sheet name: <c>'My Sheet'!$B$2:$D$10</c>.</summary>
        public string ToAbsolute(string sheetName) => QuoteSheetName(sheetName) + "!" + ToAbsolute();

        /// <summary>Quotes a sheet name for use in a formula, doubling embedded apostrophes.</summary>
        public static string QuoteSheetName(string sheetName) => "'" + sheetName.Replace("'", "''") + "'";

        /// <summary>True when the two ranges share at least one cell.</summary>
        public bool Overlaps(RangeReference other)
            => FromRow <= other.ToRow && other.FromRow <= ToRow && FromColumn <= other.ToColumn && other.FromColumn <= ToColumn;

        /// <summary>Parses "B2" or "B2:D10" (dollar signs allowed). Throws <see cref="ArgumentException"/> naming the bad text.</summary>
        public static RangeReference Parse(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException("Range cannot be empty.", nameof(text));
            var parts = text.Trim().Split(':');
            if (parts.Length > 2) throw new ArgumentException($"Range '{text}' is not a valid A1 range.", nameof(text));
            var (r1, c1) = ParseCell(parts[0], text);
            if (parts.Length == 1) return new RangeReference(r1, c1, r1, c1);
            var (r2, c2) = ParseCell(parts[1], text);
            return new RangeReference(Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
        }

        /// <summary>
        /// Parses "B2" or "B2:D10" like <see cref="Parse"/>, returning false for anything else instead of
        /// throwing. For ranges read from a workbook this library did not write: an exception costs
        /// microseconds, and a file can hold hundreds of thousands of tokens that are not ranges.
        /// </summary>
        public static bool TryParse(string? text, out RangeReference range)
        {
            range = default;
            if (string.IsNullOrWhiteSpace(text)) return false;
            var s = text!.Trim();
            var colon = s.IndexOf(':');
            if (colon < 0)
            {
                if (!TryParseCell(s, 0, s.Length, out var row, out var column)) return false;
                range = new RangeReference(row, column, row, column);
                return true;
            }
            if (s.IndexOf(':', colon + 1) >= 0) return false;
            if (!TryParseCell(s, 0, colon, out var r1, out var c1) || !TryParseCell(s, colon + 1, s.Length, out var r2, out var c2)) return false;
            range = new RangeReference(Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
            return true;
        }

        /// <summary>Parses one A1 cell reference into (row, column).</summary>
        public static (int Row, int Column) ParseCell(string cell, string? context = null)
        {
            var s = cell.Trim();
            if (!TryParseCell(s, 0, s.Length, out var row, out var column))
                throw new ArgumentException($"'{context ?? cell}' is not a valid A1 cell reference.", nameof(cell));
            return (row, column);
        }

        /// <summary>Parses the A1 cell reference in <paramref name="s"/> between two positions, without allocating.</summary>
        private static bool TryParseCell(string s, int start, int end, out int row, out int column)
        {
            row = 0;
            column = 0;
            while (start < end && char.IsWhiteSpace(s[start])) start++;
            while (end > start && char.IsWhiteSpace(s[end - 1])) end--;
            var i = start;
            if (i < end && s[i] == '$') i++;
            var letters = 0;
            while (i < end && ((s[i] >= 'A' && s[i] <= 'Z') || (s[i] >= 'a' && s[i] <= 'z')))
            {
                if (++letters > 3) return false;
                column = column * 26 + ((s[i] & ~0x20) - 'A' + 1);
                i++;
            }
            if (i < end && s[i] == '$') i++;
            var digits = 0;
            long value = 0;
            while (i < end && s[i] >= '0' && s[i] <= '9')
            {
                if (++digits > 7) return false; // Excel's last row has seven digits.
                value = value * 10 + (s[i] - '0');
                i++;
            }
            if (letters == 0 || digits == 0 || i != end || value < 1 || value > CellReference.MaxRows || column > CellReference.MaxColumns) return false;
            row = (int)value;
            return true;
        }

        /// <summary>
        /// Turns a user range into an absolute, sheet-qualified reference. Input may already carry a sheet
        /// ("'Data'!B2:B10" or "Data!B2:B10"); otherwise <paramref name="defaultSheet"/> is used.
        /// </summary>
        public static string QualifyAbsolute(string text, string defaultSheet)
        {
            if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException("Range cannot be empty.", nameof(text));
            var t = text.Trim();
            if (t.StartsWith("=", StringComparison.Ordinal)) t = t.Substring(1);
            var bang = t.LastIndexOf('!');
            if (bang < 0) return Parse(t).ToAbsolute(defaultSheet);
            var sheet = t.Substring(0, bang).Trim();
            if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
                sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
            return Parse(t.Substring(bang + 1)).ToAbsolute(sheet);
        }
    }
}
