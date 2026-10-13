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

        /// <summary>Parses one A1 cell reference into (row, column).</summary>
        public static (int Row, int Column) ParseCell(string cell, string? context = null)
        {
            var s = cell.Trim();
            var i = 0;
            if (i < s.Length && s[i] == '$') i++;
            var column = 0;
            var letters = 0;
            while (i < s.Length && ((s[i] >= 'A' && s[i] <= 'Z') || (s[i] >= 'a' && s[i] <= 'z')))
            {
                column = column * 26 + (char.ToUpperInvariant(s[i]) - 'A' + 1);
                letters++;
                i++;
            }
            if (i < s.Length && s[i] == '$') i++;
            var digitsStart = i;
            while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
            if (letters == 0 || letters > 3 || digitsStart == i || i != s.Length
                || !int.TryParse(s.Substring(digitsStart), NumberStyles.None, CultureInfo.InvariantCulture, out var row)
                || row < 1 || row > CellReference.MaxRows || column > CellReference.MaxColumns)
                throw new ArgumentException($"'{context ?? cell}' is not a valid A1 cell reference.", nameof(cell));
            return (row, column);
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
