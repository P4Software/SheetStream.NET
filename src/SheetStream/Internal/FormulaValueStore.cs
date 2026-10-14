using System;

namespace SheetStream.Internal
{
    /// <summary>What a recorded cell holds, as far as the formula evaluator is concerned.</summary>
    internal enum FormulaCellState
    {
        /// <summary>Nothing was written to the cell (or only a style).</summary>
        Empty,

        /// <summary>A number, date, time span or a formula whose numeric result is known.</summary>
        Number,

        /// <summary>Text, a boolean, an error or a formula with a text/boolean result: present but not a number.</summary>
        Text,

        /// <summary>Content whose value is not known (a formula without a cached result, a template cell).</summary>
        Unknown,
    }

    /// <summary>
    /// Compact record of the values written to a worksheet so far, kept only when the sheet caches formula
    /// results. Each column is a set of fixed-size pages of <c>double</c> allocated on demand, so a numeric
    /// cell in a dense export costs 8 bytes, while a single cell at row 1,000,000 costs one page (8 KB)
    /// rather than an 8 MB array. That bound matters in edit mode, where the rows come from a file the
    /// application did not write. The three non-numeric states are stored in the same slot as NaN values
    /// with distinct payloads, which is why no second array is needed: a real NaN is never recorded because
    /// the writer turns NaN and infinities into <c>#NUM!</c> error cells.
    /// </summary>
    internal sealed class FormulaValueStore
    {
        // NaN payloads for the non-numeric states. Compared by bit pattern since NaN != NaN.
        private const long EmptyBits = unchecked((long)0x7FF8000000000001UL);
        private const long TextBits = unchecked((long)0x7FF8000000000002UL);
        private const long UnknownBits = unchecked((long)0x7FF8000000000003UL);
        private static readonly double EmptyValue = BitConverter.Int64BitsToDouble(EmptyBits);
        private static readonly double TextValue = BitConverter.Int64BitsToDouble(TextBits);
        private static readonly double UnknownValue = BitConverter.Int64BitsToDouble(UnknownBits);

        // 1,024 rows per page: 8 KB, small enough that a stray far-away cell is cheap and large enough that a
        // dense column of a million rows needs only ~1,000 page references.
        private const int PageShift = 10;
        private const int PageSize = 1 << PageShift;
        private const int PageMask = PageSize - 1;

        // Column -> pages of that column -> values. Both outer levels grow by doubling, capped at the sheet limits.
        private double[]?[]?[] _columns = new double[]?[]?[32];

        // Rows whose content came from somewhere the writer did not see (copied template rows). A bit per row.
        private ulong[] _unknownRows = new ulong[0];

        /// <summary>Records a numeric value (numbers, dates and durations as serials, cached formula results).</summary>
        public void SetNumber(int row, int column, double value) => Slot(row, column) = value;

        /// <summary>Records that the cell holds something that is not a number (text, boolean, error).</summary>
        public void SetText(int row, int column) => Slot(row, column) = TextValue;

        /// <summary>Records that the cell has content whose value cannot be known here.</summary>
        public void SetUnknown(int row, int column) => Slot(row, column) = UnknownValue;

        /// <summary>Returns the state of a cell and, for <see cref="FormulaCellState.Number"/>, its value.</summary>
        public FormulaCellState Get(int row, int column, out double value)
        {
            value = 0;
            if (IsRowUnknown(row)) return FormulaCellState.Unknown;
            if (column >= _columns.Length) return FormulaCellState.Empty;
            var pages = _columns[column];
            var pageIndex = row >> PageShift;
            if (pages == null || pageIndex >= pages.Length) return FormulaCellState.Empty;
            var page = pages[pageIndex];
            if (page == null) return FormulaCellState.Empty;
            var v = page[row & PageMask];
            if (!double.IsNaN(v))
            {
                value = v;
                return FormulaCellState.Number;
            }
            var bits = BitConverter.DoubleToInt64Bits(v);
            return bits == TextBits ? FormulaCellState.Text : bits == UnknownBits ? FormulaCellState.Unknown : FormulaCellState.Empty;
        }

        /// <summary>Marks a whole row as opaque: a template row copied as XML, whose cells were never seen.</summary>
        public void MarkRowUnknown(int row)
        {
            var word = row >> 6;
            if (word >= _unknownRows.Length)
            {
                var grown = new ulong[Math.Max(word + 1, _unknownRows.Length * 2)];
                Array.Copy(_unknownRows, grown, _unknownRows.Length);
                _unknownRows = grown;
            }
            _unknownRows[word] |= 1UL << (row & 63);
        }

        /// <summary>True when <see cref="MarkRowUnknown"/> was called for the row.</summary>
        public bool IsRowUnknown(int row)
        {
            var word = row >> 6;
            return word < _unknownRows.Length && (_unknownRows[word] & (1UL << (row & 63))) != 0;
        }

        /// <summary>Returns the slot for a cell, growing the column list, the column's page list and the page as needed.</summary>
        private ref double Slot(int row, int column)
        {
            if (column >= _columns.Length)
            {
                var grown = new double[]?[]?[Math.Min(CellReference.MaxColumns + 1, Math.Max(column + 1, _columns.Length * 2))];
                Array.Copy(_columns, grown, _columns.Length);
                _columns = grown;
            }

            var pages = _columns[column];
            var pageIndex = row >> PageShift;
            if (pages == null || pageIndex >= pages.Length)
            {
                var oldLength = pages?.Length ?? 0;
                var grown = new double[]?[Math.Min((CellReference.MaxRows >> PageShift) + 1, Math.Max(pageIndex + 1, Math.Max(oldLength * 2, 4)))];
                if (pages != null) Array.Copy(pages, grown, oldLength);
                _columns[column] = pages = grown;
            }

            var page = pages[pageIndex];
            if (page == null)
            {
                // A fresh page is filled with the Empty sentinel because a default 0.0 would read as a legitimate number.
                page = new double[PageSize];
                for (var i = 0; i < page.Length; i++) page[i] = EmptyValue;
                pages[pageIndex] = page;
            }
            return ref page[row & PageMask];
        }
    }
}
