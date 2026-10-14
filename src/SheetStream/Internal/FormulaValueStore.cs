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
    /// results. One growable <c>double[]</c> per column, indexed by row, so a numeric cell costs 8 bytes. The
    /// three non-numeric states are stored in the same slot as NaN values with distinct payloads, which is why
    /// no second array is needed: a real NaN is never recorded because the writer turns NaN and infinities
    /// into <c>#NUM!</c> error cells.
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

        private double[]?[] _columns = new double[]?[32];

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
            var cells = _columns[column];
            if (cells == null || row >= cells.Length) return FormulaCellState.Empty;
            var v = cells[row];
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

        /// <summary>Returns the slot for a cell, growing the column array and the column's row array as needed.</summary>
        private ref double Slot(int row, int column)
        {
            if (column >= _columns.Length)
            {
                var grown = new double[]?[Math.Min(CellReference.MaxColumns + 1, Math.Max(column + 1, _columns.Length * 2))];
                Array.Copy(_columns, grown, _columns.Length);
                _columns = grown;
            }
            var cells = _columns[column];
            if (cells == null || row >= cells.Length)
            {
                // Doubling keeps the amortised cost flat; the tail is filled with the Empty sentinel because a
                // default 0.0 would read as a legitimate number.
                var oldLength = cells?.Length ?? 0;
                var grown = new double[Math.Min(CellReference.MaxRows + 1, Math.Max(row + 1, Math.Max(oldLength * 2, 256)))];
                if (cells != null) Array.Copy(cells, grown, oldLength);
                for (var i = oldLength; i < grown.Length; i++) grown[i] = EmptyValue;
                _columns[column] = cells = grown;
            }
            return ref cells[row];
        }
    }
}
