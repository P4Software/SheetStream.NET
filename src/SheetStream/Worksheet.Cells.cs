using System;
using System.Collections.Generic;
using SheetStream.Internal;
using SheetStream.Styles;

namespace SheetStream
{
    public sealed partial class Worksheet
    {
        private SortedDictionary<int, SortedDictionary<int, BufferedCell>>? _buffer;
        private CellCollection? _cells;
        private bool _streamedRows;

        /// <summary>A cell held in memory until the sheet is saved (random-access mode).</summary>
        internal sealed class BufferedCell
        {
            public object? Value;
            public CellStyle? Style;
            public string? Formula;
        }

        /// <summary>
        /// Random-access cells, EPPlus style: <c>sheet.Cells[3, 2].Value = 42</c>, in any order. Using it puts
        /// the sheet in random-access mode: cells are kept in memory and written in order when the sheet is
        /// saved, so memory grows with the cell count. Use <see cref="Row()"/> streaming for large exports;
        /// a sheet uses one mode or the other.
        /// </summary>
        public CellCollection Cells
        {
            get
            {
                if (_streamedRows)
                    throw new InvalidOperationException($"Worksheet '{Name}' already streams rows through Row(). A sheet uses either Row() streaming or Cells[...] random access, not both.");
                EnsureNotSealed();
                _buffer ??= new SortedDictionary<int, SortedDictionary<int, BufferedCell>>();
                return _cells ??= new CellCollection(this);
            }
        }

        internal BufferedCell GetBufferedCell(int row, int column, bool create)
        {
            EnsureNotSealed();
            if (!_buffer!.TryGetValue(row, out var cols))
            {
                if (!create) return null!;
                cols = new SortedDictionary<int, BufferedCell>();
                _buffer[row] = cols;
            }
            if (!cols.TryGetValue(column, out var cell))
            {
                if (!create) return null!;
                cell = new BufferedCell();
                cols[column] = cell;
            }
            return cell;
        }

        /// <summary>Writes the buffered cells through the streaming writer, in row and column order.</summary>
        private void FlushBuffer()
        {
            if (_buffer == null) return;
            foreach (var rowPair in _buffer)
            {
                var writer = RowCore(rowPair.Key, null);
                foreach (var cellPair in rowPair.Value)
                {
                    writer.Skip(cellPair.Key - writer.NextColumn);
                    var cell = cellPair.Value;
                    if (cell.Formula != null) writer.WriteFormula(cell.Formula, cell.Style);
                    else if (cell.Value == null && cell.Style != null) writer.WriteEmpty(cell.Style);
                    else writer.WriteCell(cell.Value, cell.Style);
                }
            }
            _buffer = null;
        }
    }

    /// <summary>Indexer over a sheet's cells for random access. See <see cref="Worksheet.Cells"/>.</summary>
    public sealed class CellCollection
    {
        private readonly Worksheet _sheet;

        internal CellCollection(Worksheet sheet)
        {
            _sheet = sheet;
        }

        /// <summary>The cell at a 1-based row and column.</summary>
        public ExcelCell this[int row, int column]
        {
            get
            {
                if (row < 1 || row > CellReference.MaxRows) throw new ArgumentOutOfRangeException(nameof(row), row, "Row index must be between 1 and 1048576.");
                CellReference.ColumnLetters(column);
                return new ExcelCell(_sheet, row, column);
            }
        }

        /// <summary>The cell at an A1 address such as "B7".</summary>
        public ExcelCell this[string address]
        {
            get
            {
                var (row, column) = RangeReference.ParseCell(address);
                return new ExcelCell(_sheet, row, column);
            }
        }
    }

    /// <summary>One cell in random-access mode.</summary>
    public sealed class ExcelCell
    {
        private readonly Worksheet _sheet;

        internal ExcelCell(Worksheet sheet, int row, int column)
        {
            _sheet = sheet;
            Row = row;
            Column = column;
        }

        /// <summary>1-based row.</summary>
        public int Row { get; }

        /// <summary>1-based column.</summary>
        public int Column { get; }

        /// <summary>A1 address.</summary>
        public string Address => CellReference.ToA1(Row, Column);

        /// <summary>The value: string, number, bool, DateTime, TimeSpan, RichText[] or null. Setting it clears any formula.</summary>
        public object? Value
        {
            get => _sheet.GetBufferedCell(Row, Column, false)?.Value;
            set
            {
                var c = _sheet.GetBufferedCell(Row, Column, true);
                c.Value = value;
                c.Formula = null;
            }
        }

        /// <summary>The cell style, or null for the default.</summary>
        public CellStyle? Style
        {
            get => _sheet.GetBufferedCell(Row, Column, false)?.Style;
            set => _sheet.GetBufferedCell(Row, Column, true).Style = value;
        }

        /// <summary>A formula such as "SUM(A1:A10)" (leading "=" optional). Setting it clears the value.</summary>
        public string? Formula
        {
            get => _sheet.GetBufferedCell(Row, Column, false)?.Formula;
            set
            {
                var c = _sheet.GetBufferedCell(Row, Column, true);
                c.Formula = value;
                c.Value = null;
            }
        }
    }
}
