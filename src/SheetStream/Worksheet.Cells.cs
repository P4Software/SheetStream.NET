using System;
using System.Collections.Generic;
using SheetStream.Internal;
using SheetStream.Reading;
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

            // Edit mode: a cell loaded from the workbook keeps its original XML so it is written back exactly as
            // it was (cached formula results, rich inline text, metadata attributes) unless the caller changes
            // its content. RawInner is null for cells created or changed by the caller.
            public string? RawInner;
            public string? RawType;
            public List<RawAttribute>? RawAttributes;

            /// <summary>
            /// True for the cell that holds the text of a shared formula. Its dependents only reference it, so
            /// changing its content would leave them without a formula; such a change is refused instead.
            /// </summary>
            public bool SharedFormulaMaster;

            /// <summary>Forgets the original XML because the caller is replacing the cell's content.</summary>
            public void DiscardRaw()
            {
                RawInner = null;
                RawType = null;
                RawAttributes = null;
                SharedFormulaMaster = false;
            }
        }

        /// <summary>
        /// Random-access cells, EPPlus style: <c>sheet.Cells[3, 2].Value = 42</c>, in any order. Using it puts
        /// the sheet in random-access mode: cells are kept in memory and written in order when the sheet is
        /// saved, so memory grows with the cell count. Use <see cref="Row()"/> streaming for large exports;
        /// a sheet uses one mode or the other. A sheet opened with <c>Worksheets.Edit</c> starts in this mode
        /// with every existing cell loaded.
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

        /// <summary>
        /// Returns the cell for a content change (value or formula), creating it if needed. A loaded cell drops
        /// its original XML, since the caller's content replaces it; the master cell of a shared formula is
        /// refused because its dependents would lose their formula.
        /// </summary>
        internal BufferedCell GetBufferedCellForContentChange(int row, int column)
        {
            var cell = GetBufferedCell(row, column, true);
            if (cell.SharedFormulaMaster) ThrowSharedFormulaMaster(row, column);
            cell.DiscardRaw();
            return cell;
        }

        /// <summary>Removes a cell from the buffer entirely (no value, no style, no formula).</summary>
        internal void RemoveBufferedCell(int row, int column)
        {
            EnsureNotSealed();
            if (_buffer == null || !_buffer.TryGetValue(row, out var cols) || !cols.TryGetValue(column, out var cell)) return;
            if (cell.SharedFormulaMaster) ThrowSharedFormulaMaster(row, column);
            cols.Remove(column);
            // Keep the row entry when it carries original row attributes (height, hidden): an empty <row> still
            // needs to be written for them. Otherwise an empty row would only add noise to the file.
            if (cols.Count == 0 && (_rowAttributes == null || !_rowAttributes.ContainsKey(row))) _buffer.Remove(row);
        }

        private void ThrowSharedFormulaMaster(int row, int column)
            => throw new NotSupportedException($"Cell {CellReference.ToA1(row, column)} on sheet '{Name}' holds the master copy of a shared formula; the cells that share it only point at it, so changing or clearing it would leave them without a formula. Change it in Excel, or set the formula of every cell in the shared range explicitly.");

        /// <summary>Snapshot of the buffered cells in row then column order, as handles.</summary>
        internal List<ExcelCell> SnapshotBufferedCells()
        {
            var list = new List<ExcelCell>();
            if (_buffer == null) return list;
            foreach (var rowPair in _buffer)
                foreach (var cellPair in rowPair.Value)
                    list.Add(new ExcelCell(this, rowPair.Key, cellPair.Key));
            return list;
        }

        /// <summary>Last row and column that hold a buffered cell, or (0, 0) when there is none.</summary>
        internal (int LastRow, int LastColumn) BufferDimension()
        {
            var lastRow = 0;
            var lastColumn = 0;
            if (_buffer == null) return (0, 0);
            foreach (var rowPair in _buffer)
            {
                if (rowPair.Value.Count == 0) continue;
                lastRow = rowPair.Key;
                foreach (var column in rowPair.Value.Keys)
                    if (column > lastColumn) lastColumn = column;
            }
            return (lastRow, lastColumn);
        }

        /// <summary>Writes the buffered cells through the streaming writer, in row and column order.</summary>
        private void FlushBuffer()
        {
            if (_buffer == null) return;
            foreach (var rowPair in _buffer)
            {
                List<RawAttribute>? rowAttributes = null;
                _rowAttributes?.TryGetValue(rowPair.Key, out rowAttributes);
                var writer = RowCore(rowPair.Key, null, rowAttributes);
                foreach (var cellPair in rowPair.Value)
                {
                    writer.Skip(cellPair.Key - writer.NextColumn);
                    var cell = cellPair.Value;
                    if (cell.RawInner != null)
                    {
                        writer.WriteRawCell(cell.RawType, cell.RawAttributes, cell.RawInner, cell.Style);
                        RecordLoadedCell(rowPair.Key, cellPair.Key, cell);
                    }
                    else if (cell.Formula != null) WriteBufferedFormula(writer, cell);
                    else if (cell.Value == null && cell.Style != null) writer.WriteEmpty(cell.Style);
                    else WriteBufferedValue(writer, cell);
                }
            }
            _buffer = null;
        }

        /// <summary>
        /// Writes a formula cell; a loaded cell whose style changed still carries its cached result, which is
        /// kept so readers that do not calculate (ExcelReader, viewers) see the same value as before.
        /// </summary>
        private static void WriteBufferedFormula(RowWriter writer, BufferedCell cell)
        {
            switch (cell.Value)
            {
                case double d: writer.WriteFormula(cell.Formula!, d, cell.Style); break;
                case string s: writer.WriteFormula(cell.Formula!, s, cell.Style); break;
                case bool b: writer.WriteFormula(cell.Formula!, b, cell.Style); break;
                default: writer.WriteFormula(cell.Formula!, cell.Style); break;
            }
        }

        /// <summary>Writes a value cell, covering the types a loaded workbook can yield beyond the ordinary ones.</summary>
        private void WriteBufferedValue(RowWriter writer, BufferedCell cell)
        {
            switch (cell.Value)
            {
                case ExcelErrorValue error:
                    writer.WriteErrorValue(error.Code, cell.Style);
                    break;
                case DateTime date when _templateSource != null && _templateSource.Date1904:
                    // A 1904-system workbook counts serial days from a different epoch; ToOADate assumes 1900.
                    var style = cell.Style ?? (date.TimeOfDay == TimeSpan.Zero ? Workbook.Styles.DateStyle : Workbook.Styles.DateTimeStyle);
                    writer.WriteCell(date.ToOADate() - 1462, style);
                    break;
                default:
                    writer.WriteCell(cell.Value, cell.Style);
                    break;
            }
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

        /// <summary>
        /// The cells that currently exist on the sheet (loaded from the workbook being edited or set through
        /// this collection), in row then column order. The list is a snapshot, so cells may be changed or
        /// cleared while iterating.
        /// </summary>
        /// <example><code>var lastRow = sheet.Cells.Existing.Select(c => c.Row).DefaultIfEmpty(0).Max();</code></example>
        public IEnumerable<ExcelCell> Existing => _sheet.SnapshotBufferedCells();

        /// <summary>The last row and last column that hold a cell, or (0, 0) on an empty sheet. Handy for appending.</summary>
        /// <example><code>var next = sheet.Cells.Dimension.LastRow + 1;</code></example>
        public (int LastRow, int LastColumn) Dimension => _sheet.BufferDimension();
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

        /// <summary>
        /// The value: string, number, bool, DateTime, TimeSpan, RichText[], <see cref="ExcelErrorValue"/> or
        /// null. On an edited workbook, loaded cells report their stored value (a formula cell its cached
        /// result; a number with a date format a <see cref="DateTime"/>). Setting it clears any formula;
        /// setting null empties the cell but keeps its style, see <see cref="Clear"/> to remove it.
        /// </summary>
        public object? Value
        {
            get => _sheet.GetBufferedCell(Row, Column, false)?.Value;
            set
            {
                var c = _sheet.GetBufferedCellForContentChange(Row, Column);
                c.Value = value;
                c.Formula = null;
            }
        }

        /// <summary>
        /// The cell style, or null for the default. On an edited workbook, the style index the cell had in the
        /// file, which stays valid in the saved copy. Changing only the style keeps the content as it was.
        /// </summary>
        public CellStyle? Style
        {
            get => _sheet.GetBufferedCell(Row, Column, false)?.Style;
            set => _sheet.GetBufferedCell(Row, Column, true).Style = value;
        }

        /// <summary>
        /// A formula such as "SUM(A1:A10)" (leading "=" optional). Setting it clears the value. On an edited
        /// workbook, the formula text stored in the file; a cell that shares another cell's formula reports
        /// null because the text lives in the master cell.
        /// </summary>
        public string? Formula
        {
            get => _sheet.GetBufferedCell(Row, Column, false)?.Formula;
            set
            {
                var c = _sheet.GetBufferedCellForContentChange(Row, Column);
                c.Formula = value;
                c.Value = null;
            }
        }

        /// <summary>Removes the cell altogether: value, formula and style. Nothing is written for it.</summary>
        public void Clear() => _sheet.RemoveBufferedCell(Row, Column);
    }
}
