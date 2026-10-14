using System;
using SheetStream.Internal;

namespace SheetStream
{
    public sealed partial class Worksheet
    {
        // Values written so far, kept only while CacheFormulaResults is on. Null means the feature is off and
        // the writer's per-cell hook is a single false branch.
        private FormulaValueStore? _formulaValues;

        /// <summary>
        /// Makes <see cref="RowWriter.WriteFormula(string, Styles.CellStyle?)"/> also write the formula's result
        /// as the cell's cached value whenever the library can compute it, so viewers that do not calculate
        /// (file previews, Quick Look, many mobile and web viewers, <see cref="Reading.ExcelReader"/>) show a
        /// number instead of a blank. Must be enabled before the first row is written.
        /// </summary>
        /// <remarks>
        /// <para>
        /// This is not a formula engine. Supported: numbers, references to cells of this sheet that were already
        /// written (rows above, or cells to the left on the same row), rectangular ranges, <c>+ - * /</c>, unary
        /// minus, a <c>%</c> suffix, parentheses, and <c>SUM</c>, <c>AVERAGE</c>, <c>COUNT</c>, <c>COUNTA</c>,
        /// <c>MIN</c>, <c>MAX</c>, <c>ROUND</c> and <c>ABS</c>. A formula that uses anything else, references a
        /// row not yet written or another sheet, or would produce an error (such as a division by zero) is
        /// written exactly as before: without a cached value, and Excel computes it on open. A wrong value is
        /// never written. A formula that does get a value counts as a numeric cell for later formulas, so a
        /// grand total over subtotals works.
        /// </para>
        /// <para>
        /// Cost: the sheet remembers every value written, about 8 bytes per cell of every column that holds
        /// numbers, which is why this is opt-in. Values written through <c>WriteCell</c>, <c>LoadFrom*</c> and
        /// <c>Cells[...]</c> are all seen. Rows copied from a template (<c>Worksheets.Open</c>) are not: a
        /// formula referencing a template row gets no cached value.
        /// </para>
        /// </remarks>
        /// <example>
        /// <code>
        /// sheet.CacheFormulaResults();
        /// for (var i = 0; i &lt; 100; i++) sheet.Row().WriteCell(i).WriteCell(i * 1.5);
        /// sheet.Row().WriteCell("Total").WriteFormula("SUM(B1:B100)"); // Written with its value, 7425.
        /// </code>
        /// </example>
        public Worksheet CacheFormulaResults(bool enabled = true)
        {
            EnsureNotSealed();
            if (!enabled)
            {
                _formulaValues = null;
                return this;
            }
            if (_formulaValues != null) return this;
            // Enabling after some rows have gone by would leave those rows unseen, and a SUM over them would be
            // wrong rather than missing; refusing is the only honest option.
            if (CurrentRow > 0)
                throw new InvalidOperationException($"CacheFormulaResults must be enabled on sheet '{Name}' before its first row is written, because the values of rows already streamed are gone.");
            _formulaValues = new FormulaValueStore();
            return this;
        }

        /// <summary>True when formula results are being cached; read by the row writer once per row.</summary>
        internal bool IsCachingFormulaResults => _formulaValues != null;

        /// <summary>Records a numeric cell (numbers, dates and durations as serials, cached formula results).</summary>
        internal void RecordNumber(int row, int column, double value) => _formulaValues?.SetNumber(row, column, value);

        /// <summary>Records a cell that holds non-numeric content (text, boolean, error, text or boolean formula result).</summary>
        internal void RecordText(int row, int column) => _formulaValues?.SetText(row, column);

        /// <summary>Records a cell whose value is not known here (a formula written without a cached result).</summary>
        internal void RecordUnknown(int row, int column) => _formulaValues?.SetUnknown(row, column);

        /// <summary>Records a copied template row whose cells were never seen.</summary>
        private void RecordUnknownRow(int row) => _formulaValues?.MarkRowUnknown(row);

        /// <summary>
        /// Records a cell flushed from the random-access buffer with its original XML (edit mode), from the typed
        /// value the loader produced. A formula without a cached result stays unknown.
        /// </summary>
        private void RecordLoadedCell(int row, int column, BufferedCell cell)
        {
            if (_formulaValues == null) return;
            switch (cell.Value)
            {
                case null:
                    if (cell.Formula != null) _formulaValues.SetUnknown(row, column);
                    break;
                case double d: _formulaValues.SetNumber(row, column, d); break;
                case DateTime date:
                    // Mirror what the writer would store: a 1904-system workbook counts serial days from a later epoch.
                    _formulaValues.SetNumber(row, column, date.ToOADate() - (_templateSource != null && _templateSource.Date1904 ? 1462 : 0));
                    break;
                case TimeSpan span: _formulaValues.SetNumber(row, column, span.TotalDays); break;
                default: _formulaValues.SetText(row, column); break;
            }
        }

        /// <summary>Evaluates a formula about to be written at the given cell; false when no value can be cached.</summary>
        internal bool TryEvaluateFormula(string formula, int row, int column, out double value)
        {
            if (_formulaValues == null)
            {
                value = 0;
                return false;
            }
            return FormulaEvaluator.TryEvaluate(formula, _formulaValues, row, column, Name, out value);
        }
    }
}
