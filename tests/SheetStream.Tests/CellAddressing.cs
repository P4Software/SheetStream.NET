using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SheetStream.Internal;
using XlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace SheetStream.Tests;

/// <summary>
/// Gives every cell of a worksheet its address the way a spreadsheet reader does. SheetStream leaves the
/// <c>r</c> attribute off a cell that sits in the column after the previous one (ECMA-376 makes it optional),
/// so <see cref="Cell.CellReference"/> is null for most of its cells; this works the address out from the
/// position instead: a missing row number is the row after the previous one, a missing reference is the
/// column after the previous cell, and a cell that does carry a reference moves the position to itself.
/// </summary>
internal static class CellAddressing
{
    /// <summary>Every cell of the worksheet with its address, in document order.</summary>
    public static IEnumerable<(string Address, Cell Cell)> WithAddresses(this XlWorksheet worksheet)
    {
        var rowIndex = 0;
        foreach (var row in worksheet.Descendants<Row>())
        {
            rowIndex = row.RowIndex?.Value is uint declared ? (int)declared : rowIndex + 1;
            var column = 0;
            foreach (var cell in row.Elements<Cell>())
            {
                var reference = cell.CellReference?.Value;
                if (reference != null)
                {
                    column = RangeReference.ParseCell(reference).Column;
                    yield return (reference, cell);
                }
                else
                {
                    column++;
                    yield return (SheetStream.Internal.CellReference.ToA1(rowIndex, column), cell);
                }
            }
        }
    }

    /// <summary>The worksheet's cells by address ("B2").</summary>
    public static Dictionary<string, Cell> CellsByAddress(this XlWorksheet worksheet)
        => worksheet.WithAddresses().ToDictionary(p => p.Address, p => p.Cell);
}
