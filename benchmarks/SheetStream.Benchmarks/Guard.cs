using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SheetStream.Benchmarks;

/// <summary>
/// A <see cref="MemoryStream"/> whose <c>Dispose</c>/<c>Close</c> do nothing. Every library writes into one
/// of these; it is created once with enough capacity and reset before each write, so the output buffer's own
/// growth never shows up in a library's allocation figures. It also neutralises libraries that close the
/// stream they are given.
/// </summary>
public sealed class ReusableOutput : MemoryStream
{
    /// <summary>Creates the buffer with an initial capacity in bytes.</summary>
    public ReusableOutput(int capacity) : base(capacity) { }

    /// <summary>Empties the buffer while keeping its capacity.</summary>
    public void Reset()
    {
        Position = 0;
        SetLength(0);
    }

    /// <summary>Intentionally ignores disposal so the buffer can be reused and read back.</summary>
    protected override void Dispose(bool disposing) { }
}

/// <summary>What the guard found in one library's output.</summary>
/// <param name="Rows">Number of <c>&lt;row&gt;</c> elements in the first worksheet.</param>
/// <param name="LastRowCells">Number of cells in the last row.</param>
/// <param name="HeaderBold">Whether cell A1 resolves to a bold font.</param>
/// <param name="PriceType">Cell type of the first double column in the last row ("n" for a number).</param>
/// <param name="PriceFormat">Number format that cell resolves to.</param>
/// <param name="DateType">Cell type of the date column in the last row.</param>
/// <param name="DateFormat">Number format that cell resolves to.</param>
public sealed record GuardReport(int Rows, int LastRowCells, bool HeaderBold,
    string PriceType, string PriceFormat, string DateType, string DateFormat);

/// <summary>
/// Correctness guard: reads a library's output back with the Open XML SDK (a parser none of the libraries
/// under test share their writing code with) and checks that the expected rows and cells are really there,
/// so no library can "win" by writing nothing.
/// </summary>
public static class Guard
{
    /// <summary>Rows used for the quick pre-benchmark check of the Long scenario.</summary>
    public const int SmallLongRows = 50;

    /// <summary>Rows and columns used for the quick pre-benchmark check of the Wide scenario.</summary>
    public const int SmallWide = 40;

    /// <summary>
    /// Writes the small case of <paramref name="scenario"/> with <paramref name="writer"/> and throws if the
    /// output is short. Called from every benchmark's <c>GlobalSetup</c> and from <c>--verify</c>.
    /// </summary>
    public static GuardReport CheckSmall(ILibraryWriter writer, string scenario)
    {
        // Million shares Long's shape and VeryWide shares Wide's, so the small case is the same for each pair;
        // what differs between them (the row and column counts) is exactly what the small case leaves out.
        bool longShape = Scenarios.IsLongShape(scenario);
        using var output = new ReusableOutput(1 << 20);
        if (longShape) writer.WriteLong(output, Data.CreateLong(SmallLongRows));
        else writer.WriteWide(output, Data.CreateWide(SmallWide, SmallWide));
        return Check(writer.Name, scenario, output,
            longShape ? SmallLongRows : SmallWide,
            longShape ? Data.LongHeaders.Length : SmallWide);
    }

    /// <summary>
    /// Reads <paramref name="output"/> and throws <see cref="InvalidOperationException"/> unless the first
    /// worksheet has the expected number of rows and the last row has the expected number of cells.
    /// </summary>
    /// <param name="library">Library name, for the error message.</param>
    /// <param name="scenario">A Long-shaped scenario (header row + data) or a Wide-shaped one (data only); see <see cref="Scenarios.IsLongShape"/>.</param>
    /// <param name="output">The written package.</param>
    /// <param name="dataRows">Data rows written, excluding the header.</param>
    /// <param name="columns">Columns per row.</param>
    public static GuardReport Check(string library, string scenario, MemoryStream output, int dataRows, int columns)
    {
        // Read from a copy of the bytes so the reader can never disturb the writer's buffer.
        using var copy = new MemoryStream(output.GetBuffer(), 0, (int)output.Length, writable: false);
        using var doc = SpreadsheetDocument.Open(copy, false);
        var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException($"{library}: no workbook part.");
        var firstSheet = wbPart.Workbook!.Sheets!.Elements<Sheet>().First();
        var wsPart = (WorksheetPart)wbPart.GetPartById(firstSheet.Id!.Value!);

        int rows = 0, cellsInRow = 0;
        // Attributes of every cell in the current row, keyed by zero-based position; only the last row is kept.
        var rowCells = new List<(string? Type, uint Style)>();
        var a1 = (Type: (string?)null, Style: 0u);

        // SAX-style reader: the full-size check reads 100,000-row sheets without building a DOM.
        using (var reader = OpenXmlReader.Create(wsPart))
        {
            while (reader.Read())
            {
                if (!reader.IsStartElement) continue;
                if (reader.ElementType == typeof(Row))
                {
                    rows++;
                    cellsInRow = 0;
                    rowCells.Clear();
                }
                else if (reader.ElementType == typeof(Cell))
                {
                    string? type = null;
                    uint style = 0;
                    foreach (var attr in reader.Attributes)
                    {
                        if (attr.LocalName == "t") type = attr.Value;
                        else if (attr.LocalName == "s") style = uint.Parse(attr.Value!);
                    }
                    if (rows == 1 && cellsInRow == 0) a1 = (type, style);
                    rowCells.Add((type, style));
                    cellsInRow++;
                }
            }
        }

        bool longShape = Scenarios.IsLongShape(scenario);
        int expectedRows = longShape ? dataRows + 1 : dataRows;
        if (rows != expectedRows)
            throw new InvalidOperationException($"{library} {scenario}: expected {expectedRows} rows, found {rows}.");
        if (cellsInRow != columns)
            throw new InvalidOperationException($"{library} {scenario}: expected {columns} cells in the last row, found {cellsInRow}.");

        if (!longShape) return new GuardReport(rows, cellsInRow, false, "-", "-", "-", "-");

        var styles = wbPart.WorkbookStylesPart?.Stylesheet;
        return new GuardReport(rows, cellsInRow,
            HeaderBold: IsBold(styles, a1.Style),
            PriceType: rowCells[4].Type ?? "n",
            PriceFormat: FormatOf(styles, rowCells[4].Style),
            DateType: rowCells[6].Type ?? "n",
            DateFormat: FormatOf(styles, rowCells[6].Style));
    }

    /// <summary>Resolves a cell style index to its font and reports whether that font is bold.</summary>
    private static bool IsBold(Stylesheet? styles, uint styleIndex)
    {
        var xf = styles?.CellFormats?.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
        if (xf == null) return false;
        var font = styles!.Fonts?.Elements<Font>().ElementAtOrDefault((int)(xf.FontId?.Value ?? 0));
        // <b/> means bold; <b val="0"/> explicitly means not bold.
        return font?.Bold != null && (font.Bold.Val == null || font.Bold.Val.Value);
    }

    /// <summary>Resolves a cell style index to its number format code (built-in ids are shown as "builtin:N").</summary>
    private static string FormatOf(Stylesheet? styles, uint styleIndex)
    {
        var xf = styles?.CellFormats?.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
        uint id = xf?.NumberFormatId?.Value ?? 0;
        if (id == 0) return "General";
        var custom = styles!.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(f => f.NumberFormatId?.Value == id);
        return custom?.FormatCode?.Value ?? $"builtin:{id}";
    }
}
