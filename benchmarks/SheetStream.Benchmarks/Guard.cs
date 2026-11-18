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
/// <param name="ValuesChecked">Cells whose value was compared with the generator's value (0 when only counts were checked).</param>
/// <param name="Formats">For the styled scenario, the number formats found in the first data row; "-" otherwise.</param>
public sealed record GuardReport(int Rows, int LastRowCells, bool HeaderBold,
    string PriceType, string PriceFormat, string DateType, string DateFormat,
    int ValuesChecked = 0, string Formats = "-");

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
    /// Data rows used for the quick pre-benchmark check of the string and styled scenarios. 60 rows are enough to
    /// reach every value cycle of the generators (the longest repeats after 50 rows) and every escaped character.
    /// </summary>
    public const int SmallRows = 60;

    /// <summary>
    /// Writes the small case of <paramref name="scenario"/> with <paramref name="writer"/> and throws if the
    /// output is short. For the string and styled scenarios it also reads every cell back and throws unless each
    /// value (and, for the styled scenario, each number format) equals what the generator produced. Called from
    /// every benchmark's <c>GlobalSetup</c> and from <c>--verify</c>.
    /// </summary>
    public static GuardReport CheckSmall(ILibraryWriter writer, string scenario)
    {
        // Million shares Long's shape and VeryWide shares Wide's, so the small case is the same for each pair;
        // what differs between them (the row and column counts) is exactly what the small case leaves out.
        var shape = Scenarios.Shape(scenario);
        using var output = new ReusableOutput(1 << 20);
        switch (shape)
        {
            case ScenarioShape.Long:
                writer.WriteLong(output, Data.CreateLong(SmallLongRows));
                return Check(writer.Name, scenario, output, SmallLongRows, Data.LongHeaders.Length);
            case ScenarioShape.Wide:
                writer.WriteWide(output, Data.CreateWide(SmallWide, SmallWide));
                return Check(writer.Name, scenario, output, SmallWide, SmallWide);
            case ScenarioShape.Strings:
            {
                var rows = scenario == Scenarios.SpreadCheetahOwn ? Data.CreateOwnStrings(SmallRows) : Data.CreateStringHeavy(SmallRows);
                writer.WriteStrings(output, rows);
                var report = Check(writer.Name, scenario, output, SmallRows, Data.StringColumns);
                var expected = rows.Select(r => r.Select(s => (object?)s).ToArray()).ToArray();
                return report with { ValuesChecked = CompareValues(writer.Name, scenario, output, expected, null) };
            }
            default:
            {
                var rows = Data.CreateStyled(SmallRows);
                writer.WriteStyled(output, rows);
                var report = Check(writer.Name, scenario, output, SmallRows, Data.StyledHeaders.Length);
                // Row 0 is the header; a date is compared as its serial number, which is how xlsx stores it.
                var expected = new List<object?[]> { Data.StyledHeaders.Select(h => (object?)h).ToArray() };
                expected.AddRange(rows.Select(r => new object?[]
                {
                    r.OrderDate.ToOADate(), r.Created.ToOADate(), r.ShipDate.ToOADate(),
                    r.Net, r.Tax, r.Total, r.Margin, r.Discount,
                }));
                string[] formats =
                {
                    Data.StyledDateFormat, Data.DateTimeFormat, Data.StyledDateFormat,
                    Data.MoneyFormat, Data.MoneyFormat, Data.MoneyFormat, Data.PercentFormat, Data.PercentFormat,
                };
                int compared = CompareValues(writer.Name, scenario, output, expected.ToArray(), formats);
                return report with { ValuesChecked = compared, Formats = DistinctFormatsOfRow(output, 1) };
            }
        }
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

        var shape = Scenarios.Shape(scenario);
        int expectedRows = Scenarios.HasHeader(scenario) ? dataRows + 1 : dataRows;
        if (rows != expectedRows)
            throw new InvalidOperationException($"{library} {scenario}: expected {expectedRows} rows, found {rows}.");
        if (cellsInRow != columns)
            throw new InvalidOperationException($"{library} {scenario}: expected {columns} cells in the last row, found {cellsInRow}.");

        var styles = wbPart.WorkbookStylesPart?.Stylesheet;
        // Wide and string scenarios have no header and no formats to report.
        if (shape is ScenarioShape.Wide or ScenarioShape.Strings) return new GuardReport(rows, cellsInRow, false, "-", "-", "-", "-");
        // The styled scenario reports its formats through CompareValues; here only the bold header is looked at.
        if (shape == ScenarioShape.Styled) return new GuardReport(rows, cellsInRow, IsBold(styles, a1.Style), "-", "-", "-", "-");

        return new GuardReport(rows, cellsInRow,
            HeaderBold: IsBold(styles, a1.Style),
            PriceType: rowCells[4].Type ?? "n",
            PriceFormat: FormatOf(styles, rowCells[4].Style),
            DateType: rowCells[6].Type ?? "n",
            DateFormat: FormatOf(styles, rowCells[6].Style));
    }

    /// <summary>One cell as read back: its <c>t</c> attribute, its text (shared and inline strings resolved) and its style index.</summary>
    private readonly record struct ReadCell(string? Type, string Text, uint Style);

    /// <summary>
    /// Reads every cell of the first worksheet into a list of rows. Uses the Open XML SDK's DOM, which is fine for
    /// the small cases this is called with (a few hundred cells); the full-size checks use the SAX counting in
    /// <see cref="Check"/>.
    /// </summary>
    private static (List<List<ReadCell>> Rows, Stylesheet? Styles) ReadCells(MemoryStream output)
    {
        using var copy = new MemoryStream(output.GetBuffer(), 0, (int)output.Length, writable: false);
        using var doc = SpreadsheetDocument.Open(copy, false);
        var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("No workbook part.");
        var firstSheet = wbPart.Workbook!.Sheets!.Elements<Sheet>().First();
        var wsPart = (WorksheetPart)wbPart.GetPartById(firstSheet.Id!.Value!);

        // Resolve the shared string table once; indexing it with ElementAt per cell would be quadratic.
        var shared = wbPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>()
            .Select(i => i.InnerText).ToList() ?? new List<string>();

        var rows = new List<List<ReadCell>>();
        foreach (var row in wsPart.Worksheet!.Descendants<Row>())
        {
            var cells = new List<ReadCell>();
            foreach (var cell in row.Elements<Cell>())
            {
                string? type = cell.DataType?.InnerText;
                string text = type switch
                {
                    "s" => shared[int.Parse(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture)],
                    "inlineStr" => cell.InlineString?.InnerText ?? "",
                    _ => cell.CellValue?.Text ?? "",
                };
                cells.Add(new ReadCell(type, text, cell.StyleIndex?.Value ?? 0));
            }
            rows.Add(cells);
        }
        return (rows, wbPart.WorkbookStylesPart?.Stylesheet);
    }

    /// <summary>
    /// Compares every cell of <paramref name="output"/> with <paramref name="expected"/>: a string must come back
    /// as the same text in a string cell, a double must come back as the same number in a number cell, to within
    /// 1e-7 (a date serial counts in days, so that is under 9 milliseconds and still tells one second from the
    /// next; money and percentages are exact to far more digits than that). When
    /// <paramref name="formats"/> is given, row 0 is a header and every later cell must also carry that column's
    /// number format. Returns the number of cells compared.
    /// </summary>
    private static int CompareValues(string library, string scenario, MemoryStream output, object?[][] expected, string[]? formats)
    {
        var (rows, styles) = ReadCells(output);
        string where = $"{library} {scenario}";
        if (rows.Count != expected.Length)
            throw new InvalidOperationException($"{where}: expected {expected.Length} rows, read back {rows.Count}.");

        int compared = 0;
        var formatCache = new Dictionary<uint, string>();
        for (int r = 0; r < expected.Length; r++)
        {
            if (rows[r].Count != expected[r].Length)
                throw new InvalidOperationException($"{where}: row {r + 1} has {rows[r].Count} cells, expected {expected[r].Length}.");
            for (int c = 0; c < expected[r].Length; c++)
            {
                var actual = rows[r][c];
                string cell = $"{where}, row {r + 1} column {c + 1}";
                switch (expected[r][c])
                {
                    case string text:
                        if (actual.Type is not ("s" or "inlineStr" or "str"))
                            throw new InvalidOperationException($"{cell}: expected text, found a cell of type '{actual.Type ?? "n"}'.");
                        if (!string.Equals(actual.Text, text, StringComparison.Ordinal))
                            throw new InvalidOperationException($"{cell}: expected \"{text}\", read back \"{actual.Text}\".");
                        break;
                    case double number:
                        if (actual.Type is not (null or "n"))
                            throw new InvalidOperationException($"{cell}: expected a number, found a cell of type '{actual.Type}'.");
                        if (!double.TryParse(actual.Text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var value)
                            || Math.Abs(value - number) > 1e-7)
                            throw new InvalidOperationException($"{cell}: expected {number:R}, read back \"{actual.Text}\".");
                        break;
                    default:
                        throw new InvalidOperationException($"{cell}: the guard has no rule for expected value {expected[r][c]}.");
                }
                // Row 0 is the header of a styled sheet; its look (bold) is reported, not compared.
                if (formats != null && r > 0)
                {
                    if (!formatCache.TryGetValue(actual.Style, out var found))
                        formatCache[actual.Style] = found = NormalizeFormat(FormatOf(styles, actual.Style));
                    if (!string.Equals(found, formats[c], StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException($"{cell}: expected number format '{formats[c]}', found '{found}'.");
                }
                compared++;
            }
        }
        return compared;
    }

    /// <summary>The distinct number formats used by the cells of one row, for the <c>--verify</c> table.</summary>
    private static string DistinctFormatsOfRow(MemoryStream output, int rowIndex)
    {
        var (rows, styles) = ReadCells(output);
        return string.Join(" ; ", rows[rowIndex].Select(c => NormalizeFormat(FormatOf(styles, c.Style))).Distinct());
    }

    /// <summary>Maps the built-in format ids that mean the same as one of the scenario's formats to that format's code.</summary>
    private static string NormalizeFormat(string format) => format switch
    {
        "builtin:4" => "#,##0.00",
        "builtin:10" => "0.00%",
        _ => format,
    };

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
