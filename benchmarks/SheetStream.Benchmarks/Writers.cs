using System.IO.Compression;
using ClosedXML.Excel;
using MiniExcelLibs;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using EP = OfficeOpenXml;
using SC = SpreadCheetah;
using SS = SheetStream;

namespace SheetStream.Benchmarks;

/// <summary>
/// One library's implementation of every scenario shape (long, wide, strings, styled). The benchmarks, the correctness guard and the peak-memory
/// mode all call these same methods, so what is verified is exactly what is timed.
/// </summary>
public interface ILibraryWriter
{
    /// <summary>Short display name used on the command line and in result tables.</summary>
    string Name { get; }

    /// <summary>Writes a header row (bold) plus one row per item of <paramref name="rows"/>.</summary>
    void WriteLong(Stream output, LongRow[] rows);

    /// <summary>Writes every value of <paramref name="grid"/>, no header row.</summary>
    void WriteWide(Stream output, double[,] grid);

    /// <summary>
    /// Writes <paramref name="rows"/> (one string array per row, all the same length) as plain text cells: no
    /// header row, no styles. Used by the SpreadCheetahOwn and StringHeavy scenarios.
    /// </summary>
    void WriteStrings(Stream output, string[][] rows);

    /// <summary>
    /// Writes a bold header row (<see cref="Data.StyledHeaders"/>) plus one row per item of <paramref name="rows"/>,
    /// with the date, date-time, money and percentage formats of the StyledDates scenario on every data cell.
    /// </summary>
    void WriteStyled(Stream output, StyledRow[] rows);
}

/// <summary>Registry of every library under test, in the order they appear in reports.</summary>
public static class Writers
{
    /// <summary>All writers. SheetStream (default settings) comes first because it is the baseline.</summary>
    public static readonly ILibraryWriter[] All =
    {
        new SheetStreamWriter(CompressionLevel.Optimal),
        new SheetStreamWriter(CompressionLevel.Fastest),
        // What SheetStream did before text could be stored inline: every text through the shared string table. Kept so the before/after is in one run.
        new SheetStreamWriter(CompressionLevel.Optimal, storage: SS.StringStorage.Shared),
        new SheetStreamWriter(CompressionLevel.Fastest, storage: SS.StringStorage.Shared),
        new SpreadCheetahWriter(SC.SpreadCheetahCompressionLevel.Fastest),
        new SpreadCheetahWriter(SC.SpreadCheetahCompressionLevel.Optimal),
        new EPPlusWriter(),
        new ClosedXmlWriter(),
        new MiniExcelWriter(),
        new NpoiXssfWriter(),
        new NpoiSxssfWriter(),
    };

    /// <summary>Finds a writer by its <see cref="ILibraryWriter.Name"/> (case-insensitive).</summary>
    public static ILibraryWriter Get(string name) =>
        All.FirstOrDefault(w => string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException($"Unknown library '{name}'. Known: {string.Join(", ", All.Select(w => w.Name))}");
}

/// <summary>
/// SheetStream: rows stream straight into the zip entry. Styles are created once and passed per cell, which is
/// the documented pattern. The "Fastest" variant only changes <see cref="PackageOptions.CompressionLevel"/>;
/// it is reported separately because every other library runs at its default compression.
/// </summary>
public sealed class SheetStreamWriter : ILibraryWriter
{
    private readonly CompressionLevel _level;
    private readonly bool _writeCellReferences;
    private readonly SS.StringStorage _storage;

    /// <summary>Creates a writer using the given zip compression level.</summary>
    /// <param name="level">Zip compression level.</param>
    /// <param name="writeCellReferences">
    /// Writes the address (<c>r="B2"</c>) on every cell. The writer scenarios leave it off, SheetStream's default,
    /// because that is what the writing benchmark measures. The file the reader benchmarks read turns it on: some
    /// readers under test (NPOI's SAX reader) need the address of every cell, and a reader table should compare
    /// readers on a file they can all read, not penalise one for a writer-side default.
    /// </param>
    /// <param name="storage">How text cells are stored. The default, <see cref="SS.StringStorage.Auto"/>, is SheetStream's default setting.</param>
    public SheetStreamWriter(CompressionLevel level, bool writeCellReferences = false, SS.StringStorage storage = SS.StringStorage.Auto)
    {
        _level = level;
        _writeCellReferences = writeCellReferences;
        _storage = storage;
    }

    /// <inheritdoc />
    public string Name => (_storage == SS.StringStorage.Shared ? "SheetStream-Shared" : "SheetStream") + (_level == CompressionLevel.Optimal ? "" : "-Fastest");

    private SS.PackageOptions Options => new SS.PackageOptions { CompressionLevel = _level, WriteCellReferences = _writeCellReferences, StringStorage = _storage };

    /// <inheritdoc />
    public void WriteLong(Stream output, LongRow[] rows)
    {
        using var package = new SS.ExcelPackage(output, leaveOpen: true, Options);
        var workbook = package.Workbook;
        var sheet = workbook.Worksheets.Add("Data");
        var bold = workbook.Styles.Create(o => o.Font.Bold = true);
        var number = workbook.Styles.Create(o => o.NumberFormat = Data.NumberFormat);
        var date = workbook.Styles.Create(o => o.NumberFormat = Data.DateFormat);

        var header = sheet.Row(1);
        foreach (var h in Data.LongHeaders) header.WriteCell(h, bold);

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            sheet.Row(i + 2)
                .WriteCell(r.Id)
                .WriteCell(r.Name)
                .WriteCell(r.Category)
                .WriteCell(r.Quantity)
                .WriteCell(r.Price, number)
                .WriteCell(r.Amount, number)
                .WriteCell(r.Date, date)
                .WriteCell(r.Region)
                .WriteCell(r.Score, number)
                .WriteCell(r.Code);
        }
        package.Save();
    }

    /// <inheritdoc />
    public void WriteWide(Stream output, double[,] grid)
    {
        using var package = new SS.ExcelPackage(output, leaveOpen: true, Options);
        var sheet = package.Workbook.Worksheets.Add("Data");
        int rows = grid.GetLength(0), cols = grid.GetLength(1);
        for (int r = 0; r < rows; r++)
        {
            var row = sheet.Row(r + 1);
            for (int c = 0; c < cols; c++) row.WriteCell(grid[r, c]);
        }
        package.Save();
    }

    /// <inheritdoc />
    public void WriteStrings(Stream output, string[][] rows)
    {
        // Default options on purpose (StringStorage.Auto): SheetStream decides per column whether text is shared or inline.
        using var package = new SS.ExcelPackage(output, leaveOpen: true, Options);
        var sheet = package.Workbook.Worksheets.Add("Data");
        for (int r = 0; r < rows.Length; r++)
        {
            var row = sheet.Row(r + 1);
            var values = rows[r];
            for (int c = 0; c < values.Length; c++) row.WriteCell(values[c]);
        }
        package.Save();
    }

    /// <inheritdoc />
    public void WriteStyled(Stream output, StyledRow[] rows)
    {
        using var package = new SS.ExcelPackage(output, leaveOpen: true, Options);
        var workbook = package.Workbook;
        var sheet = workbook.Worksheets.Add("Data");
        var bold = workbook.Styles.Create(o => o.Font.Bold = true);
        var date = workbook.Styles.Create(o => o.NumberFormat = Data.StyledDateFormat);
        var dateTime = workbook.Styles.Create(o => o.NumberFormat = Data.DateTimeFormat);
        var money = workbook.Styles.Create(o => o.NumberFormat = Data.MoneyFormat);
        var percent = workbook.Styles.Create(o => o.NumberFormat = Data.PercentFormat);

        var header = sheet.Row(1);
        foreach (var h in Data.StyledHeaders) header.WriteCell(h, bold);

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            sheet.Row(i + 2)
                .WriteCell(r.OrderDate, date)
                .WriteCell(r.Created, dateTime)
                .WriteCell(r.ShipDate, date)
                .WriteCell(r.Net, money)
                .WriteCell(r.Tax, money)
                .WriteCell(r.Total, money)
                .WriteCell(r.Margin, percent)
                .WriteCell(r.Discount, percent);
        }
        package.Save();
    }
}

/// <summary>
/// SpreadCheetah: a forward-only streaming writer, the closest design to SheetStream in this suite. Rows are
/// arrays of cell structs; one array is reused for every row, which is the pattern its README recommends, and
/// styles are added once and passed per cell. Its API is asynchronous only, so each scenario runs as one async
/// method that is waited on once at the end: the output is a memory buffer, so every await completes
/// synchronously and no thread is blocked on I/O.
/// </summary>
/// <remarks>
/// SpreadCheetah's default zip level is Fastest, where SheetStream and the other libraries default to Optimal.
/// It is therefore measured twice, like SheetStream: at its own default, and at Optimal, so that each
/// SpreadCheetah row has a SheetStream row at the same compression level to be read against.
/// </remarks>
public sealed class SpreadCheetahWriter : ILibraryWriter
{
    private readonly SC.SpreadCheetahCompressionLevel _level;

    /// <summary>Creates a writer using the given zip compression level.</summary>
    public SpreadCheetahWriter(SC.SpreadCheetahCompressionLevel level) => _level = level;

    /// <inheritdoc />
    public string Name => _level == SC.SpreadCheetahCompressionLevel.Fastest ? "SpreadCheetah" : "SpreadCheetah-Optimal";

    private SC.SpreadCheetahOptions Options => new SC.SpreadCheetahOptions { CompressionLevel = _level };

    /// <inheritdoc />
    public void WriteLong(Stream output, LongRow[] rows) => WriteLongAsync(output, rows).GetAwaiter().GetResult();

    /// <inheritdoc />
    public void WriteWide(Stream output, double[,] grid) => WriteWideAsync(output, grid).GetAwaiter().GetResult();

    private async Task WriteLongAsync(Stream output, LongRow[] rows)
    {
        await using var spreadsheet = await SC.Spreadsheet.CreateNewAsync(output, Options);
        await spreadsheet.StartWorksheetAsync("Data");

        var boldStyle = new SC.Styling.Style();
        boldStyle.Font.Bold = true;
        var bold = spreadsheet.AddStyle(boldStyle);
        var number = spreadsheet.AddStyle(new SC.Styling.Style { Format = SC.Styling.NumberFormat.Custom(Data.NumberFormat) });
        var date = spreadsheet.AddStyle(new SC.Styling.Style { Format = SC.Styling.NumberFormat.Custom(Data.DateFormat) });

        // One array for the whole sheet: AddRowAsync writes the row out before it returns, so the cells can be
        // overwritten for the next row.
        var cells = new SC.Cell[Data.LongHeaders.Length];
        for (int c = 0; c < cells.Length; c++) cells[c] = new SC.Cell(Data.LongHeaders[c], bold);
        await spreadsheet.AddRowAsync(cells);

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            cells[0] = new SC.Cell(r.Id);
            cells[1] = new SC.Cell(r.Name);
            cells[2] = new SC.Cell(r.Category);
            cells[3] = new SC.Cell(r.Quantity);
            cells[4] = new SC.Cell(r.Price, number);
            cells[5] = new SC.Cell(r.Amount, number);
            cells[6] = new SC.Cell(r.Date, date);
            cells[7] = new SC.Cell(r.Region);
            cells[8] = new SC.Cell(r.Score, number);
            cells[9] = new SC.Cell(r.Code);
            await spreadsheet.AddRowAsync(cells);
        }
        await spreadsheet.FinishAsync();
    }

    private async Task WriteWideAsync(Stream output, double[,] grid)
    {
        await using var spreadsheet = await SC.Spreadsheet.CreateNewAsync(output, Options);
        await spreadsheet.StartWorksheetAsync("Data");
        int rows = grid.GetLength(0), cols = grid.GetLength(1);

        // DataCell is the unstyled cell type, the cheapest one SpreadCheetah offers; reused for every row.
        var cells = new SC.DataCell[cols];
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++) cells[c] = new SC.DataCell(grid[r, c]);
            await spreadsheet.AddRowAsync(cells);
        }
        await spreadsheet.FinishAsync();
    }

    /// <inheritdoc />
    public void WriteStrings(Stream output, string[][] rows) => WriteStringsAsync(output, rows).GetAwaiter().GetResult();

    /// <inheritdoc />
    public void WriteStyled(Stream output, StyledRow[] rows) => WriteStyledAsync(output, rows).GetAwaiter().GetResult();

    private async Task WriteStringsAsync(Stream output, string[][] rows)
    {
        await using var spreadsheet = await SC.Spreadsheet.CreateNewAsync(output, Options);
        await spreadsheet.StartWorksheetAsync("Data");

        // The way SpreadCheetah recommends (and the way its own benchmark does it): one cell array reused for every
        // row, and DataCell, its unstyled and cheapest cell type, because no cell here carries a style.
        var cells = new SC.DataCell[rows.Length == 0 ? 0 : rows[0].Length];
        for (int r = 0; r < rows.Length; r++)
        {
            var values = rows[r];
            for (int c = 0; c < cells.Length; c++) cells[c] = new SC.DataCell(values[c]);
            await spreadsheet.AddRowAsync(cells);
        }
        await spreadsheet.FinishAsync();
    }

    private async Task WriteStyledAsync(Stream output, StyledRow[] rows)
    {
        await using var spreadsheet = await SC.Spreadsheet.CreateNewAsync(output, Options);
        await spreadsheet.StartWorksheetAsync("Data");

        var boldStyle = new SC.Styling.Style();
        boldStyle.Font.Bold = true;
        var bold = spreadsheet.AddStyle(boldStyle);
        var date = spreadsheet.AddStyle(new SC.Styling.Style { Format = SC.Styling.NumberFormat.Custom(Data.StyledDateFormat) });
        var dateTime = spreadsheet.AddStyle(new SC.Styling.Style { Format = SC.Styling.NumberFormat.Custom(Data.DateTimeFormat) });
        var money = spreadsheet.AddStyle(new SC.Styling.Style { Format = SC.Styling.NumberFormat.Custom(Data.MoneyFormat) });
        var percent = spreadsheet.AddStyle(new SC.Styling.Style { Format = SC.Styling.NumberFormat.Custom(Data.PercentFormat) });

        // StyledCell is SpreadCheetah's cell type for cells that always carry a style (every cell of this sheet
        // does), so it is the cheapest choice here; one array is reused for the header and for every row.
        var cells = new SC.StyledCell[Data.StyledHeaders.Length];
        for (int c = 0; c < cells.Length; c++) cells[c] = new SC.StyledCell(Data.StyledHeaders[c], bold);
        await spreadsheet.AddRowAsync(cells);

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            cells[0] = new SC.StyledCell(r.OrderDate, date);
            cells[1] = new SC.StyledCell(r.Created, dateTime);
            cells[2] = new SC.StyledCell(r.ShipDate, date);
            cells[3] = new SC.StyledCell(r.Net, money);
            cells[4] = new SC.StyledCell(r.Tax, money);
            cells[5] = new SC.StyledCell(r.Total, money);
            cells[6] = new SC.StyledCell(r.Margin, percent);
            cells[7] = new SC.StyledCell(r.Discount, percent);
            await spreadsheet.AddRowAsync(cells);
        }
        await spreadsheet.FinishAsync();
    }
}

/// <summary>
/// EPPlus 4.5.3.3 (the last LGPL release). Values go in through <c>Worksheet.SetValue(row, col, value)</c>,
/// which skips the <c>ExcelRange</c> allocation that <c>Cells[r, c].Value</c> makes per call; formats are then
/// applied once per column range, as the EPPlus samples do. <c>SetValue</c> takes <c>object</c>, so each value
/// is boxed; that is inherent to the EPPlus API (LoadFromArrays boxes too).
/// </summary>
public sealed class EPPlusWriter : ILibraryWriter
{
    /// <inheritdoc />
    public string Name => "EPPlus 4.5.3.3";

    /// <inheritdoc />
    public void WriteLong(Stream output, LongRow[] rows)
    {
        using var package = new EP.ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("Data");
        for (int c = 0; c < Data.LongHeaders.Length; c++) ws.SetValue(1, c + 1, Data.LongHeaders[c]);
        ws.Cells[1, 1, 1, Data.LongHeaders.Length].Style.Font.Bold = true;

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            int row = i + 2;
            ws.SetValue(row, 1, r.Id);
            ws.SetValue(row, 2, r.Name);
            ws.SetValue(row, 3, r.Category);
            ws.SetValue(row, 4, r.Quantity);
            ws.SetValue(row, 5, r.Price);
            ws.SetValue(row, 6, r.Amount);
            ws.SetValue(row, 7, r.Date);
            ws.SetValue(row, 8, r.Region);
            ws.SetValue(row, 9, r.Score);
            ws.SetValue(row, 10, r.Code);
        }

        int last = rows.Length + 1;
        if (rows.Length > 0)
        {
            ws.Cells[2, 5, last, 6].Style.Numberformat.Format = Data.NumberFormat;
            ws.Cells[2, 7, last, 7].Style.Numberformat.Format = Data.DateFormat;
            ws.Cells[2, 9, last, 9].Style.Numberformat.Format = Data.NumberFormat;
        }
        package.SaveAs(output);
    }

    /// <inheritdoc />
    public void WriteWide(Stream output, double[,] grid)
    {
        using var package = new EP.ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("Data");
        int rows = grid.GetLength(0), cols = grid.GetLength(1);
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                ws.SetValue(r + 1, c + 1, grid[r, c]);
        package.SaveAs(output);
    }

    /// <inheritdoc />
    public void WriteStrings(Stream output, string[][] rows)
    {
        using var package = new EP.ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("Data");
        for (int r = 0; r < rows.Length; r++)
        {
            var values = rows[r];
            for (int c = 0; c < values.Length; c++) ws.SetValue(r + 1, c + 1, values[c]);
        }
        package.SaveAs(output);
    }

    /// <inheritdoc />
    public void WriteStyled(Stream output, StyledRow[] rows)
    {
        using var package = new EP.ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("Data");
        for (int c = 0; c < Data.StyledHeaders.Length; c++) ws.SetValue(1, c + 1, Data.StyledHeaders[c]);
        ws.Cells[1, 1, 1, Data.StyledHeaders.Length].Style.Font.Bold = true;

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            int row = i + 2;
            ws.SetValue(row, 1, r.OrderDate);
            ws.SetValue(row, 2, r.Created);
            ws.SetValue(row, 3, r.ShipDate);
            ws.SetValue(row, 4, r.Net);
            ws.SetValue(row, 5, r.Tax);
            ws.SetValue(row, 6, r.Total);
            ws.SetValue(row, 7, r.Margin);
            ws.SetValue(row, 8, r.Discount);
        }

        // Formats once per column range, as the Long scenario does (the EPPlus samples' way).
        int last = rows.Length + 1;
        if (rows.Length > 0)
        {
            ws.Cells[2, 1, last, 1].Style.Numberformat.Format = Data.StyledDateFormat;
            ws.Cells[2, 2, last, 2].Style.Numberformat.Format = Data.DateTimeFormat;
            ws.Cells[2, 3, last, 3].Style.Numberformat.Format = Data.StyledDateFormat;
            ws.Cells[2, 4, last, 6].Style.Numberformat.Format = Data.MoneyFormat;
            ws.Cells[2, 7, last, 8].Style.Numberformat.Format = Data.PercentFormat;
        }
        package.SaveAs(output);
    }
}

/// <summary>
/// ClosedXML: cell loop assigning the typed <c>XLCellValue</c> (a struct, so no boxing), then one style call
/// per column range. ClosedXML keeps the whole workbook in memory and serialises it through the Open XML SDK.
/// </summary>
public sealed class ClosedXmlWriter : ILibraryWriter
{
    /// <inheritdoc />
    public string Name => "ClosedXML";

    /// <inheritdoc />
    public void WriteLong(Stream output, LongRow[] rows)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Data");
        for (int c = 0; c < Data.LongHeaders.Length; c++) ws.Cell(1, c + 1).Value = Data.LongHeaders[c];
        ws.Range(1, 1, 1, Data.LongHeaders.Length).Style.Font.Bold = true;

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            int row = i + 2;
            ws.Cell(row, 1).Value = r.Id;
            ws.Cell(row, 2).Value = r.Name;
            ws.Cell(row, 3).Value = r.Category;
            ws.Cell(row, 4).Value = r.Quantity;
            ws.Cell(row, 5).Value = r.Price;
            ws.Cell(row, 6).Value = r.Amount;
            ws.Cell(row, 7).Value = r.Date;
            ws.Cell(row, 8).Value = r.Region;
            ws.Cell(row, 9).Value = r.Score;
            ws.Cell(row, 10).Value = r.Code;
        }

        int last = rows.Length + 1;
        if (rows.Length > 0)
        {
            ws.Range(2, 5, last, 6).Style.NumberFormat.Format = Data.NumberFormat;
            ws.Range(2, 7, last, 7).Style.NumberFormat.Format = Data.DateFormat;
            ws.Range(2, 9, last, 9).Style.NumberFormat.Format = Data.NumberFormat;
        }
        wb.SaveAs(output);
    }

    /// <inheritdoc />
    public void WriteWide(Stream output, double[,] grid)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Data");
        int rows = grid.GetLength(0), cols = grid.GetLength(1);
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                ws.Cell(r + 1, c + 1).Value = grid[r, c];
        wb.SaveAs(output);
    }

    /// <inheritdoc />
    public void WriteStrings(Stream output, string[][] rows)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Data");
        for (int r = 0; r < rows.Length; r++)
        {
            var values = rows[r];
            for (int c = 0; c < values.Length; c++) ws.Cell(r + 1, c + 1).Value = values[c];
        }
        wb.SaveAs(output);
    }

    /// <inheritdoc />
    public void WriteStyled(Stream output, StyledRow[] rows)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Data");
        for (int c = 0; c < Data.StyledHeaders.Length; c++) ws.Cell(1, c + 1).Value = Data.StyledHeaders[c];
        ws.Range(1, 1, 1, Data.StyledHeaders.Length).Style.Font.Bold = true;

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            int row = i + 2;
            ws.Cell(row, 1).Value = r.OrderDate;
            ws.Cell(row, 2).Value = r.Created;
            ws.Cell(row, 3).Value = r.ShipDate;
            ws.Cell(row, 4).Value = r.Net;
            ws.Cell(row, 5).Value = r.Tax;
            ws.Cell(row, 6).Value = r.Total;
            ws.Cell(row, 7).Value = r.Margin;
            ws.Cell(row, 8).Value = r.Discount;
        }

        // Formats once per column range, as the Long scenario does.
        int last = rows.Length + 1;
        if (rows.Length > 0)
        {
            ws.Range(2, 1, last, 1).Style.NumberFormat.Format = Data.StyledDateFormat;
            ws.Range(2, 2, last, 2).Style.NumberFormat.Format = Data.DateTimeFormat;
            ws.Range(2, 3, last, 3).Style.NumberFormat.Format = Data.StyledDateFormat;
            ws.Range(2, 4, last, 6).Style.NumberFormat.Format = Data.MoneyFormat;
            ws.Range(2, 7, last, 8).Style.NumberFormat.Format = Data.PercentFormat;
        }
        wb.SaveAs(output);
    }
}

/// <summary>
/// MiniExcel: the documented fast path, <c>MiniExcel.SaveAs</c> over an <c>IEnumerable</c>. Long uses the typed
/// <see cref="LongRow"/> class with <c>[ExcelColumn(Format = ...)]</c>; Wide uses one dictionary per row
/// (MiniExcel's documented shape for dynamic columns), which boxes all 1,000 doubles of each row.
/// Differences that MiniExcel does not let us switch off without leaving its defaults are listed in the README
/// (its own default header style, an auto-filter and a frozen first row on the Long sheet).
/// </summary>
public sealed class MiniExcelWriter : ILibraryWriter
{
    /// <inheritdoc />
    public string Name => "MiniExcel";

    /// <inheritdoc />
    public void WriteLong(Stream output, LongRow[] rows) =>
        output.SaveAs(rows, printHeader: true, sheetName: "Data");

    /// <inheritdoc />
    public void WriteWide(Stream output, double[,] grid) =>
        output.SaveAs(WideRows(grid), printHeader: false, sheetName: "Data");

    /// <inheritdoc />
    public void WriteStrings(Stream output, string[][] rows) =>
        output.SaveAs(StringRows(rows), printHeader: false, sheetName: "Data");

    /// <inheritdoc />
    public void WriteStyled(Stream output, StyledRow[] rows) =>
        // The typed path, as for Long: the formats come from [ExcelColumn(Format = ...)] on StyledRow. The header
        // captions are the property names, which Data.StyledHeaders matches.
        output.SaveAs(rows, printHeader: true, sheetName: "Data");

    /// <summary>
    /// One dictionary per sheet, overwritten for every row. MiniExcel's streaming writer consumes a row before it
    /// asks for the next, so this is safe, and it is the cheapest way to feed it: MiniExcel is not charged for
    /// allocating a dictionary per row that a caller who cares about speed would not allocate either. (Wide keeps
    /// the one-dictionary-per-row shape the original scenarios used; the --verify guard proves this one still
    /// writes every row correctly.)
    /// </summary>
    private static IEnumerable<Dictionary<string, object>> StringRows(string[][] rows)
    {
        if (rows.Length == 0) yield break;
        int cols = rows[0].Length;
        var keys = new string[cols];
        for (int c = 0; c < cols; c++) keys[c] = "C" + (c + 1);
        var row = new Dictionary<string, object>(cols);
        for (int r = 0; r < rows.Length; r++)
        {
            var values = rows[r];
            for (int c = 0; c < cols; c++) row[keys[c]] = values[c];
            yield return row;
        }
    }

    /// <summary>Column keys are built once; MiniExcel only uses them for headers, which are off here.</summary>
    private static IEnumerable<Dictionary<string, object>> WideRows(double[,] grid)
    {
        int rows = grid.GetLength(0), cols = grid.GetLength(1);
        var keys = new string[cols];
        for (int c = 0; c < cols; c++) keys[c] = "C" + (c + 1);
        for (int r = 0; r < rows; r++)
        {
            var row = new Dictionary<string, object>(cols);
            for (int c = 0; c < cols; c++) row[keys[c]] = grid[r, c];
            yield return row;
        }
    }
}

/// <summary>
/// NPOI XSSF (the in-memory user model): <c>CreateRow</c> / <c>CreateCell</c> / <c>SetCellValue</c> with three
/// cell styles created once up front, as the NPOI docs recommend (one style per cell would hit Excel's limit).
/// </summary>
public sealed class NpoiXssfWriter : ILibraryWriter
{
    /// <inheritdoc />
    public string Name => "NPOI XSSF";

    /// <inheritdoc />
    public void WriteLong(Stream output, LongRow[] rows)
    {
        using var wb = new XSSFWorkbook();
        NpoiShared.WriteLong(wb, rows);
        wb.Write(output, leaveOpen: true);
    }

    /// <inheritdoc />
    public void WriteWide(Stream output, double[,] grid)
    {
        using var wb = new XSSFWorkbook();
        NpoiShared.WriteWide(wb, grid);
        wb.Write(output, leaveOpen: true);
    }

    /// <inheritdoc />
    public void WriteStrings(Stream output, string[][] rows)
    {
        using var wb = new XSSFWorkbook();
        NpoiShared.WriteStrings(wb, rows);
        wb.Write(output, leaveOpen: true);
    }

    /// <inheritdoc />
    public void WriteStyled(Stream output, StyledRow[] rows)
    {
        using var wb = new XSSFWorkbook();
        NpoiShared.WriteStyled(wb, rows);
        wb.Write(output, leaveOpen: true);
    }
}

/// <summary>
/// NPOI SXSSF (the streaming port of Apache POI's SXSSF): keeps a 100-row window in memory and flushes older
/// rows to a temporary XML file on disk, which is copied into the package on <c>Write</c>. Its time therefore
/// includes that temp-file I/O. <c>Dispose</c> deletes the temp files.
/// </summary>
public sealed class NpoiSxssfWriter : ILibraryWriter
{
    /// <inheritdoc />
    public string Name => "NPOI SXSSF";

    /// <inheritdoc />
    public void WriteLong(Stream output, LongRow[] rows)
    {
        using var wb = new SXSSFWorkbook(100);
        NpoiShared.WriteLong(wb, rows);
        wb.Write(output, leaveOpen: true);
    }

    /// <inheritdoc />
    public void WriteWide(Stream output, double[,] grid)
    {
        using var wb = new SXSSFWorkbook(100);
        NpoiShared.WriteWide(wb, grid);
        wb.Write(output, leaveOpen: true);
    }

    /// <inheritdoc />
    public void WriteStrings(Stream output, string[][] rows)
    {
        using var wb = new SXSSFWorkbook(100);
        NpoiShared.WriteStrings(wb, rows);
        wb.Write(output, leaveOpen: true);
    }

    /// <inheritdoc />
    public void WriteStyled(Stream output, StyledRow[] rows)
    {
        using var wb = new SXSSFWorkbook(100);
        NpoiShared.WriteStyled(wb, rows);
        wb.Write(output, leaveOpen: true);
    }
}

/// <summary>The NPOI user-model code is identical for XSSF and SXSSF; only the workbook class differs.</summary>
internal static class NpoiShared
{
    public static void WriteLong(IWorkbook wb, LongRow[] rows)
    {
        var sheet = wb.CreateSheet("Data");
        var format = wb.CreateDataFormat();

        var font = wb.CreateFont();
        font.IsBold = true;
        var bold = wb.CreateCellStyle();
        bold.SetFont(font);
        var number = wb.CreateCellStyle();
        number.DataFormat = format.GetFormat(Data.NumberFormat);
        var date = wb.CreateCellStyle();
        date.DataFormat = format.GetFormat(Data.DateFormat);

        var header = sheet.CreateRow(0);
        for (int c = 0; c < Data.LongHeaders.Length; c++)
        {
            var cell = header.CreateCell(c);
            cell.SetCellValue(Data.LongHeaders[c]);
            cell.CellStyle = bold;
        }

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            var row = sheet.CreateRow(i + 1);
            row.CreateCell(0).SetCellValue(r.Id);
            row.CreateCell(1).SetCellValue(r.Name);
            row.CreateCell(2).SetCellValue(r.Category);
            row.CreateCell(3).SetCellValue(r.Quantity);
            var price = row.CreateCell(4); price.SetCellValue(r.Price); price.CellStyle = number;
            var amount = row.CreateCell(5); amount.SetCellValue(r.Amount); amount.CellStyle = number;
            var d = row.CreateCell(6); d.SetCellValue(r.Date); d.CellStyle = date;
            row.CreateCell(7).SetCellValue(r.Region);
            var score = row.CreateCell(8); score.SetCellValue(r.Score); score.CellStyle = number;
            row.CreateCell(9).SetCellValue(r.Code);
        }
    }

    public static void WriteWide(IWorkbook wb, double[,] grid)
    {
        var sheet = wb.CreateSheet("Data");
        int rows = grid.GetLength(0), cols = grid.GetLength(1);
        for (int r = 0; r < rows; r++)
        {
            var row = sheet.CreateRow(r);
            for (int c = 0; c < cols; c++) row.CreateCell(c).SetCellValue(grid[r, c]);
        }
    }

    public static void WriteStrings(IWorkbook wb, string[][] rows)
    {
        var sheet = wb.CreateSheet("Data");
        for (int r = 0; r < rows.Length; r++)
        {
            var row = sheet.CreateRow(r);
            var values = rows[r];
            for (int c = 0; c < values.Length; c++) row.CreateCell(c).SetCellValue(values[c]);
        }
    }

    public static void WriteStyled(IWorkbook wb, StyledRow[] rows)
    {
        var sheet = wb.CreateSheet("Data");
        var format = wb.CreateDataFormat();

        // Five styles created once and shared by every cell, as the NPOI docs recommend.
        var font = wb.CreateFont();
        font.IsBold = true;
        var bold = wb.CreateCellStyle();
        bold.SetFont(font);
        var date = wb.CreateCellStyle();
        date.DataFormat = format.GetFormat(Data.StyledDateFormat);
        var dateTime = wb.CreateCellStyle();
        dateTime.DataFormat = format.GetFormat(Data.DateTimeFormat);
        var money = wb.CreateCellStyle();
        money.DataFormat = format.GetFormat(Data.MoneyFormat);
        var percent = wb.CreateCellStyle();
        percent.DataFormat = format.GetFormat(Data.PercentFormat);

        var header = sheet.CreateRow(0);
        for (int c = 0; c < Data.StyledHeaders.Length; c++)
        {
            var cell = header.CreateCell(c);
            cell.SetCellValue(Data.StyledHeaders[c]);
            cell.CellStyle = bold;
        }

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            var row = sheet.CreateRow(i + 1);
            var a = row.CreateCell(0); a.SetCellValue(r.OrderDate); a.CellStyle = date;
            var b = row.CreateCell(1); b.SetCellValue(r.Created); b.CellStyle = dateTime;
            var d = row.CreateCell(2); d.SetCellValue(r.ShipDate); d.CellStyle = date;
            var net = row.CreateCell(3); net.SetCellValue(r.Net); net.CellStyle = money;
            var tax = row.CreateCell(4); tax.SetCellValue(r.Tax); tax.CellStyle = money;
            var total = row.CreateCell(5); total.SetCellValue(r.Total); total.CellStyle = money;
            var margin = row.CreateCell(6); margin.SetCellValue(r.Margin); margin.CellStyle = percent;
            var discount = row.CreateCell(7); discount.SetCellValue(r.Discount); discount.CellStyle = percent;
        }
    }
}
