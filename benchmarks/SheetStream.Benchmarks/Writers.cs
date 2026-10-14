using System.IO.Compression;
using ClosedXML.Excel;
using MiniExcelLibs;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using EP = OfficeOpenXml;
using SS = SheetStream;

namespace SheetStream.Benchmarks;

/// <summary>
/// One library's implementation of both scenarios. The benchmarks, the correctness guard and the peak-memory
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
}

/// <summary>Registry of every library under test, in the order they appear in reports.</summary>
public static class Writers
{
    /// <summary>All writers. SheetStream (default settings) comes first because it is the baseline.</summary>
    public static readonly ILibraryWriter[] All =
    {
        new SheetStreamWriter(CompressionLevel.Optimal),
        new SheetStreamWriter(CompressionLevel.Fastest),
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

    /// <summary>Creates a writer using the given zip compression level.</summary>
    public SheetStreamWriter(CompressionLevel level) => _level = level;

    /// <inheritdoc />
    public string Name => _level == CompressionLevel.Optimal ? "SheetStream" : "SheetStream-Fastest";

    private SS.PackageOptions Options => new SS.PackageOptions { CompressionLevel = _level };

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
}
