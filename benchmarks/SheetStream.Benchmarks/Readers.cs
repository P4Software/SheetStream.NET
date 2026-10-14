using System.Globalization;
using System.IO.Compression;
using ClosedXML.Excel;
using ExcelDataReader;
using MiniExcelLibs;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.XSSF.EventUserModel;
using NPOI.XSSF.Extractor;
using NPOI.XSSF.UserModel;
using EP = OfficeOpenXml;
using SS = SheetStream.Reading;

namespace SheetStream.Benchmarks;

/// <summary>
/// One library's way of reading the Long sheet back. The read benchmarks, the reader guard and the <c>--peak</c>
/// read children all call this same method, so what is verified is exactly what is timed.
/// </summary>
public interface ILibraryReader
{
    /// <summary>Short display name used on the command line and in result tables.</summary>
    string Name { get; }

    /// <summary>
    /// Reads every data row of the first sheet and returns <see cref="Checksum"/>'s figure for it: the sum of
    /// the <c>Id</c> column plus the number of non-empty data cells. Returning a value computed from every row
    /// is what stops the JIT from optimising the read away, and it is what the guard compares.
    /// </summary>
    long ReadLong(Stream input);
}

/// <summary>Registry of every reader under test, in the order they appear in reports.</summary>
public static class Readers
{
    /// <summary>All readers. SheetStream's untyped reader comes first because it is the baseline.</summary>
    public static readonly ILibraryReader[] All =
    {
        new SheetStreamRowsReader(),
        new SheetStreamTypedReader(),
        new MiniExcelTypedReader(),
        new MiniExcelDynamicReader(),
        new ExcelDataReaderReader(),
        new ClosedXmlReader(),
        new EPPlusReader(),
        new NpoiXssfReader(),
        new NpoiSaxReader(),
    };

    /// <summary>Finds a reader by its <see cref="ILibraryReader.Name"/> (case-insensitive).</summary>
    public static ILibraryReader Get(string name) =>
        All.FirstOrDefault(r => string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException($"Unknown reader '{name}'. Known: {string.Join(", ", All.Select(r => r.Name))}");
}

/// <summary>
/// The figure every reader must arrive at, and the helpers that compute it from the shapes each library
/// returns. It is defined over data rows only (the header is skipped) because the typed readers never see
/// the header row.
/// </summary>
public static class Checksum
{
    /// <summary>
    /// The checksum of a Long sheet with <paramref name="rows"/> data rows, computed from the generator's
    /// definition rather than from any reader: Ids are 1..n and every one of the 10 cells is written.
    /// </summary>
    public static long Expected(int rows) => (long)rows * (rows + 1) / 2 + (long)rows * Data.LongHeaders.Length;

    /// <summary>
    /// Contribution of one row read into the typed model. The seven value-typed members cannot be null, so
    /// they count as present (every cell is written, so the untyped readers count them too); the three string
    /// members are counted only when the library really set them.
    /// </summary>
    public static long OfTyped(LongRow r) =>
        r.Id + 7 + (r.Name != null ? 1 : 0) + (r.Category != null ? 1 : 0) + (r.Region != null ? 1 : 0);

    /// <summary>
    /// The <c>Id</c> cell as a long, whatever a library hands back for a numeric cell: SheetStream and EPPlus
    /// return <see cref="double"/>, MiniExcel's dynamic rows can return <see cref="double"/> or <see cref="int"/>,
    /// and NPOI's SAX handler delivers formatted text.
    /// </summary>
    public static long IdOf(object? value) => value switch
    {
        double d => (long)d,
        int i => i,
        long l => l,
        decimal m => (long)m,
        string s => long.Parse(s, NumberStyles.Integer, CultureInfo.InvariantCulture),
        _ => throw new InvalidOperationException($"Unexpected Id cell value '{value}' ({value?.GetType().Name ?? "null"})."),
    };
}

/// <summary>
/// Produces the input file for the read benchmarks: the Long scenario written once by SheetStream at its
/// default settings. The same bytes are handed to every reader through a fresh read-only
/// <see cref="MemoryStream"/>, so no reader pays for disk I/O and none can disturb the shared buffer.
/// </summary>
public static class ReadInput
{
    /// <summary>Writes <paramref name="rows"/> Long rows and returns the finished .xlsx bytes.</summary>
    public static byte[] Create(int rows)
    {
        using var output = new ReusableOutput(Scenarios.OutputCapacity(Scenarios.Long));
        new SheetStreamWriter(CompressionLevel.Optimal).WriteLong(output, Data.CreateLong(rows));
        return output.ToArray();
    }

    /// <summary>Opens a read-only stream over the shared bytes. Readers may dispose it freely.</summary>
    public static MemoryStream Open(byte[] file) => new(file, writable: false);
}

/// <summary>
/// Reader-side correctness guard: every reader must reproduce <see cref="Checksum.Expected"/> for a file whose
/// contents are known from the generator, so a reader cannot "win" by skipping rows, cells or the last row.
/// </summary>
public static class ReaderGuard
{
    /// <summary>Checks <paramref name="reader"/> on a small Long file and throws if the checksum is off.</summary>
    public static long CheckSmall(ILibraryReader reader) => Check(reader, ReadInput.Create(Guard.SmallLongRows), Guard.SmallLongRows);

    /// <summary>Reads <paramref name="file"/> with <paramref name="reader"/> and compares the checksum with the expected one.</summary>
    public static long Check(ILibraryReader reader, byte[] file, int rows)
    {
        using var input = ReadInput.Open(file);
        long actual = reader.ReadLong(input);
        long expected = Checksum.Expected(rows);
        if (actual != expected)
            throw new InvalidOperationException($"{reader.Name}: checksum {actual:N0} for {rows:N0} rows, expected {expected:N0}.");
        return actual;
    }
}

/// <summary>
/// SheetStream's untyped reader: <c>ExcelReader.Open(stream)</c> and <c>ReadRows</c>, which streams one
/// <c>ExcelRow</c> at a time (numbers as <see cref="double"/>, shared strings resolved, the date column as
/// <see cref="DateTime"/>). The header row is the first row yielded and is skipped.
/// </summary>
public sealed class SheetStreamRowsReader : ILibraryReader
{
    /// <inheritdoc />
    public string Name => "SheetStream-Rows";

    /// <inheritdoc />
    public long ReadLong(Stream input)
    {
        using var reader = SS.ExcelReader.Open(input, leaveOpen: true);
        long sum = 0;
        bool header = true;
        foreach (var row in reader.ReadRows("Data"))
        {
            if (header) { header = false; continue; }
            sum += Checksum.IdOf(row[0]);
            var values = row.Values;
            for (int c = 0; c < values.Count; c++) if (values[c] != null) sum++;
        }
        return sum;
    }
}

/// <summary>
/// SheetStream's typed reader: <c>Read&lt;LongRow&gt;</c> maps the header row to properties by name and
/// converts each cell to the property type (the date column arrives as a date-styled number and becomes a
/// <see cref="DateTime"/>).
/// </summary>
public sealed class SheetStreamTypedReader : ILibraryReader
{
    /// <inheritdoc />
    public string Name => "SheetStream-Typed";

    /// <inheritdoc />
    public long ReadLong(Stream input)
    {
        using var reader = SS.ExcelReader.Open(input, leaveOpen: true);
        long sum = 0;
        foreach (var row in reader.Read<LongRow>("Data")) sum += Checksum.OfTyped(row);
        return sum;
    }
}

/// <summary>
/// MiniExcel's typed path: <c>stream.Query&lt;LongRow&gt;()</c>, which maps the header row to the properties of
/// <see cref="LongRow"/> (the same class the MiniExcel writer uses). Everything is left at its defaults.
/// </summary>
public sealed class MiniExcelTypedReader : ILibraryReader
{
    /// <inheritdoc />
    public string Name => "MiniExcel-Typed";

    /// <inheritdoc />
    public long ReadLong(Stream input)
    {
        long sum = 0;
        foreach (var row in input.Query<LongRow>(sheetName: "Data")) sum += Checksum.OfTyped(row);
        return sum;
    }
}

/// <summary>
/// MiniExcel's dynamic path: <c>stream.Query(useHeaderRow: true)</c>, which yields one
/// <c>IDictionary&lt;string, object&gt;</c> per data row keyed by header caption. This is MiniExcel's
/// documented shape for reading without a model class.
/// </summary>
public sealed class MiniExcelDynamicReader : ILibraryReader
{
    /// <inheritdoc />
    public string Name => "MiniExcel-Dynamic";

    /// <inheritdoc />
    public long ReadLong(Stream input)
    {
        long sum = 0;
        foreach (IDictionary<string, object?> row in input.Query(useHeaderRow: true, sheetName: "Data"))
        {
            sum += Checksum.IdOf(row["Id"]);
            foreach (var value in row.Values) if (value != null) sum++;
        }
        return sum;
    }
}

/// <summary>
/// ExcelDataReader: <c>ExcelReaderFactory.CreateReader(stream)</c> and the <c>IDataReader</c>-style
/// <c>Read()</c> / <c>GetValue(i)</c> loop from its README. The first <c>Read()</c> returns the header row,
/// which is skipped. Empty cells come back as null in 3.x (the README's <c>IsDBNull</c> check also covers
/// the 4.x behaviour, where they are <see cref="DBNull"/>), so both are treated as empty.
/// </summary>
/// <remarks>
/// The README asks every .NET Core application to register <c>CodePagesEncodingProvider</c> once at start-up.
/// It is only needed to decode the legacy code pages of binary .xls files, which this suite never opens, but
/// it is the documented set-up step, so it is done once in the static constructor (outside any timing).
/// </remarks>
public sealed class ExcelDataReaderReader : ILibraryReader
{
    static ExcelDataReaderReader() =>
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

    /// <inheritdoc />
    public string Name => "ExcelDataReader";

    /// <inheritdoc />
    public long ReadLong(Stream input)
    {
        using var reader = ExcelReaderFactory.CreateReader(input);
        long sum = 0;
        bool header = true;
        while (reader.Read())
        {
            if (header) { header = false; continue; }
            sum += Checksum.IdOf(reader.GetValue(0));
            for (int c = 0; c < reader.FieldCount; c++)
            {
                var value = reader.GetValue(c);
                if (value != null && value is not DBNull) sum++;
            }
        }
        return sum;
    }
}

/// <summary>
/// ClosedXML: <c>new XLWorkbook(stream)</c> loads the whole workbook, then <c>RowsUsed()</c> / <c>CellsUsed()</c>
/// walk the cells that hold a value. Row 1 is the header and is skipped.
/// </summary>
public sealed class ClosedXmlReader : ILibraryReader
{
    /// <inheritdoc />
    public string Name => "ClosedXML";

    /// <inheritdoc />
    public long ReadLong(Stream input)
    {
        using var wb = new XLWorkbook(input);
        var ws = wb.Worksheet("Data");
        long sum = 0;
        foreach (var row in ws.RowsUsed())
        {
            if (row.RowNumber() == 1) continue;
            sum += (long)row.Cell(1).GetDouble();
            foreach (var cell in row.CellsUsed()) sum++;
        }
        return sum;
    }
}

/// <summary>
/// EPPlus 4.5.3.3: <c>new ExcelPackage(stream)</c> loads the workbook, then a <c>Cells[r, c].Value</c> loop over
/// <c>Dimension</c>, which is how the EPPlus 4 samples read a sheet. Each <c>Cells[r, c]</c> allocates an
/// <c>ExcelRange</c>; <c>GetValue&lt;T&gt;</c> goes through the same indexer, so there is no cheaper documented
/// path on the read side. Numbers come back as <see cref="double"/>.
/// </summary>
public sealed class EPPlusReader : ILibraryReader
{
    /// <inheritdoc />
    public string Name => "EPPlus 4.5.3.3";

    /// <inheritdoc />
    public long ReadLong(Stream input)
    {
        using var package = new EP.ExcelPackage(input);
        var ws = package.Workbook.Worksheets["Data"];
        var end = ws.Dimension.End;
        long sum = 0;
        for (int r = 2; r <= end.Row; r++)
        {
            sum += Checksum.IdOf(ws.Cells[r, 1].Value);
            for (int c = 1; c <= end.Column; c++) if (ws.Cells[r, c].Value != null) sum++;
        }
        return sum;
    }
}

/// <summary>
/// NPOI XSSF user model: <c>new XSSFWorkbook(stream)</c> loads the workbook, then <c>GetRow</c> / <c>GetCell</c>
/// over <c>LastRowNum</c>. A missing cell is null and a present-but-empty one is <c>CellType.Blank</c>; neither
/// occurs in this file, but both are treated as empty so the count means the same as for the other readers.
/// </summary>
public sealed class NpoiXssfReader : ILibraryReader
{
    /// <inheritdoc />
    public string Name => "NPOI XSSF";

    /// <inheritdoc />
    public long ReadLong(Stream input)
    {
        using var wb = new XSSFWorkbook(input);
        var sheet = wb.GetSheet("Data");
        long sum = 0;
        for (int r = 1; r <= sheet.LastRowNum; r++)
        {
            var row = sheet.GetRow(r);
            if (row == null) continue;
            sum += (long)row.GetCell(0).NumericCellValue;
            for (int c = 0; c < row.LastCellNum; c++)
            {
                var cell = row.GetCell(c);
                if (cell != null && cell.CellType != CellType.Blank) sum++;
            }
        }
        return sum;
    }
}

/// <summary>
/// NPOI's event (SAX) model, the port of Apache POI's <c>XSSFReader</c> + <c>XSSFSheetXMLHandler</c>: the
/// package is opened with <c>OPCPackage.Open(stream)</c>, the shared strings and styles are loaded once, and the
/// first sheet's XML is streamed through an <c>ISheetContentsHandler</c>. NPOI drives the handler through the
/// public <c>XSSFEventBasedExcelExtractor.ProcessSheet</c>, which is the only parser entry point it exposes.
/// </summary>
/// <remarks>
/// The handler is text-extraction oriented: <c>Cell</c> receives the <em>formatted</em> value as a string, so
/// <c>Id</c> is parsed back from text. That is inherent to the API and is a cost every SAX user of NPOI pays.
/// </remarks>
public sealed class NpoiSaxReader : ILibraryReader
{
    /// <inheritdoc />
    public string Name => "NPOI SAX";

    /// <inheritdoc />
    public long ReadLong(Stream input)
    {
        // OPCPackage is not IDisposable. Open(Stream) opens it read-write, so Close() would try to save it;
        // Revert() is the documented way to close a package without saving.
        var package = OPCPackage.Open(input);
        try
        {
            var reader = new XSSFReader(package);
            var strings = new ReadOnlySharedStringsTable(package);
            var styles = reader.StylesTable;
            var handler = new SumHandler();
            using var sheets = reader.GetSheetsData();
            if (!sheets.MoveNext()) throw new InvalidOperationException("The workbook has no sheets.");
            using (var sheet = sheets.Current)
            {
                new XSSFEventBasedExcelExtractor(package).ProcessSheet(handler, styles, null, strings, sheet);
            }
            return handler.Sum;
        }
        finally
        {
            package.Revert();
        }
    }

    /// <summary>Accumulates the checksum from the handler callbacks; row 0 is the header and is ignored.</summary>
    private sealed class SumHandler : XSSFSheetXMLHandler.ISheetContentsHandler
    {
        public long Sum;
        private bool _header;

        public void StartRow(int rowNum) => _header = rowNum == 0;

        public void EndRow(int rowNum) { }

        public void EndSheet() { }

        public void Cell(string cellReference, string formattedValue, XSSFComment comment)
        {
            if (_header || formattedValue == null) return;
            // Column A holds Id; the reference's first character is enough to recognise it.
            if (cellReference[0] == 'A' && (cellReference.Length == 1 || cellReference[1] is >= '0' and <= '9'))
                Sum += Checksum.IdOf(formattedValue);
            Sum++;
        }

        public void HeaderFooter(string text, bool isHeader, string tagName) { }
    }
}
