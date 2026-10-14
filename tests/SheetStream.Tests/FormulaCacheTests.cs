// The Open XML SDK exposes nearly every member as nullable; in test assertions a null is a failure anyway.
#pragma warning disable CS8602, CS8603, CS8604
using System.IO.Compression;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Internal;
using SheetStream.Reading;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SheetStream.Tests;

/// <summary>
/// Formula result caching: the evaluator's grammar and limits in isolation, then end-to-end files read back
/// through ExcelReader, which only ever sees the cached <c>&lt;v&gt;</c>.
/// </summary>
public class FormulaCacheTests
{
    // ----- Evaluator -------------------------------------------------------------------------------------------

    /// <summary>
    /// A small sheet already written up to row 9: B1:B5 = 10..50, C1:C5 = 1.5, 2.5, "text", (empty), 4.5,
    /// D1 = "x". The formula under test is written at A10 unless a test says otherwise.
    /// </summary>
    private static FormulaValueStore SampleStore()
    {
        var store = new FormulaValueStore();
        for (var r = 1; r <= 5; r++) store.SetNumber(r, 2, r * 10);
        store.SetNumber(1, 3, 1.5);
        store.SetNumber(2, 3, 2.5);
        store.SetText(3, 3);
        store.SetNumber(5, 3, 4.5);
        store.SetText(1, 4);
        return store;
    }

    private static double? Eval(string formula, FormulaValueStore? store = null, int row = 10, int column = 1, string sheet = "Data")
        => FormulaEvaluator.TryEvaluate(formula, store ?? SampleStore(), row, column, sheet, out var v) ? v : (double?)null;

    [Theory]
    [InlineData("1+2*3", 7)]
    [InlineData("(1+2)*3", 9)]
    [InlineData("10-4-3", 3)]                 // Left associative.
    [InlineData("10/4", 2.5)]
    [InlineData("2*-3", -6)]
    [InlineData("-B1", -10)]
    [InlineData("--B1", 10)]
    [InlineData("50%", 0.5)]
    [InlineData("B1*21%", 2.1)]
    [InlineData("1E3+1", 1001)]
    [InlineData(".5+.25", 0.75)]
    [InlineData(" 1 +  2 ", 3)]
    [InlineData("B1 * B2 / B3", 20.0 / 3)]
    [InlineData("C4 + 1", 1)]                 // An empty cell is 0 in arithmetic.
    [InlineData("$B$1+B$2+$B3", 60)]
    [InlineData("b1+b2", 30)]                 // References are case-insensitive.
    [InlineData("SUM(B1:B5)", 150)]
    [InlineData("SUM($B$1:$B$5)", 150)]
    [InlineData("sum(B5:B1)", 150)]           // Reversed corners are normalised.
    [InlineData("SUM(B1:C5)", 158.5)]         // Two columns; text and empties skipped.
    [InlineData("SUM(B1:B5, 10)", 160)]
    [InlineData("SUM(B1, B2)", 30)]
    [InlineData("SUM(B1, B2*2, 1)", 51)]
    [InlineData("SUM(C1:C5)", 8.5)]
    [InlineData("SUM(C3)", 0)]                // A lone text cell in an aggregate is ignored, as in Excel.
    [InlineData("AVERAGE(B1:B5)", 30)]
    [InlineData("AVERAGE(C1:C5)", 8.5 / 3)]
    [InlineData("COUNT(C1:C5)", 3)]
    [InlineData("COUNT(C1:C5, 7)", 4)]
    [InlineData("COUNTA(C1:C5)", 4)]
    [InlineData("MIN(B1:B5)", 10)]
    [InlineData("MAX(B1:B5)", 50)]
    [InlineData("MIN(C4)", 0)]                // MIN/MAX over nothing numeric is 0 in Excel.
    [InlineData("MAX(B1:B5, 99)", 99)]
    [InlineData("ROUND(B1/3, 2)", 3.33)]
    [InlineData("ROUND(2.5, 0)", 3)]          // Half away from zero, not banker's rounding.
    [InlineData("ROUND(-2.5, 0)", -3)]
    [InlineData("ROUND(1234.5, -2)", 1200)]
    [InlineData("ROUND(B2, 1)", 20)]
    [InlineData("ABS(-B1)", 10)]
    [InlineData("ABS(B1-B2)", 10)]
    [InlineData("ABS(C4)", 0)]
    [InlineData("'Data'!B1 + Data!B2", 30)]
    [InlineData("SUM('Data'!B1:B5)", 150)]
    [InlineData("SUM('DATA'!B1:B5)", 150)]    // Sheet names are case-insensitive.
    [InlineData("SUM(B1:B5)*2+ABS(-1)", 301)]
    [InlineData("ROUND(SUM(B1:B5)/COUNT(B1:B5), 0)", 30)]
    public void Evaluator_ComputesSupportedFormulas(string formula, double expected)
    {
        var actual = Eval(formula);
        Assert.NotNull(actual);
        Assert.Equal(expected, actual.Value, 10);
    }

    [Theory]
    [InlineData("SUM(B1:B10)")]               // Reaches the row being written.
    [InlineData("B10")]                       // The cell itself.
    [InlineData("B11")]                       // A row below.
    [InlineData("SUM(B1:B5) + C3")]           // Text in arithmetic is #VALUE!.
    [InlineData("C3*2")]
    [InlineData("D1")]
    [InlineData("'Other'!B1")]                // Another sheet's values are not available.
    [InlineData("SUM(Other!B1:B5)")]
    [InlineData("B1:B5")]                     // A bare range is not a number.
    [InlineData("IF(B1>5, 1, 0)")]            // Unsupported function.
    [InlineData("SUMIF(B1:B5, \">1\")")]
    [InlineData("B1>5")]                      // Comparisons.
    [InlineData("B1=B2")]
    [InlineData("\"a\"&\"b\"")]               // Strings and concatenation.
    [InlineData("2^3")]                       // Power is not in the subset.
    [InlineData("Total")]                     // Defined names.
    [InlineData("SUM(Lines[Amount])")]        // Structured references.
    [InlineData("1/0")]                       // #DIV/0!
    [InlineData("B1/C4")]                     // Division by an empty cell is also #DIV/0!.
    [InlineData("AVERAGE(C4)")]               // AVERAGE of nothing is #DIV/0!.
    [InlineData("SUM()")]
    [InlineData("ROUND(1.5)")]                // Wrong arity.
    [InlineData("ROUND(1.5, 0, 1)")]
    [InlineData("ABS(1, 2)")]
    [InlineData("ABS(B1:B2)")]                // A scalar function with a range.
    [InlineData("(1+2")]                      // Malformed input.
    [InlineData("1+")]
    [InlineData("SUM(B1:B5")]
    [InlineData("1 2")]
    [InlineData("")]
    [InlineData("A1B+1")]
    [InlineData("1.2.3")]
    [InlineData("SUM(B1:B5,)")]
    public void Evaluator_RefusesWhatItCannotPromise(string formula)
        => Assert.Null(Eval(formula));

    [Fact]
    public void Evaluator_SeesCellsToTheLeftOnTheSameRow_ButNotToTheRight()
    {
        var store = SampleStore();
        store.SetNumber(10, 2, 3);  // B10 = 3, written before D10.
        store.SetNumber(10, 3, 4);  // C10 = 4.
        Assert.Equal(12, Eval("B10*C10", store, row: 10, column: 4));
        Assert.Equal(8, Eval("SUM(B10:C10)+1", store, row: 10, column: 4));
        Assert.Null(Eval("E10", store, row: 10, column: 4));            // Not written yet.
        Assert.Null(Eval("SUM(B10:E10)", store, row: 10, column: 4));
        Assert.Null(Eval("D10", store, row: 10, column: 4));            // The cell itself.
    }

    [Fact]
    public void Evaluator_FailsOnUnknownCellsAndOpaqueRows()
    {
        var store = SampleStore();
        store.SetUnknown(3, 2);   // B3 holds a formula whose result was never known.
        Assert.Null(Eval("SUM(B1:B5)", store));
        Assert.Null(Eval("B3+1", store));
        Assert.Equal(30, Eval("SUM(B1:B2)", store));

        var template = SampleStore();
        template.MarkRowUnknown(4);   // A row copied from a template.
        Assert.Null(Eval("SUM(B1:B5)", template));
        Assert.Null(Eval("B4", template));
        Assert.Null(Eval("COUNTA(B4:B4)", template));
        Assert.Equal(60, Eval("SUM(B1:B3)", template));
    }

    [Fact]
    public void ValueStore_DistinguishesZeroEmptyTextAndUnknown_AndGrows()
    {
        var store = new FormulaValueStore();
        store.SetNumber(1, 1, 0);
        store.SetText(2, 1);
        store.SetUnknown(3, 1);
        store.SetNumber(100_000, 5_000, -1.25);   // Far away: forces both arrays to grow.
        Assert.Equal(FormulaCellState.Number, store.Get(1, 1, out var zero));
        Assert.Equal(0, zero);
        Assert.Equal(FormulaCellState.Text, store.Get(2, 1, out _));
        Assert.Equal(FormulaCellState.Unknown, store.Get(3, 1, out _));
        Assert.Equal(FormulaCellState.Empty, store.Get(4, 1, out _));
        Assert.Equal(FormulaCellState.Empty, store.Get(1, 2, out _));
        Assert.Equal(FormulaCellState.Empty, store.Get(99_999, 5_000, out _));
        Assert.Equal(FormulaCellState.Number, store.Get(100_000, 5_000, out var far));
        Assert.Equal(-1.25, far);
        Assert.Equal(FormulaCellState.Empty, store.Get(1, 16_384, out _));
    }

    // ----- End to end ------------------------------------------------------------------------------------------

    private static SpreadsheetDocument OpenValidated(byte[] bytes)
    {
        var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc).ToList();
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Path?.PartUri} {e.Path?.XPath}: {e.Description}")));
        return doc;
    }

    private static string SheetXml(byte[] bytes, int sheetNumber = 1)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry($"xl/worksheets/sheet{sheetNumber}.xml").Open());
        return reader.ReadToEnd();
    }

    private static Dictionary<int, ExcelRow> ReadRows(byte[] bytes, string sheet)
    {
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        return reader.ReadRows(sheet).ToDictionary(r => r.Index);
    }

    /// <summary>100 amount rows under a header, then a block of aggregate formulas; <paramref name="cache"/> toggles the feature.</summary>
    private static byte[] BuildReport(bool cache)
    {
        return ExcelPackage.CreateBytes(wb =>
        {
            var money = wb.Styles.Create(o => o.NumberFormat = "#,##0.00");
            var s = wb.Worksheets.Add("Q1 Report");
            if (cache) s.CacheFormulaResults();
            s.Row().WriteCell("#").WriteCell("Item").WriteCell("Amount");
            for (var i = 1; i <= 100; i++) s.Row().WriteCell(i).WriteCell("Item " + i).WriteCell(i * 1.25m, money);
            s.Row(102).WriteCell("Total").Skip().WriteFormula("SUM(C2:C101)", money);
            s.Row(103).WriteCell("Average").Skip().WriteFormula("=AVERAGE(C2:C101)");
            s.Row(104).WriteCell("Count").Skip().WriteFormula("COUNT(C2:C101)");
            s.Row(105).WriteCell("Half, rounded").Skip().WriteFormula("ROUND(C102/2, 1)");
            s.Row(106).WriteCell("Items").Skip().WriteFormula("COUNTA(B2:B101)");
            s.Row(107).WriteCell("Qualified").Skip().WriteFormula("SUM('Q1 Report'!C2:C101)");
            s.Row(108).WriteCell("Forward").Skip().WriteFormula("SUM(C2:C200)");          // Reaches past the current row.
            s.Row(109).WriteCell("Unsupported").Skip().WriteFormula("IF(C102>0,\"ok\",\"no\")");
            s.Row(110).WriteCell("Elsewhere").Skip().WriteFormula("SUM(Other!C2:C101)");
        });
    }

    [Fact]
    public void Streaming_WritesCachedResults_ReadableByExcelReader()
    {
        var bytes = BuildReport(cache: true);
        using (var doc = OpenValidated(bytes))
            Assert.True(doc.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad.Value); // Excel still recalculates.

        var rows = ReadRows(bytes, "Q1 Report");
        Assert.Equal(6312.5, (double)rows[102][2]!, 10);          // 1.25 * (1 + ... + 100)
        Assert.Equal(63.125, (double)rows[103][2]!, 10);
        Assert.Equal(100.0, rows[104][2]);
        Assert.Equal(3156.3, (double)rows[105][2]!, 10);          // 3156.25 rounded half away from zero.
        Assert.Equal(100.0, rows[106][2]);
        Assert.Equal(6312.5, (double)rows[107][2]!, 10);
        Assert.Null(rows[108][2]);                                // Forward reference: no cached value.
        Assert.Null(rows[109][2]);                                // Unsupported function: no cached value.
        Assert.Null(rows[110][2]);                                // Other sheet: no cached value.

        var xml = SheetXml(bytes);
        Assert.Contains("<f>SUM(C2:C101)</f><v>6312.5</v>", xml);
        Assert.Contains("<f>SUM(C2:C200)</f></c>", xml);
        Assert.Contains("<f>IF(C102&gt;0,\"ok\",\"no\")</f></c>", xml);
    }

    [Fact]
    public void Disabled_WritesNoCachedValues()
    {
        var bytes = BuildReport(cache: false);
        var xml = SheetXml(bytes);
        Assert.Contains("<f>SUM(C2:C101)</f></c>", xml);
        Assert.Contains("<f>AVERAGE(C2:C101)</f></c>", xml);
        Assert.DoesNotContain("</f><v>", xml);

        var rows = ReadRows(bytes, "Q1 Report");
        Assert.Null(rows[102][2]);
        Assert.Null(rows[103][2]);
        Assert.Equal(1.25, rows[2][2]); // Plain values are of course still there.
    }

    [Fact]
    public void FormulaResults_FeedLaterFormulas_AndSameRowReferencesWork()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Lines");
            s.CacheFormulaResults();
            s.Row().WriteCell("Qty").WriteCell("Price").WriteCell("Line total").WriteCell("Check");
            s.Row().WriteCell(2).WriteCell(9.5).WriteFormula("A2*B2").WriteFormula("C2/A2");       // 19, 9.5
            s.Row().WriteCell(3).WriteCell(4).WriteFormula("A3*B3").WriteFormula("E3");           // 12; E3 is to the right: none
            s.Row().WriteCell("Subtotal").Skip().WriteFormula("SUM(C2:C3)");                      // 31
            s.Row().WriteCell("Tax").Skip().WriteFormula("ROUND(C4*21%, 2)");                     // 6.51
            s.Row().WriteCell("Total").Skip().WriteFormula("C4+C5");                              // 37.51
            s.Row().WriteCell("Grand").Skip().WriteFormula("SUM(C2:C6)");                         // Includes the subtotals: 106.02
            s.Row().WriteCell(new DateTime(2026, 1, 31)).WriteCell(new DateTime(2026, 1, 1)).WriteFormula("A8-B8"); // 30 days
            s.Row().WriteCell(TimeSpan.FromHours(36)).Skip().WriteFormula("A9*24");               // 36 hours
            s.Row().WriteCell(true).WriteCell("text").WriteFormula("COUNT(A10:B10)").WriteFormula("COUNTA(A10:B10)"); // 0, 2
        });

        using (var doc = OpenValidated(bytes)) { }
        var rows = ReadRows(bytes, "Lines");
        Assert.Equal(19.0, rows[2][2]);
        Assert.Equal(9.5, rows[2][3]);
        Assert.Equal(12.0, rows[3][2]);
        Assert.Null(rows[3][3]);
        Assert.Equal(31.0, rows[4][2]);
        Assert.Equal(6.51, (double)rows[5][2]!, 10);
        Assert.Equal(37.51, (double)rows[6][2]!, 10);
        Assert.Equal(106.02, (double)rows[7][2]!, 10);
        Assert.Equal(30.0, rows[8][2]);
        Assert.Equal(36.0, rows[9][2]);
        Assert.Equal(0.0, rows[10][2]);
        Assert.Equal(2.0, rows[10][3]);
    }

    [Fact]
    public void ExplicitCachedValues_AreKeptAsGiven_AndVisibleToLaterFormulas()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.CacheFormulaResults();
            s.Row().WriteFormula("SOMETHING()", 5.0);      // The caller knows better; the evaluator is not consulted.
            s.Row().WriteFormula("\"a\"&\"b\"", "ab");
            s.Row().WriteFormula("A1*2");                  // Sees the explicit 5.
            s.Row().WriteFormula("COUNTA(A1:A2)");         // The text result counts as content.
        });
        var rows = ReadRows(bytes, "S");
        Assert.Equal(5.0, rows[1][0]);
        Assert.Equal("ab", rows[2][0]);
        Assert.Equal(10.0, rows[3][0]);
        Assert.Equal(2.0, rows[4][0]);
    }

    [Fact]
    public void CellsMode_EvaluatesAtFlush_InAnyWriteOrder()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Grid");
            s.Cells["A4"].Formula = "SUM(A1:A3)";          // Formulas may be set before their inputs exist.
            s.Cells["B4"].Formula = "A4*2";
            s.Cells["A5"].Formula = "SUM(A1:A4)";          // Includes the subtotal: 24.
            s.Cells["A6"].Formula = "SUM(A1:A7)";          // Forward reference: no value.
            s.Cells["A2"].Value = 7;
            s.Cells["A1"].Value = 5;
            s.Cells["A3"].Value = "n/a";
            s.CacheFormulaResults();                       // Nothing is written until save, so any time before it is fine.
        });

        using (var doc = OpenValidated(bytes)) { }
        var rows = ReadRows(bytes, "Grid");
        Assert.Equal(12.0, rows[4][0]);
        Assert.Equal(24.0, rows[4][1]);
        Assert.Equal(24.0, rows[5][0]);
        Assert.Null(rows[6][0]);
    }

    [Fact]
    public void EditMode_SeesLoadedValues_IncludingCachedFormulaResults()
    {
        var original = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            s.Row().WriteCell("Sku").WriteCell("Amount").WriteCell("Flag");
            s.Row().WriteCell("A").WriteCell(10.5).WriteCell(true);
            s.Row().WriteCell("B").WriteCell(20.25).WriteCell("x");
            s.Row().WriteCell("Total").WriteFormula("SUM(B2:B3)", 30.75);
            s.Row().WriteCell("Open").WriteFormula("SUM(B2:B3)");     // No cached value in the original.
        });

        using var output = new MemoryStream();
        using (var package = ExcelPackage.Edit(new MemoryStream(original), output, leaveOpen: true))
        {
            var s = package.Workbook.Worksheets.Edit("Data");
            s.CacheFormulaResults();
            s.Cells["B6"].Formula = "SUM(B2:B4)";       // Loaded numbers plus the loaded cached result: 61.5.
            s.Cells["C6"].Formula = "COUNTA(C2:C3)";    // Loaded boolean and text: 2.
            s.Cells["B7"].Formula = "SUM(B2:B5)";       // B5 is a formula without a known result: no value.
            package.Save();
        }
        var bytes = output.ToArray();
        using (var doc = OpenValidated(bytes)) { }
        var rows = ReadRows(bytes, "Data");
        Assert.Equal(61.5, rows[6][1]);
        Assert.Equal(2.0, rows[6][2]);
        Assert.Null(rows[7][1]);
        Assert.Equal(30.75, rows[4][1]); // The original cached value is untouched.
    }

    [Fact]
    public void TemplateMode_TreatsTemplateRowsAsOpaque()
    {
        var template = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Lines");
            s.Row(1).WriteCell("Qty");
            s.Row(2).WriteCell(5);   // A value the filler never sees.
        });

        using var output = new MemoryStream();
        using (var package = ExcelPackage.OpenTemplate(new MemoryStream(template), output, leaveOpen: true))
        {
            var s = package.Workbook.Worksheets.Open("Lines");
            s.CacheFormulaResults();
            s.Row(3).WriteCell(7);
            s.Row(4).WriteCell(8);
            s.Row(5).WriteFormula("SUM(A3:A4)");   // Only written rows: 15.
            s.Row(6).WriteFormula("SUM(A2:A4)");   // Touches template row 2: no value rather than a wrong 15.
            package.Save();
        }
        var bytes = output.ToArray();
        using (var doc = OpenValidated(bytes)) { }
        var rows = ReadRows(bytes, "Lines");
        Assert.Equal(5.0, rows[2][0]);
        Assert.Equal(15.0, rows[5][0]);
        Assert.Null(rows[6][0]);
    }

    [Fact]
    public void CacheFormulaResults_MustBeEnabledBeforeTheFirstRow_AndCanBeTurnedOff()
    {
        using var package = new ExcelPackage(new MemoryStream());
        var late = package.Workbook.Worksheets.Add("Late");
        late.Row().WriteCell(1);
        var ex = Assert.Throws<InvalidOperationException>(() => late.CacheFormulaResults());
        Assert.Contains("before its first row", ex.Message);
        late.CacheFormulaResults(false); // Turning off is always allowed.

        var early = package.Workbook.Worksheets.Add("Early");
        Assert.Same(early, early.CacheFormulaResults());
        Assert.Same(early, early.CacheFormulaResults()); // Idempotent.
        Assert.True(early.IsCachingFormulaResults);
        early.CacheFormulaResults(false);
        Assert.False(early.IsCachingFormulaResults);
    }

    [Fact]
    public void LoadFromCollection_ValuesAreSeen()
    {
        var items = Enumerable.Range(1, 10).Select(i => new Line { Name = "L" + i, Amount = i * 2.5 }).ToList();
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Loaded");
            s.CacheFormulaResults();
            s.LoadFromCollection(items);
            s.Row().WriteCell("Total").WriteFormula("SUM(B2:B11)");
        });
        var rows = ReadRows(bytes, "Loaded");
        Assert.Equal(137.5, rows[12][1]);
    }

    private sealed class Line
    {
        public string Name { get; set; } = "";
        public double Amount { get; set; }
    }
}
