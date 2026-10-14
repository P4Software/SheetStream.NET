namespace SheetStream.Analyzers.Tests;

/// <summary>SS0002: streaming and random access on the same sheet.</summary>
public class RowAndCellsMixedTests
{
    [Fact]
    public Task Cells_after_Row_is_reported_once() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row().WriteCell(1);
        {|SS0002:sheet.Cells[1, 1]|}.Value = 42;
        sheet.Cells[""B2""].Value = 43;
    ");

    [Fact]
    public Task Row_after_Cells_is_reported() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Cells[1, 1].Value = 42;
        {|SS0002:sheet.Row()|}.WriteCell(1);
    ");

    [Fact]
    public Task LoadFromCollection_counts_as_a_row_write() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.LoadFromCollection(orders);
        {|SS0002:sheet.Cells[1, 1]|}.Value = ""Total"";
    ");

    [Fact]
    public Task Message_tells_the_user_what_to_do()
    {
        var expected = Microsoft.CodeAnalysis.Testing.DiagnosticResult.CompilerWarning("SS0002")
            .WithLocation(0)
            .WithMessage("Worksheet 'sheet' uses both Row() and Cells[], which SheetStream rejects at run time. Write a sheet either with Row() streaming or with Cells[] random access, and move the other data to a separate sheet.");
        return AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row().WriteCell(1);
        {|#0:sheet.Cells[1, 1]|}.Value = 42;
    ", expected);
    }

    [Fact]
    public Task Different_sheets_are_fine() => AnalyzerTest.Verify(@"
        var streamed = workbook.Worksheets.Add(""Streamed"");
        streamed.Row().WriteCell(1);
        var random = workbook.Worksheets.Add(""Random"");
        random.Cells[1, 1].Value = 42;
    ");

    [Fact]
    public Task One_mode_per_branch_is_fine() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        if (flag) sheet.Row().WriteCell(1);
        else sheet.Cells[1, 1].Value = 42;
    ");

    [Fact]
    public Task Cells_only_is_fine() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Cells[3, 2].Value = 42;
        sheet.Cells[1, 1].Value = ""Id"";
        sheet.FreezePanes(1);
    ");
}
