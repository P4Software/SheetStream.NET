namespace SheetStream.Analyzers.Tests;

/// <summary>SS0003: literal row indexes going backwards.</summary>
public class RowsMustAscendTests
{
    [Fact]
    public Task Smaller_literal_after_larger_is_reported() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row(5).WriteCell(1);
        {|SS0003:sheet.Row(3)|}.WriteCell(2);
    ");

    [Fact]
    public Task Message_tells_the_user_what_to_do()
    {
        var expected = Microsoft.CodeAnalysis.Testing.DiagnosticResult.CompilerWarning("SS0003")
            .WithLocation(0)
            .WithMessage("Row(3) comes after Row(5) on worksheet 'sheet', but rows are streamed forward-only and this will throw at run time. Write the rows in ascending order, or use Cells[row, column] for random access.");
        return AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row(5).WriteCell(1);
        {|#0:sheet.Row(3)|}.WriteCell(2);
    ", expected);
    }

    [Fact]
    public Task Ascending_literals_are_fine() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row(1).WriteCell(1);
        sheet.Row(3).WriteCell(2);
        sheet.Row(3).WriteCell(3); // same row again returns the open row
        sheet.Row(10).WriteCell(4);
    ");

    [Fact]
    public Task Row_with_options_overload_is_checked() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row(4, new RowOptions { Height = 30 }).WriteCell(1);
        {|SS0003:sheet.Row(2, null, 20)|}.WriteCell(2);
    ");

    [Fact]
    public Task Non_literal_indexes_are_skipped() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        const int start = 5;
        sheet.Row(start).WriteCell(1);
        sheet.Row(n).WriteCell(2);
        sheet.Row(3).WriteCell(3);
    ");

    [Fact]
    public Task Different_sheets_are_fine() => AnalyzerTest.Verify(@"
        var first = workbook.Worksheets.Add(""First"");
        first.Row(5).WriteCell(1);
        var second = workbook.Worksheets.Add(""Second"");
        second.Row(3).WriteCell(2);
    ");

    [Fact]
    public Task Larger_literal_inside_a_branch_is_not_certain() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        if (flag) sheet.Row(5).WriteCell(1);
        sheet.Row(3).WriteCell(2);
    ");

    [Fact]
    public Task Smaller_literal_inside_a_branch_after_a_larger_one_is_reported() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row(5).WriteCell(1);
        if (flag) {|SS0003:sheet.Row(3)|}.WriteCell(2);
    ");
}
