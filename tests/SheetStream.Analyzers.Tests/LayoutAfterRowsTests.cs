namespace SheetStream.Analyzers.Tests;

/// <summary>SS0001: header-level settings applied after rows were written.</summary>
public class LayoutAfterRowsTests
{
    [Fact]
    public Task FreezePanes_after_Row_is_reported() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row().WriteCell(""Id"").WriteCell(""Name"");
        {|SS0001:sheet.FreezePanes(1)|};
    ");

    [Fact]
    public Task TabColor_after_LoadFromCollection_is_reported() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.LoadFromCollection(orders);
        {|SS0001:sheet.TabColor = ""#C00000""|};
    ");

    [Fact]
    public Task Column_width_after_LoadFromDataTable_is_reported() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Data"");
        sheet.LoadFromDataTable(table);
        {|SS0001:sheet.Column(1)|}.SetWidth(20);
        {|SS0001:sheet.Columns(2, 5)|}.SetWidth(12);
    ");

    [Fact]
    public Task PageSetup_FitToWidth_after_Row_is_reported() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row(1).WriteCell(1);
        {|SS0001:sheet.PageSetup.FitToWidth = 1|};
        sheet.PageSetup.Orientation = PageOrientation.Landscape; // trailer setting: allowed any time
    ");

    [Fact]
    public Task ShowGridLines_Zoom_RightToLeft_after_Row_are_reported() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row().WriteCell(1);
        {|SS0001:sheet.ShowGridLines = false|};
        {|SS0001:sheet.Zoom = 80|};
        {|SS0001:sheet.RightToLeft = true|};
    ");

    [Fact]
    public Task Message_tells_the_user_what_to_do()
    {
        // The whole point of the analyzer is the message, so pin it once.
        var expected = Microsoft.CodeAnalysis.Testing.DiagnosticResult.CompilerWarning("SS0001")
            .WithLocation(0)
            .WithMessage("FreezePanes() is set on worksheet 'sheet' after rows were written, so SheetStream will throw at run time. Move this before the first Row() call, or call AutoFitColumns() first.");
        return AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row().WriteCell(1);
        {|#0:sheet.FreezePanes(1)|};
    ", expected);
    }

    [Fact]
    public Task Layout_set_first_is_fine() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.TabColor = ""#C00000"";
        sheet.FreezePanes(1);
        sheet.Column(1).SetWidth(20);
        sheet.PageSetup.FitToWidth = 1;
        sheet.Row().WriteCell(""Id"");
        sheet.LoadFromCollection(orders);
    ");

    [Fact]
    public Task AutoFitColumns_lifts_the_rule() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.AutoFitColumns();
        sheet.LoadFromCollection(orders);
        sheet.FreezePanes(1);
        sheet.TabColor = ""#C00000"";
    ");

    [Fact]
    public Task Cells_lifts_the_rule() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Cells[1, 1].Value = ""Id"";
        sheet.Cells[2, 1].Value = 42;
        sheet.FreezePanes(1);
        sheet.Column(1).SetWidth(20);
    ");

    [Fact]
    public Task Cells_lifts_the_rule_even_when_Row_is_also_used() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row().WriteCell(1);
        {|SS0002:sheet.Cells[1, 1]|}.Value = 42;
        sheet.FreezePanes(1);
    ");

    [Fact]
    public Task Different_worksheet_variables_do_not_interfere() => AnalyzerTest.Verify(@"
        var first = workbook.Worksheets.Add(""First"");
        first.Row().WriteCell(1);
        var second = workbook.Worksheets.Add(""Second"");
        second.FreezePanes(1);
        second.TabColor = ""#C00000"";
        second.Row().WriteCell(2);
    ");

    [Fact]
    public Task Parameter_receiver_is_tracked() => AnalyzerTest.Verify(@"
        given.Row().WriteCell(1);
        {|SS0001:given.FreezePanes(1)|};
    ");

    [Fact]
    public Task Row_inside_a_branch_is_not_certain() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        if (flag) sheet.Row().WriteCell(1);
        sheet.FreezePanes(1);
    ");

    [Fact]
    public Task Row_inside_a_loop_is_not_certain() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        foreach (var order in orders) sheet.Row().WriteCell(order.Id);
        sheet.FreezePanes(1);
    ");

    [Fact]
    public Task Row_in_one_branch_and_layout_in_the_other_is_fine() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        if (flag) { sheet.Row().WriteCell(1); } else { sheet.FreezePanes(1); }
    ");

    [Fact]
    public Task Layout_inside_a_branch_after_an_unconditional_Row_is_reported() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row().WriteCell(1);
        if (flag) { {|SS0001:sheet.FreezePanes(1)|}; }
    ");

    [Fact]
    public Task Both_inside_the_same_loop_body_is_reported() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        foreach (var order in orders)
        {
            sheet.Row().WriteCell(order.Id);
            {|SS0001:sheet.Zoom = 90|};
        }
    ");

    [Fact]
    public Task Reassigned_variable_is_skipped() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""First"");
        sheet.Row().WriteCell(1);
        sheet = workbook.Worksheets.Add(""Second"");
        sheet.FreezePanes(1);
    ");

    [Fact]
    public Task Chained_receiver_is_skipped() => AnalyzerTest.Verify(@"
        workbook.Worksheets.Add(""First"").Row().WriteCell(1);
        workbook.Worksheets.Add(""First"").FreezePanes(1);
    ");

    [Fact]
    public Task Reading_a_layout_property_is_fine() => AnalyzerTest.Verify(@"
        var sheet = workbook.Worksheets.Add(""Orders"");
        sheet.Row().WriteCell(1);
        var zoom = sheet.Zoom;
        Console.WriteLine(zoom);
    ");
}
