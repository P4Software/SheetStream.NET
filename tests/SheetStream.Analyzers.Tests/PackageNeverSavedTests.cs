using SheetStream.Analyzers;

namespace SheetStream.Analyzers.Tests;

/// <summary>SS0004: a package that a method creates and never saves writes no workbook.</summary>
public class PackageNeverSavedTests
{
    private static Task Verify(string body, params Microsoft.CodeAnalysis.Testing.DiagnosticResult[] expected)
        => AnalyzerTest.Verify<PackageSaveAnalyzer>(body, expected);

    [Fact]
    public Task Using_declaration_without_Save_is_reported() => Verify(@"
        using var package = {|SS0004:ExcelPackage.Create(""report.xlsx"")|};
        package.Workbook.Worksheets.Add(""Data"").Row().WriteCell(1);
    ");

    [Fact]
    public Task Using_statement_without_Save_is_reported() => Verify(@"
        using (var package = {|SS0004:ExcelPackage.Create(""report.xlsx"")|})
        {
            package.Workbook.Worksheets.Add(""Data"").Row().WriteCell(1);
        }
    ");

    [Fact]
    public Task Every_way_of_creating_a_package_is_covered() => Verify(@"
        using var a = {|SS0004:new ExcelPackage(new System.IO.MemoryStream())|};
        using var b = {|SS0004:ExcelPackage.OpenTemplate(""t.xlsx"", ""out.xlsx"")|};
        using var c = {|SS0004:ExcelPackage.Edit(""book.xlsx"")|};
        using var d = {|SS0004:ExcelPackage.Edit(""in.xlsx"", ""out.xlsx"")|};
        ExcelPackage e = {|SS0004:ExcelPackage.Create(""plain-local.xlsx"")|};
        e.Dispose();
    ");

    [Fact]
    public Task Message_tells_the_user_what_to_do()
    {
        var expected = Microsoft.CodeAnalysis.Testing.DiagnosticResult.CompilerWarning("SS0004")
            .WithLocation(0)
            .WithMessage("Package 'package' is never saved, so this code writes no workbook. Call package.Save() once the workbook is complete; disposing a package does not save it.");
        return Verify(@"
        using var package = {|#0:ExcelPackage.Create(""report.xlsx"")|};
        package.Workbook.Worksheets.Add(""Data"");
    ", expected);
    }

    [Fact]
    public Task Saved_package_is_fine() => Verify(@"
        using var package = ExcelPackage.Create(""report.xlsx"");
        package.Workbook.Worksheets.Add(""Data"").Row().WriteCell(1);
        package.Save();
    ");

    [Fact]
    public Task Save_on_one_branch_is_enough_to_stay_silent() => Verify(@"
        using var package = ExcelPackage.Create(""report.xlsx"");
        package.Workbook.Worksheets.Add(""Data"").Row().WriteCell(1);
        if (flag) package.Save();
    ");

    [Fact]
    public Task Save_inside_a_lambda_or_local_function_counts() => Verify(@"
        using var package = ExcelPackage.Create(""report.xlsx"");
        Action finish = () => package.Save();
        finish();

        using var other = ExcelPackage.Create(""other.xlsx"");
        void Finish() { other.Save(); }
        Finish();
    ");

    [Fact]
    public Task Save_taken_as_a_delegate_counts() => Verify(@"
        using var offloaded = ExcelPackage.Create(""a.xlsx"");
        System.Threading.Tasks.Task.Run(offloaded.Save).Wait();

        using var deferred = ExcelPackage.Create(""b.xlsx"");
        Action finish = deferred.Save;
        finish();
    ");

    [Fact]
    public Task Another_method_taken_as_a_delegate_is_not_a_save() => Verify(@"
        using var package = {|SS0004:ExcelPackage.Create(""a.xlsx"")|};
        Action release = package.Dispose;
    ");

    [Fact]
    public Task Package_handed_to_other_code_is_not_reported() => Verify(@"
        using var passed = ExcelPackage.Create(""a.xlsx"");
        Hand(passed);

        var stored = ExcelPackage.Create(""b.xlsx"");
        _kept = stored;

        var listed = ExcelPackage.Create(""c.xlsx"");
        var all = new List<ExcelPackage> { listed };

        var copied = ExcelPackage.Create(""d.xlsx"");
        var alias = copied;
        alias.Save();

        var disposable = ExcelPackage.Create(""e.xlsx"");
        IDisposable asInterface = disposable;
    ");

    [Fact]
    public Task Returned_package_is_not_reported() => AnalyzerTest.Verify<PackageSaveAnalyzer>(@"
        ExcelPackage Open()
        {
            var package = ExcelPackage.Create(""report.xlsx"");
            package.Workbook.Worksheets.Add(""Data"");
            return package;
        }
        Open().Save();
    ");

    [Fact]
    public Task Package_not_created_here_is_not_reported() => Verify(@"
        ExcelPackage fromField = _kept;
        fromField.Workbook.Worksheets.Add(""Data"");

        ExcelPackage later;
        later = ExcelPackage.Create(""later.xlsx"");
        later.Workbook.Worksheets.Add(""Data"");

        var chosen = flag ? ExcelPackage.Create(""x.xlsx"") : ExcelPackage.Create(""y.xlsx"");
        chosen.Workbook.Worksheets.Add(""Data"");
    ");

    [Fact]
    public Task CreateBytes_and_reading_are_not_packages_to_save() => Verify(@"
        var bytes = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add(""Data"").Row().WriteCell(1));
        using var reader = SheetStream.Reading.ExcelReader.Open(new System.IO.MemoryStream(bytes));
    ");
}