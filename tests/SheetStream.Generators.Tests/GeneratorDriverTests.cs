using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace SheetStream.Generators.Tests;

/// <summary>Runs the generator in isolation through <see cref="CSharpGeneratorDriver"/> to check its diagnostics and output.</summary>
public class GeneratorDriverTests
{
    private static (GeneratorDriverRunResult Run, Compilation Output) Run(string source, LanguageVersion language = LanguageVersion.Latest)
    {
        // Every assembly the test process itself loaded, so System.Runtime, System.ComponentModel.Primitives
        // and SheetStream.dll are all there without hard-coding paths.
        var tpa = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!).Split(Path.PathSeparator);
        var references = tpa.Select(p => (MetadataReference)MetadataReference.CreateFromFile(p)).ToList();
        Assert.Contains(tpa, p => p.EndsWith("SheetStream.dll", StringComparison.OrdinalIgnoreCase));

        var compilation = CSharpCompilation.Create(
            "Consumer",
            new[] { CSharpSyntaxTree.ParseText(source, new CSharpParseOptions(language), path: "Models.cs") },
            references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, nullableContextOptions: NullableContextOptions.Enable));

        var driver = CSharpGeneratorDriver.Create(new[] { new SheetRowGenerator().AsSourceGenerator() }, parseOptions: new CSharpParseOptions(language));
        driver = (CSharpGeneratorDriver)driver.RunGeneratorsAndUpdateCompilation(compilation, out var output, out _);
        return (driver.GetRunResult(), output);
    }

    private static IEnumerable<string> Errors(Compilation c) => c.GetDiagnostics().Where(d => d.Severity == DiagnosticSeverity.Error).Select(d => d.ToString());

    [Fact]
    public void NonPartialType_ReportsSSG001_AndGeneratesNothing()
    {
        var (run, _) = Run("""
            using SheetStream;
            namespace N { [SheetRow] public class NotPartial { public int A { get; set; } } }
            """);
        var d = Assert.Single(run.Diagnostics);
        Assert.Equal("SSG001", d.Id);
        Assert.Equal(DiagnosticSeverity.Error, d.Severity);
        Assert.Contains("NotPartial", d.GetMessage());
        Assert.Equal(2, d.Location.GetLineSpan().StartLinePosition.Line + 1);
        Assert.Empty(run.GeneratedTrees);
    }

    [Fact]
    public void NonPartialContainer_ReportsSSG002()
    {
        var (run, _) = Run("""
            using SheetStream;
            public class Outer { [SheetRow] public partial class Inner { public int A { get; set; } } }
            """);
        var d = Assert.Single(run.Diagnostics);
        Assert.Equal("SSG002", d.Id);
        Assert.Contains("'Outer'", d.GetMessage());
        Assert.Empty(run.GeneratedTrees);
    }

    [Fact]
    public void NoMembers_ReportsSSG003Warning()
    {
        var (run, _) = Run("""
            using SheetStream;
            [SheetRow] public partial class Empty { [SheetColumn(Ignore = true)] public int A { get; set; } private int B { get; set; } public static int C { get; set; } }
            """);
        var d = Assert.Single(run.Diagnostics);
        Assert.Equal("SSG003", d.Id);
        Assert.Equal(DiagnosticSeverity.Warning, d.Severity);
        Assert.Empty(run.GeneratedTrees);
    }

    [Fact]
    public void GenericType_ReportsSSG004()
    {
        var (run, _) = Run("""
            using SheetStream;
            [SheetRow] public partial class Box<T> { public T? Value { get; set; } }
            """);
        Assert.Equal("SSG004", Assert.Single(run.Diagnostics).Id);
    }

    [Fact]
    public void ValidType_GeneratesCompilableMapperWithModuleInitializer()
    {
        var (run, output) = Run("""
            using System.ComponentModel;
            using SheetStream;
            namespace Shop.Models
            {
                public partial class Outer
                {
                    [SheetRow]
                    public partial class Item
                    {
                        [SheetColumn("Code", Order = 0)] public string Sku { get; set; } = "";
                        [DisplayName("Unit price")] [SheetColumn(NumberFormat = "0.00")] public decimal Price { get; set; }
                        public int? Stock;
                        public System.DateTime Added { get; init; }
                        public Kind Kind { get; set; }
                    }
                }
                public enum Kind { A, B }
            }
            """);
        Assert.Empty(run.Diagnostics);
        var tree = Assert.Single(run.GeneratedTrees);
        Assert.EndsWith("Shop.Models.Outer.Item.SheetMapper.g.cs", tree.FilePath);
        var text = tree.ToString();
        Assert.Contains("[global::System.Runtime.CompilerServices.ModuleInitializer]", text);
        // Column order: Sku (Order 0), then the unordered members with the field first (metadata token order,
        // matching reflection), then the properties in declaration order.
        Assert.Contains("row.WriteCell(item.Sku, styles[0]);", text);
        Assert.Contains("if (item.Stock.HasValue) row.WriteCell(item.Stock.GetValueOrDefault(), styles[1]); else row.WriteEmpty(styles[1]);", text);
        Assert.Contains("row.WriteCell(item.Price, styles[2]);", text);
        Assert.Contains("row.WriteCell(item.Kind.ToString(), styles[4]);", text);
        Assert.Contains("case 3: return; // read-only member", text);
        Assert.Contains("\"Code\",", text);
        Assert.Contains("\"Unit price\",", text);
        Assert.Contains("\"0.00\",", text);
        Assert.Empty(Errors(output));
        // No warnings either: the generated file must stay clean under TreatWarningsAsErrors.
        Assert.Empty(output.GetDiagnostics().Where(d => d.Severity == DiagnosticSeverity.Warning && d.Location.SourceTree == tree));
    }

    [Fact]
    public void OlderLanguageVersion_SkipsModuleInitializer()
    {
        var (run, output) = Run("""
            using SheetStream;
            [SheetRow] public partial class Row { public int A { get; set; } }
            """, LanguageVersion.CSharp8);
        var text = Assert.Single(run.GeneratedTrees).ToString();
        Assert.DoesNotContain("ModuleInitializer]", text);
        Assert.Contains("public static void RegisterSheetMapper()", text);
        Assert.Empty(Errors(output));
    }
}
