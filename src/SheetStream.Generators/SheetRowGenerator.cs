using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;

namespace SheetStream.Generators
{
    /// <summary>
    /// Generates an <c>ISheetRowMapper&lt;T&gt;</c> for every partial type marked with <c>[SheetStream.SheetRow]</c>,
    /// so <c>LoadFromCollection</c> and <c>ExcelReader.Read&lt;T&gt;</c> work without reflection (trimming and
    /// Native AOT safe). The mapper is exposed as a static <c>SheetMapper</c> property and registered through a
    /// module initializer when the target framework supports one.
    /// </summary>
    [Generator(LanguageNames.CSharp)]
    public sealed class SheetRowGenerator : IIncrementalGenerator
    {
        private const string SheetRowAttributeName = "SheetStream.SheetRowAttribute";

        /// <inheritdoc />
        public void Initialize(IncrementalGeneratorInitializationContext context)
        {
            var models = context.SyntaxProvider.ForAttributeWithMetadataName(
                SheetRowAttributeName,
                predicate: static (node, _) => node is TypeDeclarationSyntax,
                transform: static (ctx, _) => SheetRowModelBuilder.Build(ctx));

            // Module initializers need the attribute type (net5.0+) and C# 9. Computed once per compilation
            // and combined into each model so the output stays cacheable.
            var canUseModuleInitializer = context.CompilationProvider.Select(static (compilation, _) =>
                compilation.GetTypeByMetadataName("System.Runtime.CompilerServices.ModuleInitializerAttribute") != null
                && compilation is CSharpCompilation csharp
                && csharp.LanguageVersion >= LanguageVersion.CSharp9);

            context.RegisterSourceOutput(models.Combine(canUseModuleInitializer), static (spc, pair) =>
            {
                var (model, moduleInitializer) = pair;
                foreach (var diagnostic in model.Diagnostics) spc.ReportDiagnostic(diagnostic.ToDiagnostic());
                if (model.HasErrors || model.Columns.Count == 0) return;
                var source = SheetRowEmitter.Emit(model, moduleInitializer && model.IsAccessibleAtModuleLevel);
                spc.AddSource(model.HintName + ".SheetMapper.g.cs", SourceText.From(source, System.Text.Encoding.UTF8));
            });
        }
    }
}
