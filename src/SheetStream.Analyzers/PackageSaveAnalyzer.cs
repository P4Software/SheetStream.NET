using System.Collections.Generic;
using System.Collections.Immutable;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Operations;

namespace SheetStream.Analyzers
{
    /// <summary>
    /// Reports SS0004: a method creates an <c>ExcelPackage</c>, keeps it in a local and never calls
    /// <c>Save()</c> on it. Since disposing a package does not save, such code runs without error and writes
    /// nothing, which is exactly the kind of mistake that is cheap to catch while typing and expensive later.
    /// </summary>
    /// <remarks>
    /// The rule is deliberately narrow, because a false warning breaks the build of anyone who treats warnings
    /// as errors. It only looks at a local that is initialised directly from <c>new ExcelPackage(...)</c> or
    /// one of the static factories, and it stays silent as soon as the local is used for anything other than
    /// reaching its own members: passed as an argument, returned, assigned somewhere, captured in a tuple or
    /// compared. Any of those means other code may save the package, and that code is out of sight.
    /// </remarks>
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class PackageSaveAnalyzer : DiagnosticAnalyzer
    {
        /// <inheritdoc />
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics { get; } = ImmutableArray.Create(Rules.PackageNeverSaved);

        /// <inheritdoc />
        public override void Initialize(AnalysisContext context)
        {
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            context.EnableConcurrentExecution();
            context.RegisterCompilationStartAction(start =>
            {
                // Resolved once per compilation; a project that does not reference SheetStream pays nothing more.
                var package = start.Compilation.GetTypeByMetadataName("SheetStream.ExcelPackage");
                if (package == null) return;
                start.RegisterOperationBlockAction(block => AnalyzeBlock(block, package));
            });
        }

        private static void AnalyzeBlock(OperationBlockAnalysisContext context, INamedTypeSymbol package)
        {
            // Locals that hold a freshly created package, with the expression that created each.
            Dictionary<ILocalSymbol, IOperation>? created = null;
            foreach (var block in context.OperationBlocks)
            {
                foreach (var operation in block.Descendants())
                {
                    if (!(operation is IVariableDeclaratorOperation declarator)) continue;
                    var value = declarator.GetVariableInitializer()?.Value;
                    while (value is IConversionOperation conversion) value = conversion.Operand;
                    if (value == null || !CreatesPackage(value, package)) continue;
                    (created ??= new Dictionary<ILocalSymbol, IOperation>(SymbolEqualityComparer.Default))[declarator.Symbol] = value;
                }
            }
            if (created == null) return;

            // One pass over every use of those locals: a Save() call or anything that lets the package leave
            // the method clears the local.
            foreach (var block in context.OperationBlocks)
            {
                foreach (var operation in block.Descendants())
                {
                    if (!(operation is ILocalReferenceOperation reference) || !created.ContainsKey(reference.Local)) continue;
                    if (IsSaved(reference) || Escapes(reference)) created.Remove(reference.Local);
                    if (created.Count == 0) return;
                }
            }

            foreach (var pair in created)
                context.ReportDiagnostic(Diagnostic.Create(Rules.PackageNeverSaved, pair.Value.Syntax.GetLocation(), pair.Key.Name));
        }

        /// <summary>True for <c>new ExcelPackage(...)</c> and for the static factories that return a package.</summary>
        private static bool CreatesPackage(IOperation value, INamedTypeSymbol package)
        {
            if (value is IObjectCreationOperation creation)
                return SymbolEqualityComparer.Default.Equals(creation.Type, package);
            return value is IInvocationOperation call
                && call.TargetMethod.IsStatic
                && SymbolEqualityComparer.Default.Equals(call.TargetMethod.ContainingType, package)
                && SymbolEqualityComparer.Default.Equals(call.TargetMethod.ReturnType, package);
        }

        /// <summary>The operation that uses the reference, looking through conversions and parentheses.</summary>
        private static IOperation? User(ILocalReferenceOperation reference, out IOperation used)
        {
            used = reference;
            var parent = reference.Parent;
            while (parent is IConversionOperation || parent is IParenthesizedOperation)
            {
                used = parent;
                parent = parent.Parent;
            }
            return parent;
        }

        /// <summary>
        /// True for <c>package.Save()</c> and for <c>package.Save</c> taken as a delegate
        /// (<c>Task.Run(package.Save)</c>, <c>Action finish = package.Save</c>), which will be called by code
        /// this rule cannot follow.
        /// </summary>
        private static bool IsSaved(ILocalReferenceOperation reference)
        {
            switch (User(reference, out var used))
            {
                case IInvocationOperation call:
                    return ReferenceEquals(call.Instance, used) && call.TargetMethod.Name == "Save";
                case IMethodReferenceOperation method:
                    return ReferenceEquals(method.Instance, used) && method.Method.Name == "Save";
                default:
                    return false;
            }
        }

        /// <summary>
        /// True when the reference is used for anything but reaching a member of the package itself. Reading
        /// <c>package.Workbook</c> or calling <c>package.Dispose()</c> keeps the package in the method; every
        /// other use is treated as handing it to code that may save it.
        /// </summary>
        private static bool Escapes(ILocalReferenceOperation reference)
        {
            var parent = User(reference, out var used);
            switch (parent)
            {
                case IInvocationOperation call:
                    return !ReferenceEquals(call.Instance, used);
                case IMemberReferenceOperation member:
                    return !ReferenceEquals(member.Instance, used);
                default:
                    return true;
            }
        }
    }
}