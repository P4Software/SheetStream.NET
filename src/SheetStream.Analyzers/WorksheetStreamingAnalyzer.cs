using System.Collections.Generic;
using System.Collections.Immutable;
using System.Globalization;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Operations;

namespace SheetStream.Analyzers
{
    /// <summary>
    /// Warns at design time about code that breaks SheetStream's streaming rules and would throw at run time:
    /// SS0001 (layout after rows), SS0002 (Row() mixed with Cells[]) and SS0003 (rows not ascending). The three
    /// rules share one walk over each method body because they all need the same thing, the list of worksheet
    /// member accesses grouped by the variable that holds the sheet.
    /// </summary>
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class WorksheetStreamingAnalyzer : DiagnosticAnalyzer
    {
        /// <inheritdoc />
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics
            => ImmutableArray.Create(Rules.LayoutAfterRows, Rules.RowAndCellsMixed, Rules.RowsMustAscend);

        /// <inheritdoc />
        public override void Initialize(AnalysisContext context)
        {
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            context.EnableConcurrentExecution();
            context.RegisterCompilationStartAction(start =>
            {
                // Resolve the SheetStream types once; projects that do not reference the library get no work at all.
                var symbols = WorksheetSymbols.TryCreate(start.Compilation);
                if (symbols == null) return;
                start.RegisterOperationBlockAction(block => Analyze(block, symbols));
            });
        }

        /// <summary>One member access on a worksheet variable, with what the analyzer needs to know about it.</summary>
        private readonly struct Usage
        {
            public Usage(UsageKind kind, ISymbol receiver, IOperation operation, SyntaxNode reportAt, string member, int? rowLiteral)
            {
                Kind = kind;
                Receiver = receiver;
                Operation = operation;
                ReportAt = reportAt;
                Member = member;
                RowLiteral = rowLiteral;
            }

            public UsageKind Kind { get; }
            public ISymbol Receiver { get; }
            /// <summary>The operation used for ordering (the invocation, property reference or assignment).</summary>
            public IOperation Operation { get; }
            /// <summary>The syntax a diagnostic is reported on; may be wider than <see cref="Operation"/> for readability.</summary>
            public SyntaxNode ReportAt { get; }
            /// <summary>Display name for messages, e.g. "FreezePanes()", "TabColor" or "PageSetup.FitToWidth".</summary>
            public string Member { get; }
            /// <summary>The integer literal passed to <c>Row(n)</c>, or <c>null</c> for any other argument.</summary>
            public int? RowLiteral { get; }
        }

        private static void Analyze(OperationBlockAnalysisContext context, WorksheetSymbols symbols)
        {
            var usages = new List<Usage>();
            var reassigned = new HashSet<ISymbol>(SymbolEqualityComparer.Default);

            foreach (var block in context.OperationBlocks)
            {
                foreach (var operation in block.DescendantsAndSelf())
                {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    switch (operation)
                    {
                        case IInvocationOperation invocation:
                            CollectInvocation(invocation, symbols, usages);
                            break;
                        case IPropertyReferenceOperation property:
                            CollectPropertyReference(property, symbols, usages);
                            break;
                        case IAssignmentOperation assignment:
                            // A variable that is re-pointed inside the method may hold a different sheet later on,
                            // so nothing can be said about ordering on it. Track and skip it.
                            var target = WorksheetSymbols.ResolveReceiver(assignment.Target);
                            if (target != null) reassigned.Add(target);
                            break;
                        case IArgumentOperation argument when argument.Parameter != null && argument.Parameter.RefKind != RefKind.None:
                            var byRef = WorksheetSymbols.ResolveReceiver(argument.Value);
                            if (byRef != null) reassigned.Add(byRef);
                            break;
                        case IBranchOperation branch when branch.BranchKind == BranchKind.GoTo:
                            // goto can run statements in any order; lexical reasoning no longer holds.
                            return;
                    }
                }
            }
            if (usages.Count == 0) return;

            // Group by the variable holding the sheet; every rule is about one sheet at a time.
            var byReceiver = new Dictionary<ISymbol, List<Usage>>(SymbolEqualityComparer.Default);
            foreach (var usage in usages)
            {
                if (reassigned.Contains(usage.Receiver)) continue;
                if (!byReceiver.TryGetValue(usage.Receiver, out var list)) byReceiver[usage.Receiver] = list = new List<Usage>();
                list.Add(usage);
            }

            foreach (var pair in byReceiver)
            {
                ReportLayoutAfterRows(context, pair.Value);
                ReportRowAndCellsMixed(context, pair.Value);
                ReportRowsMustAscend(context, pair.Value);
            }
        }

        private static void CollectInvocation(IInvocationOperation invocation, WorksheetSymbols symbols, List<Usage> usages)
        {
            var kind = symbols.ClassifyMethod(invocation.TargetMethod);
            if (kind == null) return;
            var receiver = WorksheetSymbols.ResolveReceiver(invocation.Instance);
            if (receiver == null) return;

            int? rowLiteral = null;
            if (kind == UsageKind.RowWrite && invocation.TargetMethod.Name == "Row")
            {
                // Only an actual integer literal counts for SS0003; constants, variables and expressions are skipped
                // because the rule promises to be certain, not clever.
                foreach (var argument in invocation.Arguments)
                {
                    if (argument.Parameter?.Name == "index" && argument.Value is ILiteralOperation literal && literal.ConstantValue.HasValue && literal.ConstantValue.Value is int n)
                        rowLiteral = n;
                }
            }
            usages.Add(new Usage(kind.Value, receiver, invocation, invocation.Syntax, invocation.TargetMethod.Name + "()", rowLiteral));
        }

        private static void CollectPropertyReference(IPropertyReferenceOperation property, WorksheetSymbols symbols, List<Usage> usages)
        {
            if (symbols.IsCells(property.Property))
            {
                var receiver = WorksheetSymbols.ResolveReceiver(property.Instance);
                if (receiver == null) return;
                // Report on "sheet.Cells[1, 1]" rather than the bare "sheet.Cells" when the indexer follows.
                var reportAt = property.Parent is IPropertyReferenceOperation indexer && indexer.Property.IsIndexer && indexer.Instance == property
                    ? indexer.Syntax
                    : property.Syntax;
                usages.Add(new Usage(UsageKind.Cells, receiver, property, reportAt, "Cells[]", null));
                return;
            }

            // Only the setter matters: reading sheet.Zoom after rows is harmless.
            if (!(property.Parent is IAssignmentOperation assignment) || assignment.Target != property) return;

            if (symbols.IsLayoutProperty(property.Property))
            {
                var receiver = WorksheetSymbols.ResolveReceiver(property.Instance);
                if (receiver == null) return;
                usages.Add(new Usage(UsageKind.Layout, receiver, assignment, assignment.Syntax, property.Property.Name, null));
            }
            else if (symbols.IsPageSetupLayoutProperty(property.Property)
                && property.Instance is IPropertyReferenceOperation pageSetup
                && symbols.IsPageSetupAccessor(pageSetup.Property))
            {
                var receiver = WorksheetSymbols.ResolveReceiver(pageSetup.Instance);
                if (receiver == null) return;
                usages.Add(new Usage(UsageKind.Layout, receiver, assignment, assignment.Syntax, "PageSetup." + property.Property.Name, null));
            }
        }

        /// <summary>SS0001: a layout setting that certainly runs after a row write, unless the header is deferred.</summary>
        private static void ReportLayoutAfterRows(OperationBlockAnalysisContext context, List<Usage> usages)
        {
            foreach (var usage in usages)
            {
                // AutoFitColumns() and Cells[] both defer the header to save time, anywhere in the method.
                if (usage.Kind == UsageKind.AutoFit || usage.Kind == UsageKind.Cells) return;
            }

            foreach (var layout in usages)
            {
                if (layout.Kind != UsageKind.Layout) continue;
                foreach (var write in usages)
                {
                    if (write.Kind != UsageKind.RowWrite || !OperationOrder.RunsBefore(write.Operation, layout.Operation)) continue;
                    context.ReportDiagnostic(Diagnostic.Create(Rules.LayoutAfterRows, layout.ReportAt.GetLocation(), layout.Member, layout.Receiver.Name));
                    break;
                }
            }
        }

        /// <summary>SS0002: a Row() write and a Cells[] access that certainly both run, reported once per sheet variable.</summary>
        private static void ReportRowAndCellsMixed(OperationBlockAnalysisContext context, List<Usage> usages)
        {
            Usage? report = null;
            foreach (var a in usages)
            {
                if (a.Kind != UsageKind.RowWrite && a.Kind != UsageKind.Cells) continue;
                foreach (var b in usages)
                {
                    if (b.Kind == a.Kind || (b.Kind != UsageKind.RowWrite && b.Kind != UsageKind.Cells)) continue;
                    if (!OperationOrder.RunsBefore(a.Operation, b.Operation)) continue;
                    // Point at the earliest place where the second mode is introduced.
                    if (report == null || b.Operation.Syntax.SpanStart < report.Value.Operation.Syntax.SpanStart) report = b;
                }
            }
            if (report != null)
                context.ReportDiagnostic(Diagnostic.Create(Rules.RowAndCellsMixed, report.Value.ReportAt.GetLocation(), report.Value.Receiver.Name));
        }

        /// <summary>SS0003: a literal Row(n) that certainly runs after a literal Row(m) with n &lt; m.</summary>
        private static void ReportRowsMustAscend(OperationBlockAnalysisContext context, List<Usage> usages)
        {
            foreach (var later in usages)
            {
                if (later.RowLiteral == null) continue;
                foreach (var earlier in usages)
                {
                    if (earlier.RowLiteral == null || earlier.RowLiteral.Value <= later.RowLiteral.Value) continue;
                    if (!OperationOrder.RunsBefore(earlier.Operation, later.Operation)) continue;
                    context.ReportDiagnostic(Diagnostic.Create(Rules.RowsMustAscend, later.ReportAt.GetLocation(),
                        later.RowLiteral.Value.ToString(CultureInfo.InvariantCulture),
                        earlier.RowLiteral.Value.ToString(CultureInfo.InvariantCulture),
                        later.Receiver.Name));
                    break;
                }
            }
        }
    }
}
