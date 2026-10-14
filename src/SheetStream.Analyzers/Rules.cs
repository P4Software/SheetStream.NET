using Microsoft.CodeAnalysis;

namespace SheetStream.Analyzers
{
    /// <summary>
    /// The diagnostics SheetStream reports at design time. Each one mirrors an <c>InvalidOperationException</c>
    /// the library would otherwise throw at run time, so the message says what the fix is, not only what is wrong.
    /// </summary>
    public static class Rules
    {
        /// <summary>Category shown in the IDE's error list and used in <c>.editorconfig</c> severity rules.</summary>
        public const string Category = "SheetStream";

        private const string HelpPage = "https://p4software.github.io/SheetStream/analyzers.html";

        /// <summary>
        /// SS0001: a header-level setting (column widths, frozen panes, tab colour, gridlines, zoom, right-to-left,
        /// fit-to-page) is applied after the sheet already wrote rows, so the header has been flushed.
        /// </summary>
        public static readonly DiagnosticDescriptor LayoutAfterRows = new DiagnosticDescriptor(
            id: "SS0001",
            title: "Layout setting after rows were written",
            messageFormat: "{0} is set on worksheet '{1}' after rows were written, so SheetStream will throw at run time. Move this before the first Row() call, or call AutoFitColumns() first.",
            category: Category,
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true,
            description: "SheetStream streams each worksheet forward-only and the OpenXML schema places column widths, frozen panes, tab colour, gridlines, zoom, right-to-left and fit-to-page before the sheet data. Once the first row is written the header is already in the file. Set these before the first Row() or LoadFrom*() call, or call AutoFitColumns() so the header is deferred until the sheet is saved.",
            helpLinkUri: HelpPage + "#ss0001");

        /// <summary>SS0002: the same worksheet is written both through <c>Row()</c> and through <c>Cells[...]</c>.</summary>
        public static readonly DiagnosticDescriptor RowAndCellsMixed = new DiagnosticDescriptor(
            id: "SS0002",
            title: "Row() and Cells[] used on the same worksheet",
            messageFormat: "Worksheet '{0}' uses both Row() and Cells[], which SheetStream rejects at run time. Write a sheet either with Row() streaming or with Cells[] random access, and move the other data to a separate sheet.",
            category: Category,
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true,
            description: "A SheetStream worksheet is either streamed forward-only through Row() and LoadFrom*(), or buffered in memory through Cells[row, column]. Mixing the two on one sheet throws InvalidOperationException. Pick one mode per sheet.",
            helpLinkUri: HelpPage + "#ss0002");

        /// <summary>SS0003: a literal <c>Row(n)</c> is written after a literal <c>Row(m)</c> with <c>n &lt; m</c>.</summary>
        public static readonly DiagnosticDescriptor RowsMustAscend = new DiagnosticDescriptor(
            id: "SS0003",
            title: "Rows must ascend",
            messageFormat: "Row({0}) comes after Row({1}) on worksheet '{2}', but rows are streamed forward-only and this will throw at run time. Write the rows in ascending order, or use Cells[row, column] for random access.",
            category: Category,
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true,
            description: "SheetStream writes rows straight into the file, so a row index lower than the current row cannot be written any more and throws InvalidOperationException. Write rows in ascending order, or use Cells[row, column], which buffers the sheet and sorts the cells when it is saved.",
            helpLinkUri: HelpPage + "#ss0003");
    }
}
