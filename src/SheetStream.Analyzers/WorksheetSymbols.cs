using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Operations;

namespace SheetStream.Analyzers
{
    /// <summary>What a member access on a worksheet means for the streaming rules.</summary>
    internal enum UsageKind
    {
        /// <summary>Starts or writes rows: <c>Row(...)</c> and the <c>LoadFrom*</c> family. Flushes the header.</summary>
        RowWrite,

        /// <summary>A setting the schema stores in the sheet header, so it must come before the first row.</summary>
        Layout,

        /// <summary><c>AutoFitColumns()</c>: defers the header until save, which lifts the ordering rule.</summary>
        AutoFit,

        /// <summary><c>Cells[...]</c>: puts the sheet in random-access mode (header deferred, <c>Row()</c> forbidden).</summary>
        Cells,
    }

    /// <summary>
    /// Resolves SheetStream's types once per compilation and classifies member symbols. Classification goes
    /// through the symbol's containing type, never through the name alone, so a user's own <c>Row()</c> method on an
    /// unrelated class is never mistaken for a worksheet write.
    /// </summary>
    internal sealed class WorksheetSymbols
    {
        private readonly INamedTypeSymbol _worksheet;
        private readonly INamedTypeSymbol? _pageSetup;

        private WorksheetSymbols(INamedTypeSymbol worksheet, INamedTypeSymbol? pageSetup)
        {
            _worksheet = worksheet;
            _pageSetup = pageSetup;
        }

        /// <summary>Returns <c>null</c> when the compilation does not reference SheetStream; the analyzer then stays idle.</summary>
        public static WorksheetSymbols? TryCreate(Compilation compilation)
        {
            var worksheet = compilation.GetTypeByMetadataName("SheetStream.Worksheet");
            if (worksheet == null) return null;
            return new WorksheetSymbols(worksheet, compilation.GetTypeByMetadataName("SheetStream.PageSetup"));
        }

        /// <summary>Classifies a method called on <c>SheetStream.Worksheet</c>, or returns <c>null</c> for anything else.</summary>
        public UsageKind? ClassifyMethod(IMethodSymbol method)
        {
            if (!SymbolEqualityComparer.Default.Equals(method.ContainingType, _worksheet)) return null;
            switch (method.Name)
            {
                case "Row":
                case "LoadFromCollection":
                case "LoadFromDataTable":
                case "LoadFromDataReader":
                case "LoadFromAsyncEnumerable":
                    return UsageKind.RowWrite;
                case "Column":
                case "Columns":
                case "FreezePanes":
                    return UsageKind.Layout;
                case "AutoFitColumns":
                    return UsageKind.AutoFit;
                default:
                    return null;
            }
        }

        /// <summary>True for the <c>Worksheet.Cells</c> property (reading it already switches the sheet's mode).</summary>
        public bool IsCells(IPropertySymbol property)
            => property.Name == "Cells" && SymbolEqualityComparer.Default.Equals(property.ContainingType, _worksheet);

        /// <summary>True for a <c>Worksheet</c> property whose setter writes into the sheet header.</summary>
        public bool IsLayoutProperty(IPropertySymbol property)
        {
            if (!SymbolEqualityComparer.Default.Equals(property.ContainingType, _worksheet)) return false;
            switch (property.Name)
            {
                case "TabColor":
                case "ShowGridLines":
                case "Zoom":
                case "RightToLeft":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>True for <c>PageSetup.FitToWidth</c> / <c>FitToHeight</c>, whose "fit to page" flag lives in <c>sheetPr</c>.</summary>
        public bool IsPageSetupLayoutProperty(IPropertySymbol property)
        {
            if (_pageSetup == null || !SymbolEqualityComparer.Default.Equals(property.ContainingType, _pageSetup)) return false;
            return property.Name == "FitToWidth" || property.Name == "FitToHeight";
        }

        /// <summary>True for the <c>Worksheet.PageSetup</c> property, the hop between a sheet and its print settings.</summary>
        public bool IsPageSetupAccessor(IPropertySymbol property)
            => property.Name == "PageSetup" && SymbolEqualityComparer.Default.Equals(property.ContainingType, _worksheet);

        /// <summary>
        /// Resolves the receiver of a member access to the local or parameter that holds the worksheet. Fields,
        /// properties and chained calls return <c>null</c>: they can be re-pointed by code outside the method
        /// (another method swapping a <c>_sheet</c> field, for instance), and a false warning costs more trust than
        /// a missed one.
        /// </summary>
        public static ISymbol? ResolveReceiver(IOperation? instance)
        {
            while (instance is IConversionOperation conversion) instance = conversion.Operand;
            switch (instance)
            {
                case ILocalReferenceOperation local:
                    return local.Local;
                case IParameterReferenceOperation parameter:
                    return parameter.Parameter;
                default:
                    return null;
            }
        }
    }
}
