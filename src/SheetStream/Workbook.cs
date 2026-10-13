using SheetStream.Styles;

namespace SheetStream
{
    /// <summary>
    /// The logical workbook: a collection of worksheets plus the style sheet and shared string table they
    /// share. There is exactly one workbook per <see cref="ExcelPackage"/>.
    /// </summary>
    public sealed class Workbook
    {
        internal Workbook(ExcelPackage package)
        {
            Package = package;
            Styles = new StyleSheet();
            SharedStrings = new SharedStringTable();
            Worksheets = new WorksheetCollection(this);
        }

        /// <summary>The package this workbook is written into.</summary>
        public ExcelPackage Package { get; }

        /// <summary>The worksheets, in the order they appear in Excel's tab strip.</summary>
        public WorksheetCollection Worksheets { get; }

        /// <summary>Registry of cell styles. Create styles here and pass them to the cell-writing methods.</summary>
        public StyleSheet Styles { get; }

        /// <summary>The shared string table used by every sheet. Exposed for diagnostics (counts).</summary>
        public SharedStringTable SharedStrings { get; }

        /// <summary>True once any sheet has written a formula; the package then asks Excel to recalculate on open.</summary>
        internal bool HasFormulas { get; set; }
    }
}
