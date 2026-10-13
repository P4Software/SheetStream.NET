using System;
using System.Collections.Generic;
using SheetStream.Internal;
using SheetStream.Styles;

namespace SheetStream
{
    /// <summary>
    /// The logical workbook: a collection of worksheets plus the style sheet, shared string table, defined
    /// names and document properties they share. There is exactly one workbook per <see cref="ExcelPackage"/>.
    /// </summary>
    public sealed class Workbook
    {
        private readonly HashSet<string> _tableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        internal Workbook(ExcelPackage package)
        {
            Package = package;
            Styles = new StyleSheet();
            SharedStrings = new SharedStringTable();
            Worksheets = new WorksheetCollection(this);
            DefinedNames = new DefinedNameCollection();
            Properties = new WorkbookProperties();
        }

        /// <summary>The package this workbook is written into.</summary>
        public ExcelPackage Package { get; }

        /// <summary>The worksheets, in the order they appear in Excel's tab strip.</summary>
        public WorksheetCollection Worksheets { get; }

        /// <summary>Registry of cell styles. Create styles here and pass them to the cell-writing methods.</summary>
        public StyleSheet Styles { get; }

        /// <summary>The shared string table used by every sheet. Exposed for diagnostics (counts).</summary>
        public SharedStringTable SharedStrings { get; }

        /// <summary>Named ranges and named formulas.</summary>
        public DefinedNameCollection DefinedNames { get; }

        /// <summary>Document properties (title, author, company...).</summary>
        public WorkbookProperties Properties { get; }

        /// <summary>True once any sheet has written a formula; the package then asks Excel to recalculate on open.</summary>
        internal bool HasFormulas { get; set; }

        /// <summary>Workbook protection verifier (hash, salt) and flags, or null when unprotected.</summary>
        internal (string? Hash, string? Salt, bool Structure, bool Windows)? Protection { get; private set; }

        // Workbook-wide part counters: table ids, drawing / chart / comment / image part numbers.
        internal int TableCount { get; set; }
        internal int DrawingCount { get; set; }
        internal int ChartCount { get; set; }
        internal int CommentsCount { get; set; }
        internal int ImageCount { get; set; }

        /// <summary>
        /// Protects the workbook structure so sheets cannot be added, deleted, renamed, moved, hidden or
        /// unhidden without the password. Cell editing is controlled per sheet with <see cref="Worksheet.Protect"/>.
        /// </summary>
        public void Protect(string? password = null, bool lockStructure = true, bool lockWindows = false)
        {
            string? hash = null, salt = null;
            if (!string.IsNullOrEmpty(password)) (hash, salt) = PasswordHasher.Hash(password!);
            Protection = (hash, salt, lockStructure, lockWindows);
        }

        /// <summary>Reserves a table name, enforcing Excel's workbook-wide uniqueness (case-insensitive).</summary>
        internal void ReserveTableName(string name)
        {
            ExcelTable.ValidateName(name);
            if (!_tableNames.Add(name)) throw new ArgumentException($"A table named '{name}' already exists in this workbook.", nameof(name));
            foreach (var dn in DefinedNames)
                if (string.Equals(dn.Name, name, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException($"'{name}' is already used as a defined name.", nameof(name));
        }
    }
}
