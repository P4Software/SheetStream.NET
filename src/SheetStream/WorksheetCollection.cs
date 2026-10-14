using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>
    /// The worksheets of a workbook. Adding a sheet seals the previous one, because the package streams
    /// one sheet at a time straight into the zip archive.
    /// </summary>
    public sealed class WorksheetCollection : IReadOnlyList<Worksheet>
    {
        /// <summary>Characters Excel forbids in a sheet name.</summary>
        private static readonly char[] ForbiddenNameChars = { ':', '\\', '/', '?', '*', '[', ']' };

        private readonly Workbook _workbook;
        private readonly List<Worksheet> _sheets = new List<Worksheet>();
        private readonly HashSet<string> _names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        internal WorksheetCollection(Workbook workbook)
        {
            _workbook = workbook;
        }

        /// <inheritdoc />
        public int Count => _sheets.Count;

        /// <inheritdoc />
        public Worksheet this[int index] => _sheets[index];

        /// <summary>The sheet currently accepting rows, or <c>null</c> before the first sheet is added.</summary>
        public Worksheet? Current => _sheets.Count == 0 ? null : _sheets[_sheets.Count - 1];

        /// <summary>
        /// Adds a worksheet and makes it the current one. The previously current sheet is sealed and can
        /// no longer receive rows.
        /// </summary>
        /// <param name="name">
        /// Tab name: 1 to 31 characters, unique within the workbook (case-insensitive), without any of
        /// <c>: \ / ? * [ ]</c>, and not starting or ending with an apostrophe.
        /// </param>
        public Worksheet Add(string name)
        {
            if (_workbook.Template != null)
                throw new NotSupportedException("A workbook filled from a template keeps the template's sheets. Use Worksheets.Open(name) to fill one, or add the sheet to the template in Excel.");
            ValidateName(name);
            if (_workbook.Package.IsSaved) throw new InvalidOperationException("Cannot add a worksheet after the package has been saved.");

            Current?.Seal();

            var index = _sheets.Count + 1;
            var writer = _workbook.Package.OpenPart("xl/worksheets/sheet" + index.ToString(CultureInfo.InvariantCulture) + ".xml");
            var sheet = new Worksheet(_workbook, index, name, writer);
            _sheets.Add(sheet);
            _names.Add(name);
            return sheet;
        }

        /// <summary>
        /// Opens a sheet of the template for filling (template mode only) and makes it the current one; the
        /// previously current sheet is sealed. Sheets you do not open are copied unchanged.
        /// </summary>
        /// <param name="name">The sheet's tab name in the template (case-insensitive).</param>
        public Worksheet Open(string name)
        {
            var template = _workbook.Template ?? throw new InvalidOperationException("Worksheets.Open is for workbooks created with ExcelPackage.OpenTemplate. Use Worksheets.Add to create sheets.");
            return OpenCore(template, FindTemplateSheet(template, name), edit: false);
        }

        /// <summary>
        /// Opens a sheet of the workbook being edited with every existing cell loaded into <see cref="Worksheet.Cells"/>
        /// (edit mode), and makes it the current one; the previously current sheet is sealed. Read values with
        /// <c>sheet.Cells["B2"].Value</c>, change them, add rows below <c>sheet.Cells.Dimension.LastRow</c>, and
        /// <see cref="ExcelPackage.Save"/> writes the whole sheet back. Cells you do not touch are written exactly
        /// as they were; sheets you do not open are copied unchanged. Meant for small workbooks: a sheet with
        /// more than <see cref="PackageOptions.MaxEditableCells"/> cells is refused.
        /// </summary>
        /// <param name="name">The sheet's tab name (case-insensitive).</param>
        /// <example>
        /// <code>
        /// using var package = ExcelPackage.Edit("budget.xlsx");
        /// var sheet = package.Workbook.Worksheets.Edit("Q4");
        /// sheet.Cells["B4"].Value = 1250;
        /// package.Save();
        /// </code>
        /// </example>
        public Worksheet Edit(string name)
        {
            var template = _workbook.Template ?? throw new InvalidOperationException("Worksheets.Edit is for workbooks opened with ExcelPackage.Edit (or OpenTemplate). Use Worksheets.Add to create sheets.");
            return OpenCore(template, FindTemplateSheet(template, name), edit: true);
        }

        /// <summary>Opens a sheet for editing by its zero-based position in the workbook's tab order. See <see cref="Edit(string)"/>.</summary>
        public Worksheet Edit(int index)
        {
            var template = _workbook.Template ?? throw new InvalidOperationException("Worksheets.Edit is for workbooks opened with ExcelPackage.Edit (or OpenTemplate). Use Worksheets.Add to create sheets.");
            if (index < 0 || index >= template.Sheets.Count) throw new ArgumentOutOfRangeException(nameof(index), index, $"The workbook has {template.Sheets.Count} sheet(s).");
            return OpenCore(template, index, edit: true);
        }

        private static int FindTemplateSheet(TemplateSource template, string name)
        {
            var position = template.Sheets.FindIndex(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
            if (position >= 0) return position;
            var names = new List<string>();
            foreach (var s in template.Sheets) names.Add(s.Name);
            throw new ArgumentException($"The workbook has no sheet named '{name}'. Sheets: {string.Join(", ", names)}.", nameof(name));
        }

        /// <summary>
        /// Opens template sheet <paramref name="position"/> for streaming (template mode) or, with
        /// <paramref name="edit"/>, with its cells loaded (edit mode). The size check runs before the current
        /// sheet is sealed so a refused sheet leaves the package exactly as it was.
        /// </summary>
        private Worksheet OpenCore(TemplateSource template, int position, bool edit)
        {
            if (_workbook.Package.IsSaved) throw new InvalidOperationException("Cannot open a worksheet after the package has been saved.");
            var (sheetName, part) = template.Sheets[position];
            if (_names.Contains(sheetName))
                throw new InvalidOperationException($"Sheet '{sheetName}' has already been opened. Each sheet of the workbook can be opened once.");
            if (edit) Worksheet.EnsureEditable(template, part, sheetName, _workbook.Package.Options.MaxEditableCells);
            _names.Add(sheetName);

            Current?.Seal();
            var writer = _workbook.Package.OpenPart(part);
            var sheet = new Worksheet(_workbook, position + 1, sheetName, writer);
            sheet.AttachTemplate(template, part);
            _sheets.Add(sheet);
            if (edit)
            {
                // The sheet's part is already open in the archive; a load failure (damaged XML) must stop the
                // package from saving a half-loaded sheet when the caller's using block disposes it.
                try { sheet.LoadTemplateCells(); }
                catch { _workbook.Package.MarkSaveFailed(); throw; }
            }
            return sheet;
        }

        /// <summary>Seals the current sheet. Called by the package before writing the trailing parts.</summary>
        internal void SealAll() => Current?.Seal();

        private void ValidateName(string name)
        {
            if (name == null) throw new ArgumentNullException(nameof(name));
            if (name.Length == 0) throw new ArgumentException("Sheet name cannot be empty.", nameof(name));
            if (name.Length > 31) throw new ArgumentException($"Sheet name '{name}' is longer than Excel's limit of 31 characters.", nameof(name));
            if (name.IndexOfAny(ForbiddenNameChars) >= 0) throw new ArgumentException($"Sheet name '{name}' contains a forbidden character (: \\ / ? * [ ]).", nameof(name));
            if (name[0] == '\'' || name[name.Length - 1] == '\'') throw new ArgumentException($"Sheet name '{name}' cannot start or end with an apostrophe.", nameof(name));
            if (_names.Contains(name)) throw new ArgumentException($"A sheet named '{name}' already exists (names are case-insensitive).", nameof(name));
        }

        /// <inheritdoc />
        public IEnumerator<Worksheet> GetEnumerator() => _sheets.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
