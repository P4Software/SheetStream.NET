using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;

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
