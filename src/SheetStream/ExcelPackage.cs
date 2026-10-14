using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>
    /// The .xlsx container. Owns the zip archive, streams each worksheet into it as the sheet is written,
    /// and on <see cref="Save"/> emits the trailing parts (shared strings, styles, workbook, relationships
    /// and content types) that can only be known once every sheet is complete.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The package writes straight to the destination stream and never holds a sheet in memory. The price
    /// of that design is forward-only access: a <see cref="ZipArchive"/> in create mode can hold a single
    /// entry open at a time, so adding a worksheet seals the previous one, and rows must be written in
    /// ascending order. Column widths and frozen panes must be declared before the first row because the
    /// schema places them ahead of the sheet data.
    /// </para>
    /// <para>
    /// Disposing the package without calling <see cref="Save"/> saves it, so a plain <c>using</c> block
    /// always produces a complete file. Not thread-safe.
    /// </para>
    /// </remarks>
    /// <example>
    /// <code>
    /// using var package = ExcelPackage.Create("report.xlsx");
    /// var sheet = package.Workbook.Worksheets.Add("Data");
    /// sheet.Row().WriteCell("Hello").WriteCell(42);
    /// package.Save();
    /// </code>
    /// </example>
    public sealed partial class ExcelPackage : IDisposable
    {
        private static readonly XmlWriterSettings WriterSettings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            CloseOutput = true,
            Indent = false,
            // Leave newlines inside cell text untouched; Excel shows them as line breaks.
            NewLineHandling = NewLineHandling.None,
            OmitXmlDeclaration = false,
        };

        private readonly Stream _output;
        private readonly bool _leaveOpen;
        private ZipArchive? _archive;
        private XmlWriter? _openPart;
        private bool _saved;
        private bool _saveFailed;
        private bool _disposed;

        // Parts beyond the fixed set (tables, drawings, charts, comments) and extension defaults (png, vml...).
        private readonly List<(string Part, string ContentType)> _overrides = new List<(string, string)>();
        private readonly Dictionary<string, string> _defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Creates a package that writes to <paramref name="output"/>.
        /// </summary>
        /// <param name="output">Destination stream. It only needs to support forward writes.</param>
        /// <param name="leaveOpen">
        /// <c>true</c> to leave <paramref name="output"/> open after the package is saved, for example when
        /// writing into a <see cref="MemoryStream"/> you still need to read.
        /// </param>
        /// <param name="options">Optional tuning settings.</param>
        public ExcelPackage(Stream output, bool leaveOpen = false, PackageOptions? options = null)
        {
            _output = output ?? throw new ArgumentNullException(nameof(output));
            if (!output.CanWrite) throw new ArgumentException("The output stream must be writable.", nameof(output));
            _leaveOpen = leaveOpen;
            Options = options ?? PackageOptions.Default;
            _archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);
            Workbook = new Workbook(this);
        }

        /// <summary>Creates a package that writes to a file, replacing any existing file at <paramref name="path"/>.</summary>
        public static ExcelPackage Create(string path, PackageOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Path cannot be empty.", nameof(path));
            var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16);
            try
            {
                return new ExcelPackage(stream, leaveOpen: false, options);
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Builds a complete workbook in memory and returns its bytes. Convenient for web endpoints that
        /// return a download, mirroring EPPlus' <c>GetAsByteArray()</c>.
        /// </summary>
        public static byte[] CreateBytes(Action<Workbook> build, PackageOptions? options = null)
        {
            if (build == null) throw new ArgumentNullException(nameof(build));
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage(ms, leaveOpen: true, options))
                {
                    build(package.Workbook);
                    package.Save();
                }
                return ms.ToArray();
            }
        }

        /// <summary>The single workbook of this package.</summary>
        public Workbook Workbook { get; }

        /// <summary>The options this package was created with.</summary>
        public PackageOptions Options { get; }

        /// <summary>True once <see cref="Save"/> has completed.</summary>
        public bool IsSaved => _saved;

        /// <summary>
        /// Seals the current worksheet, writes the shared strings, styles, workbook, relationship and
        /// content-type parts, and closes the zip archive. Safe to call more than once; later calls do nothing.
        /// After saving, no further worksheets or rows can be written.
        /// </summary>
        public void Save()
        {
            if (_saved) return;
            if (_archive == null) throw new ObjectDisposedException(nameof(ExcelPackage));
            try
            {
                SaveCore();
                CompleteInPlaceSave();
            }
            catch
            {
                // Remember the failure so Dispose does not retry and throw a second, misleading exception.
                _saveFailed = true;
                throw;
            }
        }

        /// <summary>
        /// Marks the package as unsaveable after a failure while a part was being built (a sheet that could not
        /// be loaded for editing), so disposing it does not write a half-finished workbook.
        /// </summary>
        internal void MarkSaveFailed() => _saveFailed = true;

        private void SaveCore()
        {
            if (Workbook.Template != null)
            {
                SaveTemplate(Workbook.Template);
                _saved = true;
                return;
            }

            // Excel refuses a workbook without sheets, and one whose sheets are all hidden.
            if (Workbook.Worksheets.Count == 0) Workbook.Worksheets.Add("Sheet1");
            var anyVisible = false;
            foreach (var sheet in Workbook.Worksheets) anyVisible |= sheet.Visibility == SheetVisibility.Visible;
            if (!anyVisible) throw new InvalidOperationException("At least one worksheet must stay visible; Excel cannot open a workbook whose sheets are all hidden.");

            Workbook.Worksheets.SealAll();

            var savedUtc = DateTime.UtcNow;
            using (var w = OpenPart("docProps/core.xml")) Workbook.Properties.WriteCore(w, savedUtc);
            using (var w = OpenPart("docProps/app.xml")) Workbook.Properties.WriteApp(w);
            using (var w = OpenPart("xl/sharedStrings.xml")) Workbook.SharedStrings.WriteTo(w);
            using (var w = OpenPart("xl/styles.xml")) Workbook.Styles.WriteTo(w);
            using (var w = OpenPart("xl/workbook.xml")) WriteWorkbookPart(w);
            using (var w = OpenPart("xl/_rels/workbook.xml.rels")) WriteWorkbookRelationships(w);
            using (var w = OpenPart("_rels/.rels")) WriteRootRelationships(w);
            using (var w = OpenPart("[Content_Types].xml")) WriteContentTypes(w);

            _archive!.Dispose(); // Writes the zip central directory.
            _archive = null;
            _output.Flush();
            _saved = true;
        }

        /// <summary>Saves the package if it has not been saved yet, then releases the output stream unless it was left open.</summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                // A save that already failed is not retried: its exception is the one the caller needs to see.
                if (!_saved && !_saveFailed && _archive != null) Save();
            }
            finally
            {
                _openPart?.Dispose();
                _archive?.Dispose();
                _archive = null;
                if (!_leaveOpen) _output.Dispose();
                DiscardInPlaceTemp(); // Only has something to do when an in-place save did not complete.
            }
        }

        /// <summary>
        /// Creates a zip entry and returns an <see cref="XmlWriter"/> positioned at its start. Only one part
        /// may be open at a time; disposing the writer closes the entry.
        /// </summary>
        internal XmlWriter OpenPart(string entryName)
        {
            if (_archive == null) throw new ObjectDisposedException(nameof(ExcelPackage));
            if (_openPart != null) throw new InvalidOperationException($"Cannot open part '{entryName}' while another part is still being written.");

            var entry = _archive.CreateEntry(entryName, Options.CompressionLevel);
            var buffered = new BufferedStream(entry.Open(), Options.BufferSize);
            var writer = XmlWriter.Create(buffered, WriterSettings);
            _openPart = new PartWriter(writer, this);
            return _openPart;
        }

        private void PartClosed() => _openPart = null;

        /// <summary>Declares the content type of a part written outside the fixed set (path without leading slash).</summary>
        internal void RegisterPart(string partName, string contentType) => _overrides.Add(("/" + partName, contentType));

        /// <summary>Declares the content type for every part with a given file extension.</summary>
        internal void RegisterDefault(string extension, string contentType) => _defaults[extension] = contentType;

        /// <summary>
        /// Writes a complete binary (or pre-built text) part in one go. Images are already compressed, so they
        /// are stored without deflate to save time.
        /// </summary>
        internal void WriteBinaryPart(string entryName, byte[] data)
        {
            if (_archive == null) throw new ObjectDisposedException(nameof(ExcelPackage));
            if (_openPart != null) throw new InvalidOperationException($"Cannot write part '{entryName}' while another part is still being written.");
            var isImage = entryName.EndsWith(".png", StringComparison.OrdinalIgnoreCase) || entryName.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) || entryName.EndsWith(".gif", StringComparison.OrdinalIgnoreCase);
            var entry = _archive.CreateEntry(entryName, isImage ? CompressionLevel.NoCompression : Options.CompressionLevel);
            using (var stream = entry.Open()) stream.Write(data, 0, data.Length);
        }

        /// <summary><c>xl/workbook.xml</c>: protection, views, the sheet list, defined names and calc settings.</summary>
        private void WriteWorkbookPart(XmlWriter w)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("workbook", OpenXmlNamespaces.Main);
            w.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            // CT_Workbook order: workbookProtection, bookViews, sheets, definedNames, calcPr.
            if (Workbook.Protection is var (hash, salt, structure, windows))
            {
                w.WriteStartElement("workbookProtection");
                if (hash != null)
                {
                    w.WriteAttributeString("workbookAlgorithmName", "SHA-512");
                    w.WriteAttributeString("workbookHashValue", hash);
                    w.WriteAttributeString("workbookSaltValue", salt);
                    w.WriteAttributeString("workbookSpinCount", PasswordHasher.SpinCount.ToString(CultureInfo.InvariantCulture));
                }
                if (structure) w.WriteAttributeString("lockStructure", "1");
                if (windows) w.WriteAttributeString("lockWindows", "1");
                w.WriteEndElement();
            }

            // The active tab must be a visible sheet, otherwise Excel opens on a hidden one.
            var activeTab = 0;
            foreach (var sheet in Workbook.Worksheets)
            {
                if (sheet.Visibility == SheetVisibility.Visible) { activeTab = sheet.Index - 1; break; }
            }
            w.WriteStartElement("bookViews");
            w.WriteStartElement("workbookView");
            if (activeTab > 0)
            {
                w.WriteAttributeString("firstSheet", activeTab.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("activeTab", activeTab.ToString(CultureInfo.InvariantCulture));
            }
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("sheets");
            foreach (var sheet in Workbook.Worksheets)
            {
                w.WriteStartElement("sheet");
                w.WriteAttributeString("name", sheet.Name);
                w.WriteAttributeString("sheetId", sheet.Index.ToString(CultureInfo.InvariantCulture));
                if (sheet.Visibility == SheetVisibility.Hidden) w.WriteAttributeString("state", "hidden");
                else if (sheet.Visibility == SheetVisibility.VeryHidden) w.WriteAttributeString("state", "veryHidden");
                w.WriteAttributeString("id", OpenXmlNamespaces.Relationships, SheetRelationshipId(sheet.Index));
                w.WriteEndElement();
            }
            w.WriteEndElement();

            // Built-in names first. Excel registers each AutoFilter through a hidden _FilterDatabase name;
            // without it some versions report the file as needing repair.
            var names = new List<(string Name, int? LocalSheet, bool Hidden, string Formula)>();
            foreach (var sheet in Workbook.Worksheets)
            {
                var local = sheet.Index - 1;
                if (sheet.AutoFilterRange != null)
                    names.Add(("_xlnm._FilterDatabase", local, true, RangeReference.QuoteSheetName(sheet.Name) + "!" + sheet.AutoFilterRange));
                if (sheet.PageSetup.PrintAreaFormula is string area) names.Add(("_xlnm.Print_Area", local, false, area));
                if (sheet.PageSetup.PrintTitlesFormula is string titles) names.Add(("_xlnm.Print_Titles", local, false, titles));
            }
            foreach (var dn in Workbook.DefinedNames)
                names.Add((dn.Name, dn.Scope == null ? (int?)null : dn.Scope.Index - 1, dn.Hidden, dn.Formula));
            if (names.Count > 0)
            {
                w.WriteStartElement("definedNames");
                foreach (var (name, local, hidden, formula) in names)
                {
                    w.WriteStartElement("definedName");
                    w.WriteAttributeString("name", name);
                    if (local.HasValue) w.WriteAttributeString("localSheetId", local.Value.ToString(CultureInfo.InvariantCulture));
                    if (hidden) w.WriteAttributeString("hidden", "1");
                    w.WriteString(XmlTextUtil.Sanitize(formula));
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }

            // Formulas are written without cached values, so ask Excel to calculate everything on open.
            if (Workbook.HasFormulas)
            {
                w.WriteStartElement("calcPr");
                w.WriteAttributeString("fullCalcOnLoad", "1");
                w.WriteEndElement();
            }

            w.WriteEndElement();
            w.WriteEndDocument();
        }

        /// <summary><c>xl/_rels/workbook.xml.rels</c>: one relationship per sheet, then styles and shared strings.</summary>
        private void WriteWorkbookRelationships(XmlWriter w)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("Relationships", OpenXmlNamespaces.PackageRelationships);
            foreach (var sheet in Workbook.Worksheets)
                WriteRelationship(w, SheetRelationshipId(sheet.Index), OpenXmlNamespaces.RelWorksheet, "worksheets/sheet" + sheet.Index.ToString(CultureInfo.InvariantCulture) + ".xml");
            var next = Workbook.Worksheets.Count + 1;
            WriteRelationship(w, "rId" + next.ToString(CultureInfo.InvariantCulture), OpenXmlNamespaces.RelStyles, "styles.xml");
            WriteRelationship(w, "rId" + (next + 1).ToString(CultureInfo.InvariantCulture), OpenXmlNamespaces.RelSharedStrings, "sharedStrings.xml");
            w.WriteEndElement();
            w.WriteEndDocument();
        }

        /// <summary><c>_rels/.rels</c>: points the package at the workbook part.</summary>
        private static void WriteRootRelationships(XmlWriter w)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("Relationships", OpenXmlNamespaces.PackageRelationships);
            WriteRelationship(w, "rId1", OpenXmlNamespaces.RelOfficeDocument, "xl/workbook.xml");
            WriteRelationship(w, "rId2", OpenXmlNamespaces.RelCoreProperties, "docProps/core.xml");
            WriteRelationship(w, "rId3", OpenXmlNamespaces.RelExtendedProperties, "docProps/app.xml");
            w.WriteEndElement();
            w.WriteEndDocument();
        }

        private static void WriteRelationship(XmlWriter w, string id, string type, string target)
        {
            w.WriteStartElement("Relationship");
            w.WriteAttributeString("Id", id);
            w.WriteAttributeString("Type", type);
            w.WriteAttributeString("Target", target);
            w.WriteEndElement();
        }

        /// <summary><c>[Content_Types].xml</c>: declares the media type of every part in the package.</summary>
        private void WriteContentTypes(XmlWriter w)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("Types", OpenXmlNamespaces.ContentTypes);

            w.WriteStartElement("Default");
            w.WriteAttributeString("Extension", "rels");
            w.WriteAttributeString("ContentType", OpenXmlNamespaces.CtRelationships);
            w.WriteEndElement();
            w.WriteStartElement("Default");
            w.WriteAttributeString("Extension", "xml");
            w.WriteAttributeString("ContentType", OpenXmlNamespaces.CtXml);
            w.WriteEndElement();
            foreach (var pair in _defaults)
            {
                w.WriteStartElement("Default");
                w.WriteAttributeString("Extension", pair.Key);
                w.WriteAttributeString("ContentType", pair.Value);
                w.WriteEndElement();
            }

            WriteOverride(w, "/xl/workbook.xml", OpenXmlNamespaces.CtWorkbook);
            foreach (var sheet in Workbook.Worksheets)
                WriteOverride(w, "/xl/worksheets/sheet" + sheet.Index.ToString(CultureInfo.InvariantCulture) + ".xml", OpenXmlNamespaces.CtWorksheet);
            WriteOverride(w, "/xl/styles.xml", OpenXmlNamespaces.CtStyles);
            WriteOverride(w, "/xl/sharedStrings.xml", OpenXmlNamespaces.CtSharedStrings);
            WriteOverride(w, "/docProps/core.xml", OpenXmlNamespaces.CtCoreProperties);
            WriteOverride(w, "/docProps/app.xml", OpenXmlNamespaces.CtExtendedProperties);
            foreach (var (part, contentType) in _overrides) WriteOverride(w, part, contentType);

            w.WriteEndElement();
            w.WriteEndDocument();
        }

        private static void WriteOverride(XmlWriter w, string partName, string contentType)
        {
            w.WriteStartElement("Override");
            w.WriteAttributeString("PartName", partName);
            w.WriteAttributeString("ContentType", contentType);
            w.WriteEndElement();
        }

        private static string SheetRelationshipId(int sheetIndex) => "rId" + sheetIndex.ToString(CultureInfo.InvariantCulture);

        /// <summary>
        /// Thin wrapper that forwards to the real writer and tells the package when the entry is closed, so
        /// the one-open-part invariant can be enforced without the sheets knowing about the archive.
        /// </summary>
        private sealed class PartWriter : XmlWriter
        {
            private readonly XmlWriter _inner;
            private readonly ExcelPackage _owner;
            private bool _closed;

            public PartWriter(XmlWriter inner, ExcelPackage owner)
            {
                _inner = inner;
                _owner = owner;
            }

            public override WriteState WriteState => _inner.WriteState;
            public override void Flush() => _inner.Flush();
            public override string? LookupPrefix(string ns) => _inner.LookupPrefix(ns);
            public override void WriteBase64(byte[] buffer, int index, int count) => _inner.WriteBase64(buffer, index, count);
            public override void WriteCData(string? text) => _inner.WriteCData(text);
            public override void WriteCharEntity(char ch) => _inner.WriteCharEntity(ch);
            public override void WriteChars(char[] buffer, int index, int count) => _inner.WriteChars(buffer, index, count);
            public override void WriteComment(string? text) => _inner.WriteComment(text);
            public override void WriteDocType(string name, string? pubid, string? sysid, string? subset) => _inner.WriteDocType(name, pubid, sysid, subset);
            public override void WriteEndAttribute() => _inner.WriteEndAttribute();
            public override void WriteEndDocument() => _inner.WriteEndDocument();
            public override void WriteEndElement() => _inner.WriteEndElement();
            public override void WriteEntityRef(string name) => _inner.WriteEntityRef(name);
            public override void WriteFullEndElement() => _inner.WriteFullEndElement();
            public override void WriteProcessingInstruction(string name, string? text) => _inner.WriteProcessingInstruction(name, text);
            public override void WriteRaw(char[] buffer, int index, int count) => _inner.WriteRaw(buffer, index, count);
            public override void WriteRaw(string data) => _inner.WriteRaw(data);
            public override void WriteStartAttribute(string? prefix, string localName, string? ns) => _inner.WriteStartAttribute(prefix, localName, ns);
            public override void WriteStartDocument() => _inner.WriteStartDocument();
            public override void WriteStartDocument(bool standalone) => _inner.WriteStartDocument(standalone);
            public override void WriteStartElement(string? prefix, string localName, string? ns) => _inner.WriteStartElement(prefix, localName, ns);
            public override void WriteString(string? text) => _inner.WriteString(text);
            public override void WriteSurrogateCharEntity(char lowChar, char highChar) => _inner.WriteSurrogateCharEntity(lowChar, highChar);
            public override void WriteWhitespace(string? ws) => _inner.WriteWhitespace(ws);
            public override void WriteValue(int value) => _inner.WriteValue(value);
            public override void WriteValue(long value) => _inner.WriteValue(value);
            public override void WriteValue(double value) => _inner.WriteValue(value);
            public override void WriteValue(float value) => _inner.WriteValue(value);
            public override void WriteValue(decimal value) => _inner.WriteValue(value);
            public override void WriteValue(bool value) => _inner.WriteValue(value);
            public override void WriteValue(string? value) => _inner.WriteValue(value);

            protected override void Dispose(bool disposing)
            {
                if (_closed) return;
                _closed = true;
                if (disposing)
                {
                    _inner.Dispose(); // CloseOutput=true: also closes the buffered stream and the zip entry.
                    _owner.PartClosed();
                }
                base.Dispose(disposing);
            }
        }
    }
}
