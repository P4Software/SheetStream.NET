using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>Visibility of a worksheet tab.</summary>
    public enum SheetVisibility
    {
        /// <summary>Visible (default).</summary>
        Visible,
        /// <summary>Hidden; users can unhide it from Excel's menu.</summary>
        Hidden,
        /// <summary>Hidden and not listed in Excel's Unhide dialog; only code can show it again.</summary>
        VeryHidden,
    }

    /// <summary>Document properties shown in File, Info and written to <c>docProps/core.xml</c> / <c>app.xml</c>.</summary>
    public sealed class WorkbookProperties
    {
        /// <summary>Title.</summary>
        public string? Title { get; set; }
        /// <summary>Subject.</summary>
        public string? Subject { get; set; }
        /// <summary>Author (creator).</summary>
        public string? Author { get; set; }
        /// <summary>Keywords / tags.</summary>
        public string? Keywords { get; set; }
        /// <summary>Comments / description.</summary>
        public string? Description { get; set; }
        /// <summary>Category.</summary>
        public string? Category { get; set; }
        /// <summary>Company.</summary>
        public string? Company { get; set; }
        /// <summary>Creation time; defaults to the moment the package is saved.</summary>
        public DateTime? Created { get; set; }

        /// <summary>Writes <c>docProps/core.xml</c> (Dublin Core metadata).</summary>
        internal void WriteCore(XmlWriter w, DateTime savedUtc)
        {
            const string cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
            const string dc = "http://purl.org/dc/elements/1.1/";
            const string dcterms = "http://purl.org/dc/terms/";
            const string xsi = "http://www.w3.org/2001/XMLSchema-instance";

            w.WriteStartDocument(true);
            w.WriteStartElement("cp", "coreProperties", cp);
            w.WriteAttributeString("xmlns", "dc", null, dc);
            w.WriteAttributeString("xmlns", "dcterms", null, dcterms);
            w.WriteAttributeString("xmlns", "dcmitype", null, "http://purl.org/dc/dcmitype/");
            w.WriteAttributeString("xmlns", "xsi", null, xsi);
            if (Title != null) w.WriteElementString("dc", "title", dc, XmlTextUtil.Sanitize(Title));
            if (Subject != null) w.WriteElementString("dc", "subject", dc, XmlTextUtil.Sanitize(Subject));
            if (Author != null) w.WriteElementString("dc", "creator", dc, XmlTextUtil.Sanitize(Author));
            if (Keywords != null) w.WriteElementString("cp", "keywords", cp, XmlTextUtil.Sanitize(Keywords));
            if (Description != null) w.WriteElementString("dc", "description", dc, XmlTextUtil.Sanitize(Description));
            if (Author != null) w.WriteElementString("cp", "lastModifiedBy", cp, XmlTextUtil.Sanitize(Author));
            var created = (Created?.ToUniversalTime() ?? savedUtc).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
            w.WriteStartElement("dcterms", "created", dcterms);
            w.WriteAttributeString("xsi", "type", xsi, "dcterms:W3CDTF");
            w.WriteString(created);
            w.WriteEndElement();
            w.WriteStartElement("dcterms", "modified", dcterms);
            w.WriteAttributeString("xsi", "type", xsi, "dcterms:W3CDTF");
            w.WriteString(savedUtc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture));
            w.WriteEndElement();
            if (Category != null) w.WriteElementString("cp", "category", cp, XmlTextUtil.Sanitize(Category));
            w.WriteEndElement();
            w.WriteEndDocument();
        }

        /// <summary>Writes <c>docProps/app.xml</c> (application name and company).</summary>
        internal void WriteApp(XmlWriter w)
        {
            const string ns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
            w.WriteStartDocument(true);
            w.WriteStartElement("Properties", ns);
            w.WriteElementString("Application", ns, "SheetStream");
            if (Company != null) w.WriteElementString("Company", ns, XmlTextUtil.Sanitize(Company));
            w.WriteEndElement();
            w.WriteEndDocument();
        }
    }

    /// <summary>A workbook-level or sheet-level defined name (named range or named formula).</summary>
    public sealed class DefinedName
    {
        internal DefinedName(string name, string formula, Worksheet? scope, bool hidden)
        {
            Name = name;
            Formula = formula;
            Scope = scope;
            Hidden = hidden;
        }

        /// <summary>The name, e.g. <c>TaxRate</c>.</summary>
        public string Name { get; }

        /// <summary>What it refers to, without the leading "=", e.g. <c>'Settings'!$B$2</c>.</summary>
        public string Formula { get; }

        /// <summary>The sheet it is local to, or null for workbook scope.</summary>
        public Worksheet? Scope { get; }

        /// <summary>Hidden from Excel's Name Manager.</summary>
        public bool Hidden { get; }
    }

    /// <summary>Defined names of a workbook. Add them any time before the package is saved.</summary>
    public sealed class DefinedNameCollection : IReadOnlyList<DefinedName>
    {
        private readonly List<DefinedName> _names = new List<DefinedName>();

        internal DefinedNameCollection()
        {
        }

        /// <inheritdoc />
        public int Count => _names.Count;

        /// <inheritdoc />
        public DefinedName this[int index] => _names[index];

        /// <summary>
        /// Adds a name. <paramref name="refersTo"/> is a range ("'Settings'!B2", "Data!A1:A10") or any formula
        /// ("=0.07"). Plain ranges are made absolute and sheet-qualified with <paramref name="scope"/>'s name when
        /// they lack a sheet.
        /// </summary>
        public DefinedName Add(string name, string refersTo, Worksheet? scope = null, bool hidden = false)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Name cannot be empty.", nameof(name));
            if (name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Built-in names (_xlnm.*) are managed by SheetStream; use PageSetup for print areas and titles.", nameof(name));
            ExcelTable.ValidateName(name); // Same lexical rules as table names.
            foreach (var n in _names)
                if (string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase) && n.Scope == scope)
                    throw new ArgumentException($"The name '{name}' is already defined in this scope.", nameof(name));
            if (string.IsNullOrWhiteSpace(refersTo)) throw new ArgumentException("refersTo cannot be empty.", nameof(refersTo));

            var formula = refersTo.Trim().TrimStart('=');
            if (formula.IndexOf('!') >= 0)
            {
                // "Sheet!B2" style: normalise to an absolute, quoted reference when it is a plain range.
                try { formula = RangeReference.QualifyAbsolute(formula, ""); }
                catch (ArgumentException) { /* A formula mentioning a sheet: keep as written. */ }
            }
            else if (IsPlainRange(formula))
            {
                if (scope == null)
                    throw new ArgumentException($"'{refersTo}' has no sheet. Write it as \"'Sheet'!{formula}\" or pass the sheet as scope.", nameof(refersTo));
                formula = RangeReference.Parse(formula).ToAbsolute(scope.Name);
            }
            // Anything else (a constant such as 0.07, or a calculation) is kept as written.
            var dn = new DefinedName(name, formula, scope, hidden);
            _names.Add(dn);
            return dn;
        }

        private static bool IsPlainRange(string text)
        {
            try
            {
                RangeReference.Parse(text);
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        /// <inheritdoc />
        public IEnumerator<DefinedName> GetEnumerator() => _names.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
