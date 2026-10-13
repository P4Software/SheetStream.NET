namespace SheetStream.Internal
{
    /// <summary>
    /// The ECMA-376 namespace, relationship-type and content-type constants used across every part of the
    /// package. Centralised so a single typo cannot silently produce a file Excel refuses to open.
    /// </summary>
    internal static class OpenXmlNamespaces
    {
        /// <summary>SpreadsheetML main namespace (worksheet, workbook, styles, shared strings).</summary>
        public const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        /// <summary>Relationship-reference namespace used for the <c>r:id</c> attribute.</summary>
        public const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>Namespace of the <c>.rels</c> part documents.</summary>
        public const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>Namespace of <c>[Content_Types].xml</c>.</summary>
        public const string ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

        // Relationship types (the Type attribute of a <Relationship/> element).
        public const string RelOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public const string RelWorksheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        public const string RelStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        public const string RelSharedStrings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

        // Content types (the ContentType attribute of <Override/> and <Default/> elements).
        public const string CtRelationships = "application/vnd.openxmlformats-package.relationships+xml";
        public const string CtXml = "application/xml";
        public const string CtWorkbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        public const string CtWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        public const string CtStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        public const string CtSharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    }
}
