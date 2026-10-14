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

        /// <summary>SpreadsheetML drawing namespace (anchors of pictures and charts on a sheet).</summary>
        public const string SpreadsheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

        /// <summary>DrawingML main namespace (shapes, fills, text).</summary>
        public const string DrawingMain = "http://schemas.openxmlformats.org/drawingml/2006/main";

        /// <summary>DrawingML chart namespace.</summary>
        public const string Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        // Relationship types (the Type attribute of a <Relationship/> element).
        public const string RelOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public const string RelWorksheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        public const string RelStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        public const string RelSharedStrings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        public const string RelHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        public const string RelTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
        public const string RelDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
        public const string RelImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        public const string RelChart = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
        public const string RelComments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
        public const string RelVmlDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
        public const string RelPivotTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
        public const string RelPivotCacheDefinition = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
        public const string RelPivotCacheRecords = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
        public const string RelCoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        public const string RelExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

        // Content types (the ContentType attribute of <Override/> and <Default/> elements).
        public const string CtRelationships = "application/vnd.openxmlformats-package.relationships+xml";
        public const string CtXml = "application/xml";
        public const string CtWorkbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        public const string CtWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        public const string CtStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        public const string CtSharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        public const string CtTable = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
        public const string CtDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";
        public const string CtChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
        public const string CtComments = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
        public const string CtVmlDrawing = "application/vnd.openxmlformats-officedocument.vmlDrawing";
        public const string CtPivotTable = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
        public const string CtPivotCacheDefinition = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
        public const string CtPivotCacheRecords = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
        public const string CtCoreProperties = "application/vnd.openxmlformats-package.core-properties+xml";
        public const string CtExtendedProperties = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    }
}
