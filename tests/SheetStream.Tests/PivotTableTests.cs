// The Open XML SDK exposes nearly every member as nullable; in test assertions a null is a failure anyway.
#pragma warning disable CS8602, CS8603, CS8604
using System.IO.Compression;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SheetStream.Tests;

/// <summary>
/// Pivot tables: package structure (parts, relationships, content types, workbook cache list), schema validity,
/// field mapping errors, the computed location for each field combination, and a round trip of the cache
/// definition through the Open XML SDK. Whether Excel really refreshes the result is covered by
/// <c>tests/excel-check-pivot.ps1</c>, which needs an Excel installation.
/// </summary>
public class PivotTableTests
{
    private static readonly string[] Headers = { "Region", "Category", "Product", "Amount", "Qty" };

    /// <summary>Writes a small sales sheet and returns its range so the pivot can point at it.</summary>
    private static string WriteSales(Workbook wb, string sheetName = "Sales")
    {
        WriteSalesCore(wb, sheetName, null);
        return "A1:E13";
    }

    /// <summary>Writes the same sales sheet as an Excel table and returns the table.</summary>
    private static ExcelTable WriteSalesTable(Workbook wb, string tableName) => WriteSalesCore(wb, "Sales", tableName)!;

    private static ExcelTable? WriteSalesCore(Workbook wb, string sheetName, string? tableName)
    {
        var s = wb.Worksheets.Add(sheetName);
        var table = tableName == null ? null : s.AddTable(tableName, "A1:E13"); // Before the header row, so the names are captured from it.
        var header = s.Row();
        foreach (var h in Headers) header.WriteCell(h);
        var regions = new[] { "North", "South" };
        var categories = new[] { "Hardware", "Software" };
        for (var i = 0; i < 12; i++)
            s.Row().WriteCell(regions[i % 2]).WriteCell(categories[i / 6]).WriteCell("P" + (i % 4)).WriteCell(100 + i * 10.5).WriteCell(i + 1);
        return table;
    }

    private static SpreadsheetDocument OpenValidated(byte[] bytes)
    {
        var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc).ToList();
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Path?.PartUri} {e.Path?.XPath}: {e.Description}")));
        return doc;
    }

    private static string ReadEntry(byte[] bytes, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var entry = zip.GetEntry(name);
        Assert.True(entry != null, $"Part {name} is missing. Parts: {string.Join(", ", zip.Entries.Select(e => e.FullName))}");
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    [Fact]
    public void RangeSource_WritesAllPartsRelationshipsAndContentTypes()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var range = WriteSales(wb);
            var p = wb.Worksheets.Add("Pivot").AddPivotTable("SalesByRegion", "Sales", range, 1, 1, Headers);
            p.AddRowField("Region").AddColumnField("Category").AddDataField("Amount", numberFormat: "#,##0.0 \"units\"");
        });

        using var doc = OpenValidated(bytes);
        var wbPart = doc.WorkbookPart;

        // Workbook: one cache listed, related to the definition part.
        var caches = wbPart.Workbook.GetFirstChild<X.PivotCaches>().Elements<X.PivotCache>().ToList();
        Assert.Single(caches);
        var cacheDef = (PivotTableCacheDefinitionPart)wbPart.GetPartById(caches[0].Id);
        Assert.Equal("/xl/pivotCache/pivotCacheDefinition1.xml", cacheDef.Uri.OriginalString);
        Assert.NotNull(cacheDef.PivotTableCacheRecordsPart);
        Assert.Equal("/xl/pivotCache/pivotCacheRecords1.xml", cacheDef.PivotTableCacheRecordsPart.Uri.OriginalString);

        // Sheet: the pivot part hangs off the pivot sheet only, and points back to the same cache.
        var pivotSheet = wbPart.Workbook.Sheets.Elements<X.Sheet>().Single(s => s.Name == "Pivot");
        var pivotSheetPart = (WorksheetPart)wbPart.GetPartById(pivotSheet.Id);
        var pivotPart = pivotSheetPart.PivotTableParts.Single();
        Assert.Equal("/xl/pivotTables/pivotTable1.xml", pivotPart.Uri.OriginalString);
        Assert.Same(cacheDef, pivotPart.PivotTableCacheDefinitionPart);
        var salesSheet = wbPart.Workbook.Sheets.Elements<X.Sheet>().Single(s => s.Name == "Sales");
        Assert.Empty(((WorksheetPart)wbPart.GetPartById(salesSheet.Id)).PivotTableParts);

        // Content types for the three part kinds.
        var contentTypes = ReadEntry(bytes, "[Content_Types].xml");
        Assert.Contains("PartName=\"/xl/pivotTables/pivotTable1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml\"", contentTypes);
        Assert.Contains("PartName=\"/xl/pivotCache/pivotCacheDefinition1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml\"", contentTypes);
        Assert.Contains("PartName=\"/xl/pivotCache/pivotCacheRecords1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml\"", contentTypes);

        // Definition content.
        var def = pivotPart.PivotTableDefinition;
        Assert.Equal("SalesByRegion", def.Name.Value);
        Assert.Equal(caches[0].CacheId.Value, def.CacheId.Value);
        Assert.Equal("Values", def.DataCaption.Value);
        var fields = def.PivotFields.Elements<X.PivotField>().ToList();
        Assert.Equal(5, fields.Count);
        Assert.Equal(X.PivotTableAxisValues.AxisRow, fields[0].Axis.Value);
        Assert.Equal(X.PivotTableAxisValues.AxisColumn, fields[1].Axis.Value);
        Assert.True(fields[3].DataField.Value);
        Assert.Equal(0, def.RowFields.Elements<X.Field>().Single().Index.Value);
        Assert.Equal(1, def.ColumnFields.Elements<X.Field>().Single().Index.Value);
        var data = def.DataFields.Elements<X.DataField>().Single();
        Assert.Equal("Sum of Amount", data.Name.Value);
        Assert.Equal(3u, data.Field.Value);
        Assert.Equal(X.DataConsolidateFunctionValues.Sum, data.Subtotal.Value);
        // The custom number format landed in styles.xml under the id the data field references.
        var numFmt = wbPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Elements<X.NumberingFormat>().Single(n => n.NumberFormatId.Value == data.NumberFormatId.Value);
        Assert.Equal("#,##0.0 \"units\"", numFmt.FormatCode.Value);
        Assert.Equal("PivotStyleMedium9", def.PivotTableStyle.Name.Value);
    }

    [Fact]
    public void CacheDefinition_RoundTrips_RefreshOnLoadAndEmptyRecords()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var range = WriteSales(wb);
            wb.Worksheets.Add("Pivot").AddPivotTable("P", "Sales", range, 1, 1, Headers).AddRowField("Region").AddDataField("Qty", PivotFunction.Count);
        });

        using var doc = OpenValidated(bytes);
        var cacheDef = doc.WorkbookPart.PivotTableCacheDefinitionParts.Single().PivotCacheDefinition;
        Assert.True(cacheDef.RefreshOnLoad.Value);
        Assert.False(cacheDef.SaveData.Value);
        Assert.Equal(0u, cacheDef.RecordCount.Value);
        Assert.Equal("SheetStream", cacheDef.RefreshedBy.Value);
        Assert.Equal(6, (int)cacheDef.CreatedVersion.Value);
        Assert.Equal(6, (int)cacheDef.RefreshedVersion.Value);
        Assert.Equal(3, (int)cacheDef.MinRefreshableVersion.Value);
        var source = cacheDef.CacheSource.WorksheetSource;
        Assert.Equal("A1:E13", source.Reference.Value);
        Assert.Equal("Sales", source.Sheet.Value);
        Assert.Null(source.Name);
        Assert.Equal(Headers, cacheDef.CacheFields.Elements<X.CacheField>().Select(f => f.Name.Value));
        Assert.All(cacheDef.CacheFields.Elements<X.CacheField>(), f => Assert.Empty(f.SharedItems.ChildElements));
        var records = doc.WorkbookPart.PivotTableCacheDefinitionParts.Single().PivotTableCacheRecordsPart.PivotCacheRecords;
        Assert.Equal(0u, records.Count.Value);
        Assert.Empty(records.ChildElements);
    }

    [Fact]
    public void TableSource_UsesTableNameAndItsColumns()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var table = WriteSalesTable(wb, "SalesTbl");
            var pivot = wb.Worksheets.Add("Pivot").AddPivotTable("FromTable", table, 2, 2);
            Assert.Equal(Headers, pivot.Headers); // Captured from the header row that was written after AddTable.
            pivot.AddRowField("product").AddDataField("Amount", PivotFunction.Average, "Avg amount");
        });

        using var doc = OpenValidated(bytes);
        var source = doc.WorkbookPart.PivotTableCacheDefinitionParts.Single().PivotCacheDefinition.CacheSource.WorksheetSource;
        Assert.Equal("SalesTbl", source.Name.Value);
        Assert.Null(source.Reference);
        var def = doc.WorkbookPart.WorksheetParts.SelectMany(p => p.PivotTableParts).Single().PivotTableDefinition;
        Assert.Equal(2, def.RowFields.Elements<X.Field>().Single().Index.Value); // "product" matched case-insensitively.
        Assert.Equal("Avg amount", def.DataFields.Elements<X.DataField>().Single().Name.Value);
        Assert.Equal(X.DataConsolidateFunctionValues.Average, def.DataFields.Elements<X.DataField>().Single().Subtotal.Value);
    }

    [Fact]
    public void TableSource_OnSameSheet_NeedsHeadersFirst()
    {
        using var package = new ExcelPackage(new MemoryStream());
        var s = package.Workbook.Worksheets.Add("Sales");
        var table = s.AddTable("Tbl", "A1:E13");
        // The header row is not written yet, so the field names are unknown.
        var ex = Assert.Throws<InvalidOperationException>(() => s.AddPivotTable("P", table, 1, 8));
        Assert.Contains("header", ex.Message);
        var header = s.Row();
        foreach (var h in Headers) header.WriteCell(h);
        var pivot = s.AddPivotTable("P", table, 1, 8);
        Assert.Equal(Headers, pivot.Headers);
        pivot.AddRowField("Region").AddDataField("Amount"); // So the package can save when disposed.
        for (var i = 0; i < 12; i++) s.Row().WriteCell("N").WriteCell("C").WriteCell("P").WriteCell(1.0).WriteCell(1);
    }

    [Fact]
    public void TwoPivots_GetUniqueCacheIdsAndPartNumbers()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var range = WriteSales(wb);
            var sheet = wb.Worksheets.Add("Pivots");
            sheet.AddPivotTable("A", "Sales", range, 1, 1, Headers).AddRowField("Region").AddDataField("Amount");
            sheet.AddPivotTable("B", "Sales", range, 1, 8, Headers).AddRowField("Category").AddDataField("Qty");
            // A third one on another sheet, sourcing the same range.
            wb.Worksheets.Add("More").AddPivotTable("C", "Sales", range, 1, 1, Headers).AddFilterField("Region").AddRowField("Product").AddDataField("Amount");
        });

        using var doc = OpenValidated(bytes);
        var caches = doc.WorkbookPart.Workbook.GetFirstChild<X.PivotCaches>().Elements<X.PivotCache>().Select(c => c.CacheId.Value).ToList();
        Assert.Equal(3, caches.Count);
        Assert.Equal(3, caches.Distinct().Count());
        var defs = doc.WorkbookPart.WorksheetParts.SelectMany(p => p.PivotTableParts).ToList();
        Assert.Equal(3, defs.Count);
        Assert.Equal(caches.OrderBy(c => c), defs.Select(d => d.PivotTableDefinition.CacheId.Value).OrderBy(c => c));
        Assert.Equal(3, doc.WorkbookPart.PivotTableCacheDefinitionParts.Count());
        Assert.Equal(new[] { "/xl/pivotTables/pivotTable1.xml", "/xl/pivotTables/pivotTable2.xml", "/xl/pivotTables/pivotTable3.xml" }, defs.Select(d => d.Uri.OriginalString).OrderBy(u => u));
        // The definition is the right one for each sheet: pivot C carries the page field.
        var c = defs.Single(d => d.PivotTableDefinition.Name.Value == "C").PivotTableDefinition;
        Assert.Equal(0, c.PageFields.Elements<X.PageField>().Single().Field.Value);
        Assert.Equal(1u, c.Location.RowPageCount.Value);
    }

    [Fact]
    public void SameSheetSource_IsAllowed()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var range = WriteSales(wb);
            // The pivot lives next to its own data.
            wb.Worksheets.Current.AddPivotTable("Inline", "sales", range, 1, 8, Headers).AddRowField("Region").AddDataField("Amount");
        });
        using var doc = OpenValidated(bytes);
        Assert.Equal("Sales", doc.WorkbookPart.PivotTableCacheDefinitionParts.Single().PivotCacheDefinition.CacheSource.WorksheetSource.Sheet.Value);
    }

    [Fact]
    public void UnknownHeader_ThrowsNamingAvailableHeaders()
    {
        using var package = new ExcelPackage(new MemoryStream());
        var range = WriteSales(package.Workbook);
        var pivot = package.Workbook.Worksheets.Add("Pivot").AddPivotTable("P", "Sales", range, 1, 1, Headers);
        var ex = Assert.Throws<ArgumentException>(() => pivot.AddRowField("Reggion"));
        Assert.Contains("Reggion", ex.Message);
        Assert.Contains("Region, Category, Product, Amount, Qty", ex.Message);
        Assert.Throws<ArgumentException>(() => pivot.AddDataField("Total"));
        pivot.AddDataField("Amount"); // So the package can save when disposed.
    }

    [Fact]
    public void ArgumentErrors_AreReported()
    {
        using var package = new ExcelPackage(new MemoryStream());
        var wb = package.Workbook;
        var range = WriteSales(wb);
        var sheet = wb.Worksheets.Add("Pivot");
        Assert.Throws<ArgumentException>(() => sheet.AddPivotTable("P", "Missing", range, 1, 1, Headers));             // Unknown source sheet.
        Assert.Throws<ArgumentException>(() => sheet.AddPivotTable("P", "Sales", range, 1, 1, new[] { "A", "B" }));   // Header count mismatch.
        Assert.Throws<ArgumentException>(() => sheet.AddPivotTable("P", "Sales", "A1:E1", 1, 1, Headers));            // No data rows.
        Assert.Throws<ArgumentException>(() => sheet.AddPivotTable("P", "Sales", range, 1, 1, new[] { "A", "A", "C", "D", "E" })); // Duplicate headers.
        var pivot = sheet.AddPivotTable("P", "Sales", range, 1, 1, Headers);
        Assert.Throws<ArgumentException>(() => sheet.AddPivotTable("p", "Sales", range, 20, 1, Headers));             // Name taken (case-insensitive).
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddPivotTable("Q", "Sales", range, 0, 1, Headers));    // Row out of range.
        // A data field named like a source column makes Excel fail with "field name already exists".
        var ex = Assert.Throws<ArgumentException>(() => pivot.AddDataField("Amount", displayName: "Qty"));
        Assert.Contains("Qty", ex.Message);
        pivot.AddDataField("Amount");
        Assert.Throws<ArgumentException>(() => pivot.AddDataField("Amount"));                                           // Same caption twice.
        pivot.AddDataField("Amount", PivotFunction.Count);                                                               // Different caption: fine.
    }

    [Fact]
    public void PivotWithoutFields_FailsAtSave()
    {
        using var package = new ExcelPackage(new MemoryStream());
        var range = WriteSales(package.Workbook);
        package.Workbook.Worksheets.Add("Pivot").AddPivotTable("Empty", "Sales", range, 1, 1, Headers);
        var ex = Assert.Throws<InvalidOperationException>(() => package.Save());
        Assert.Contains("no fields", ex.Message);
    }

    [Fact]
    public void TemplateSheet_RefusesPivots()
    {
        var template = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Data").Row().WriteCell("x"));
        using var package = ExcelPackage.OpenTemplate(new MemoryStream(template), new MemoryStream());
        var sheet = package.Workbook.Worksheets.Open("Data");
        Assert.Throws<NotSupportedException>(() => sheet.AddPivotTable("P", "Data", "A1:A2", 5, 1, new[] { "x" }));
    }

    /// <summary>
    /// The location element as Excel writes it for a fresh pivot: body ref sized for one item per axis plus
    /// headers and grand totals, header/data offsets per field combination. Excel resizes it on refresh.
    /// </summary>
    [Theory]
    [InlineData("row", 3, 1, "A3:B5", 1, 1, 1)]              // "Row Labels | Sum" header, one item, grand total.
    [InlineData("row,col", 3, 1, "A3:C6", 1, 2, 1)]          // Caption row + header row, one item, grand total; item column + total column.
    [InlineData("row,col,data2", 3, 1, "A3:E7", 1, 3, 1)]    // Values row added; two data columns per item and per total.
    [InlineData("row,data2", 3, 1, "A3:C5", 1, 1, 1)]        // Values side by side without column fields: still one header row.
    [InlineData("page,row", 3, 1, "A5:B7", 1, 1, 1)]         // One filter row + blank row push the body down two rows.
    [InlineData("page,page,row,col", 1, 2, "B4:D7", 1, 2, 1)] // Two filters: body starts at row 4, column B.
    [InlineData("data", 3, 1, "A3:A4", 1, 1, 0)]             // Caption over a single value, no row axis.
    [InlineData("row,nogrand", 3, 1, "A3:B4", 1, 1, 1)]      // No grand total row.
    [InlineData("row,col,nogrand", 3, 1, "A3:B5", 1, 2, 1)]  // No grand total row or column.
    [InlineData("row,row,outline", 3, 1, "A3:C5", 1, 1, 2)]  // Outline form: one column per row field.
    public void Location_IsComputedForEachFieldCombination(string spec, int row, int column, string expectedRef, int headerRow, int dataRow, int dataCol)
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var range = WriteSales(wb);
            var p = wb.Worksheets.Add("Pivot").AddPivotTable("P", "Sales", range, row, column, Headers);
            var rowFields = new Queue<string>(new[] { "Region", "Category", "Product" });
            var dataAdded = false;
            foreach (var part in spec.Split(','))
            {
                switch (part)
                {
                    case "row": p.AddRowField(rowFields.Dequeue()); break;
                    case "col": p.AddColumnField("Qty"); break;
                    case "page": p.AddFilterField(rowFields.Dequeue()); break;
                    case "data2": p.AddDataField("Amount"); p.AddDataField("Amount", PivotFunction.Count); dataAdded = true; break;
                    case "data": p.AddDataField("Amount"); dataAdded = true; break;
                    case "nogrand": p.ShowRowGrandTotals = false; p.ShowColumnGrandTotals = false; break;
                    case "outline": p.Compact = false; break;
                }
            }
            if (!dataAdded && !spec.Contains("data")) p.AddDataField("Amount");
            var loc = p.ComputeLocation();
            Assert.Equal(expectedRef, loc.Ref.ToString());
        });

        using var doc = OpenValidated(bytes);
        var location = doc.WorkbookPart.WorksheetParts.SelectMany(p => p.PivotTableParts).Single().PivotTableDefinition.Location;
        Assert.Equal(expectedRef, location.Reference.Value);
        Assert.Equal((uint)headerRow, location.FirstHeaderRow.Value);
        Assert.Equal((uint)dataRow, location.FirstDataRow.Value);
        Assert.Equal((uint)dataCol, location.FirstDataColumn.Value);
    }

    [Fact]
    public void TwoDataFields_PutValuesAxisOnColumns()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var range = WriteSales(wb);
            var p = wb.Worksheets.Add("Pivot").AddPivotTable("P", "Sales", range, 1, 1, Headers);
            p.ShowRowGrandTotals = false;
            p.ShowRowStripes = true;
            p.AddRowField("Region").AddDataField("Amount").AddDataField("Qty", PivotFunction.Max, numberFormat: "0");
        });
        using var doc = OpenValidated(bytes);
        var def = doc.WorkbookPart.WorksheetParts.SelectMany(p => p.PivotTableParts).Single().PivotTableDefinition;
        Assert.Equal(-2, def.ColumnFields.Elements<X.Field>().Single().Index.Value); // The "Values" pseudo-field.
        Assert.Equal(2, def.DataFields.Count());
        Assert.False(def.RowGrandTotals.Value);
        Assert.Null(def.ColumnGrandTotals); // Default (true) is not written.
        Assert.True(def.PivotTableStyle.ShowRowStripes.Value);
        // Built-in format "0" maps to numFmtId 1 without a custom numFmt.
        Assert.Equal(1u, def.DataFields.Elements<X.DataField>().Last().NumberFormatId.Value);
    }

    [Fact]
    public void LoadResult_ExposesHeadersRangeAndTable()
    {
        ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            s.Row().WriteCell("Title");
            var plain = s.LoadFromCollection(new[] { new { Name = "a", Value = 1 }, new { Name = "b", Value = 2 } }, new LoadOptions { StartColumn = 2 });
            Assert.Equal(new[] { "Name", "Value" }, plain.Headers);
            Assert.Equal("B2:C4", plain.Range);
            Assert.Null(plain.Table);
            var withTable = s.LoadFromCollection(new[] { new { Name = "c", Value = 3 } }, new LoadOptions { TableName = "T" });
            Assert.NotNull(withTable.Table);
            Assert.Equal("T", withTable.Table.Name);
            Assert.Equal("A5:B6", withTable.Range);
            var empty = s.LoadFromCollection(Array.Empty<System.Collections.Generic.IDictionary<string, object?>>());
            Assert.Empty(empty.Headers);
            Assert.Equal("A7", empty.Range);
        });
    }
}
