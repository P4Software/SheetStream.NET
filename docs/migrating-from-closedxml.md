# Migrating from ClosedXML

ClosedXML is a good library: MIT licensed, a pleasant API, a calculation engine, and it can change any part of an existing workbook. It keeps the whole workbook in memory to do that, which is the one thing that hurts when an export grows. If your exports are large, or run many at once on a web server, this page shows what moving them to SheetStream.NET looks like. If you edit large workbooks or need formulas calculated in .NET, stay with ClosedXML for that part; the two libraries live happily in one application.

## What you gain and what you give up

| | ClosedXML 0.105 | SheetStream.NET |
|---|---|---|
| Model | The workbook is an object tree in memory; saved in one go at the end | Rows are written to the file as you produce them |
| 100,000 rows of 10 columns, written | 3,103 ms, 1,104 MB allocated, 231 MB peak | 1,013 ms, 27 MB allocated, 30 MB peak |
| The same file, read | 3,034 ms, 1,598 MB allocated | 533 ms, 119 MB allocated |
| Cells in any order | Yes | Rows ascend. `Cells[row, column]` gives random access for one small sheet at a time |
| Change a workbook that exists | Yes, any part | Cell values of small sheets (`ExcelPackage.Edit`), or fill a template |
| Formulas calculated in .NET | Yes | No. Simple aggregates can be cached; Excel calculates the rest on open |
| Dependencies | Open XML SDK and others | None |

The figures are from this repository's benchmarks (BenchmarkDotNet, .NET 8, one developer workstation, 2026-10-13; see [benchmarks/README.md](https://github.com/P4Software/SheetStream.NET/blob/main/benchmarks/README.md)). Both libraries wrote the same sheet with the same formats. Run them on your own hardware before relying on them.

## Before: ClosedXML

```csharp
using ClosedXML.Excel;

public void WriteOrders(Stream output, Order[] orders)
{
    using var wb = new XLWorkbook();
    var ws = wb.Worksheets.Add("Orders");

    string[] headers = { "Id", "Customer", "Quantity", "Amount", "Date" };
    for (int c = 0; c < headers.Length; c++) ws.Cell(1, c + 1).Value = headers[c];
    ws.Range(1, 1, 1, headers.Length).Style.Font.Bold = true;

    for (int i = 0; i < orders.Length; i++)
    {
        var o = orders[i];
        int row = i + 2;
        ws.Cell(row, 1).Value = o.Id;
        ws.Cell(row, 2).Value = o.Customer;
        ws.Cell(row, 3).Value = o.Quantity;
        ws.Cell(row, 4).Value = o.Amount;
        ws.Cell(row, 5).Value = o.Date;
    }

    int last = orders.Length + 1;
    ws.Range(2, 4, last, 4).Style.NumberFormat.Format = "#,##0.00";
    ws.Range(2, 5, last, 5).Style.NumberFormat.Format = "yyyy-mm-dd";
    ws.SheetView.FreezeRows(1);
    ws.Columns().AdjustToContents();

    wb.SaveAs(output);
}
```

## After: SheetStream.NET

```csharp
using SheetStream;

public void WriteOrders(Stream output, Order[] orders)
{
    using var package = new ExcelPackage(output, leaveOpen: true);
    var workbook = package.Workbook;

    // Styles are made once and passed with the cell. Nothing is restyled afterwards.
    var bold = workbook.Styles.Create(o => o.Font.Bold = true);
    var money = workbook.Styles.Create(o => o.NumberFormat = "#,##0.00");
    var date = workbook.Styles.Create(o => o.NumberFormat = "yyyy-mm-dd");

    var sheet = workbook.Worksheets.Add("Orders");
    sheet.AutoFitColumns();          // Replaces Columns().AdjustToContents() at the end
    sheet.FreezePanes(rows: 1);      // Replaces SheetView.FreezeRows(1)

    sheet.Row(1).WriteCells(bold, new[] { "Id", "Customer", "Quantity", "Amount", "Date" });
    foreach (var o in orders)
        sheet.Row().WriteCell(o.Id).WriteCell(o.Customer).WriteCell(o.Quantity).WriteCell(o.Amount, money).WriteCell(o.Date, date);

    package.Save();                  // Nothing is written until Save
}
```

What changed:

1. **The format travels with the cell.** ClosedXML writes the values first and styles ranges afterwards, which it can do because every cell is still in memory. Here a cell is gone once written, so its style is given when it is written.
2. **Layout is declared before the rows.** Column widths and frozen panes are stored ahead of the data in the file. `AutoFitColumns()` is called first and measures the rows as they go past.
3. **Rows go top to bottom.** `sheet.Row()` is the next row; `sheet.Row(n)` jumps forward. To fill cells in any order on a small sheet, use `sheet.Cells[row, column].Value`, which buffers that one sheet.
4. **`Save()` is required.** As in ClosedXML, disposing without saving writes nothing. An analyzer rule (SS0004) reports a package that is never saved.

When the data is a list of objects, the whole loop is one call, with headers taken from the property names (or from `[DisplayName]` / `[SheetColumn]`):

```csharp
sheet.LoadFromCollection(orders, new LoadOptions { TableName = "Orders" });   // ClosedXML: ws.Cell(1, 1).InsertTable(orders)
```

## Reading

```csharp
// ClosedXML: the whole workbook is loaded, then walked
using var wb = new XLWorkbook(stream);
foreach (var row in wb.Worksheet("Orders").RowsUsed())
{
    var id = row.Cell(1).GetDouble();
    var customer = row.Cell(2).GetString();
}

// SheetStream.NET: rows are parsed as you enumerate them
using var reader = ExcelReader.Open(stream);
foreach (var row in reader.ReadRows("Orders"))
{
    var id = row.GetDouble(0);            // Columns are zero-based here
    var customer = row.GetString(1);
}
```

`reader.Read<Order>("Orders")` maps rows to objects by header, and `reader.ReadDataTable("Orders")` gives a `DataTable`. See [Reading workbooks](reading-workbooks.md), including the limits that make it safe to read a file a user uploaded.

## Quick reference

| ClosedXML | SheetStream.NET |
|---|---|
| `new XLWorkbook()` … `wb.SaveAs(stream)` | `new ExcelPackage(stream)` … `package.Save()`, or `ExcelPackage.CreateBytes(wb => …)` for a `byte[]` |
| `wb.SaveAs("file.xlsx")` | `ExcelPackage.Create("file.xlsx")` … `package.Save()` |
| `wb.Worksheets.Add("Name")` | `package.Workbook.Worksheets.Add("Name")` |
| `ws.Cell(r, c).Value = v` | `sheet.Row(r).WriteCell(v)` in order, or `sheet.Cells[r, c].Value = v` |
| `ws.Cell(r, c).FormulaA1 = "SUM(A1:A9)"` | `row.WriteFormula("SUM(A1:A9)")` |
| `cell.Style.Font.Bold = true` | `var bold = Styles.Create(o => o.Font.Bold = true);` then `WriteCell(v, bold)` |
| `cell.Style.Fill.BackgroundColor = XLColor.Cyan` | `Styles.Create(o => o.Fill.BackgroundColor = "#00FFFF")` |
| `cell.Style.Font.FontColor = XLColor.DarkBlue` | `Styles.Create(o => o.Font.Color = "#00008B")` |
| `cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center` | `Styles.Create(o => o.HorizontalAlignment = HorizontalAlignment.Center)` |
| `cell.Style.NumberFormat.Format = "#,##0.00"` | `Styles.Create(o => o.NumberFormat = "#,##0.00")` |
| `ws.Cell(1, 1).InsertTable(list)` | `sheet.LoadFromCollection(list, new LoadOptions { TableName = "Name" })` |
| `ws.Cell(1, 1).InsertData(list)` | `sheet.LoadFromCollection(list)` |
| `ws.Columns().AdjustToContents()` at the end | `sheet.AutoFitColumns()` **before the first row** |
| `ws.Column(2).Width = 20` | `sheet.Column(2).SetWidth(20)` before the first row |
| `ws.SheetView.FreezeRows(1)` | `sheet.FreezePanes(rows: 1)` before the first row |
| `ws.Range("A1:D1").Merge()` | `sheet.MergeCells("A1:D1")` |
| `ws.RangeUsed().SetAutoFilter()` | `sheet.AutoFilter(1, 1, 4)` |
| `ws.Cell("A5").DataValidation.List(range)` | `sheet.AddDataValidation("A5").List("Yes", "No")` |
| `cell.Hyperlink = new XLHyperlink(url, tooltip)` | `row.WriteHyperlink(text, url)` |
| `ws.Protect("password")` | `sheet.Protect("password")` |
| `new XLWorkbook(file)`, change cells, `Save()` | `ExcelPackage.Edit(path)` + `Worksheets.Edit("Sheet")` + `Cells[…]` + `Save()`, for small sheets |
| `row.Cell(n).GetString()` / `GetDouble()` / `GetDateTime()` | `row.GetString(n - 1)` / `GetDouble(n - 1)` / `GetDateTime(n - 1)` |

## What does not translate

Stay with ClosedXML for: calculating formulas in .NET, inserting or deleting rows and columns in an existing sheet, changing the layout (widths, merges, formatting rules) of an existing workbook, copying worksheets between workbooks, and editing a sheet too large to hold in memory twice. SheetStream.NET edits cell values of small sheets and fills templates; everything else about an existing workbook it copies as it is.
