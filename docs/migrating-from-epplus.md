# Migrating from EPPlus to SheetStream

This guide converts a real exporter, taken from a production SaaS app, from EPPlus to SheetStream. Most EPPlus export code maps line for line; the one thing to unlearn is setting styles cell by cell.

## The mental model in one minute

| EPPlus | SheetStream |
|---|---|
| Builds the whole workbook in memory, then saves it | Streams rows into the file as you write them |
| `ws.Cells[r, c]` in any order | `ws.Row(r).WriteCell(...)` left to right, rows ascending (or `ws.Cells[r, c]` for small, irregular sheets) |
| `cell.Style.Font.Bold = true` on every cell | Create a style once with `Styles.Create(...)`, pass it to each cell |
| Column widths any time | Column widths before the first row, or call `AutoFitColumns()` first |
| Commercial licence for business use (EPPlus 5+) | MIT |

## Before: EPPlus

```csharp
public async Task<byte[]> ExportToExcelAsync(QueryResultDto result)
{
    using var package = new ExcelPackage();
    var worksheet = package.Workbook.Worksheets.Add("Query Results");

    for (int col = 0; col < result.Columns.Count; col++)
    {
        var cell = worksheet.Cells[1, col + 1];
        cell.Value = result.Columns[col];
        cell.Style.Font.Bold = true;
        cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189));
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    }

    for (int row = 0; row < result.Rows.Count; row++)
    {
        for (int col = 0; col < result.Columns.Count; col++)
        {
            var columnName = result.Columns[col];
            var cellValue = result.Rows[row].ContainsKey(columnName) ? result.Rows[row][columnName] : null;
            var cell = worksheet.Cells[row + 2, col + 1];
            if (cellValue != null)
            {
                if (cellValue is DateTime dateValue) { cell.Value = dateValue; cell.Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss"; }
                else if (cellValue is decimal or double or float) { cell.Value = cellValue; cell.Style.Numberformat.Format = "#,##0.00"; }
                else if (cellValue is int or long) { cell.Value = cellValue; cell.Style.Numberformat.Format = "#,##0"; }
                else cell.Value = cellValue.ToString();
            }
            if (row % 2 == 1)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 242, 242));
            }
        }
    }

    if (result.Rows.Count > 0)
        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
    worksheet.View.FreezePanes(2, 1);
    return package.GetAsByteArray();
}
```

## After: SheetStream

```csharp
using SheetStream;
using SheetStream.Styles;

public Task<byte[]> ExportToExcelAsync(QueryResultDto result)
{
    var bytes = ExcelPackage.CreateBytes(wb =>
    {
        // Styles are created once and reused; SheetStream de-duplicates them in the file.
        var header = wb.Styles.Create(o =>
        {
            o.Font.Bold = true;
            o.Font.Color = Colors.White;
            o.Fill.BackgroundColor = Colors.Rgb(79, 129, 189);
            o.HorizontalAlignment = HorizontalAlignment.Center;
        });
        var stripe = wb.Styles.Create(o => o.Fill.BackgroundColor = Colors.Rgb(242, 242, 242));
        Func<string, bool, CellStyle> fmt = (code, striped) =>
            wb.Styles.Create(o => { o.NumberFormat = code; if (striped) o.Fill.BackgroundColor = Colors.Rgb(242, 242, 242); });
        var dateTime = new[] { fmt("yyyy-mm-dd hh:mm:ss", false), fmt("yyyy-mm-dd hh:mm:ss", true) };
        var money = new[] { fmt("#,##0.00", false), fmt("#,##0.00", true) };
        var whole = new[] { fmt("#,##0", false), fmt("#,##0", true) };

        var sheet = wb.Worksheets.Add("Query Results");
        sheet.AutoFitColumns();              // Replaces AutoFitColumns() at the end
        sheet.FreezePanes(rows: 1);          // EPPlus FreezePanes(2, 1) = "freeze row 1"

        var head = sheet.Row(1);
        foreach (var column in result.Columns) head.WriteCell(column, header);

        for (var r = 0; r < result.Rows.Count; r++)
        {
            var striped = r % 2 == 1 ? 1 : 0;
            var row = sheet.Row();
            foreach (var column in result.Columns)
            {
                result.Rows[r].TryGetValue(column, out var value);
                switch (value)
                {
                    case DateTime d: row.WriteCell(d, dateTime[striped]); break;
                    case decimal or double or float: row.WriteCell(value, money[striped]); break;
                    case int or long: row.WriteCell(value, whole[striped]); break;
                    case null: row.WriteEmpty(striped == 1 ? stripe : null); break;
                    default: row.WriteCell(value.ToString(), striped == 1 ? stripe : null); break;
                }
            }
        }
    });
    return Task.FromResult(bytes);
}
```

What changed:

1. **Styles are values, not per-cell mutations.** EPPlus sets seven style properties per header cell; SheetStream creates `header` once. For a 100,000-row export that is the difference between millions of style operations and a handful.
2. **Rows stream.** Nothing is held in memory, so the same code handles 10 rows or 1,000,000.
3. **AutoFit and freeze are declared up front.** SheetStream writes column widths before the data, so `AutoFitColumns()` is called first and measures the rows as they stream past.
4. **No `System.Drawing`.** Colours are `"#RRGGBB"` strings or `Colors.Rgb(r, g, b)`, so the code runs on Linux containers without libgdiplus.

## Returning it from ASP.NET Core

With the `SheetStream.AspNetCore` package, the controller no longer needs the byte array at all:

```csharp
[HttpPost("export/excel")]
public IResult ExportExcel(QueryResultDto result) =>
    ExcelResults.Excel("query-results.xlsx", wb => BuildQuerySheet(wb, result));
```

The file is built into a temporary file and streamed to the response, so memory stays flat even for very large results.

## Quick reference

| EPPlus | SheetStream |
|---|---|
| `new ExcelPackage()` … `GetAsByteArray()` | `ExcelPackage.CreateBytes(wb => …)` |
| `new ExcelPackage(file)` … `SaveAs(file)` | `ExcelPackage.Create(path)` … `Save()` |
| `new ExcelPackage(templateFile)` to fill a template | `ExcelPackage.OpenTemplate(template, output)` + `Worksheets.Open("Sheet")` |
| `ws.Cells["A1"].LoadFromCollection(list, true)` | `ws.LoadFromCollection(list)` |
| `ws.Cells["A1"].LoadFromDataTable(table, true)` | `ws.LoadFromDataTable(table)` |
| `ws.Cells[r, c].Value = v` | `ws.Row(r).WriteCell(v)` or `ws.Cells[r, c].Value = v` |
| `ws.Cells[r, c].Formula = "SUM(A1:A9)"` | `row.WriteFormula("SUM(A1:A9)")` |
| `cell.Style.Numberformat.Format = "0.00%"` | `Styles.Create(o => o.NumberFormat = "0.00%")` |
| `cell.Style.Border.BorderAround(ExcelBorderStyle.Thin)` | `Styles.Create(o => o.Border.Style = BorderStyle.Thin)` |
| `ws.Column(2).Width = 20` | `ws.Column(2).SetWidth(20)` |
| `ws.Cells["A1:D1"].Merge = true` | `ws.MergeCells("A1:D1")` |
| `ws.Cells["A1:D1"].AutoFilter = true` | `ws.AutoFilter(1, 1, 4)` |
| `ws.Tables.Add(ws.Cells["A1:D20"], "Sales")` | `ws.AddTable("Sales", "A1:D20")` |
| `ws.ConditionalFormatting.AddLessThan(range)` | `ws.AddConditionalFormatting("D2:D20").CellValue(ConditionalOperator.LessThan, "0", style)` |
| `ws.DataValidations.AddListValidation("C2:C20")` | `ws.AddDataValidation("C2:C20").List("Open", "Closed")` |
| `ws.Drawings.AddChart("c", eChartType.ColumnClustered)` | `ws.AddChart(ChartType.Column, "F2:M20")` |
| `ws.Drawings.AddPicture("logo", file)` | `ws.AddPicture(path, row, column)` |
| `ws.Cells["A1"].Hyperlink = new Uri(url)` | `row.WriteHyperlink(text, url)` |
| `ws.Comments.Add(cell, text, author)` | `ws.AddComment(row, column, text, author)` |
| `ws.Protection.IsProtected = true` | `ws.Protect(password)` |
| `ws.PrinterSettings.Orientation = eOrientation.Landscape` | `ws.PageSetup.Orientation = PageOrientation.Landscape` |
| Reading: `ws.Cells[r, c].Value` | `ExcelReader.Open(file).ReadRows("Sheet")` or `.Read<T>("Sheet")` |

## What does not translate

These EPPlus features have no SheetStream equivalent today: formula calculation (`Calculate()`; SheetStream only writes cached results for simple aggregates, see `CacheFormulaResults`), pivot tables with calculated fields, grouping or slicers (plain pivots are supported and refreshed on open), sparklines, editing large arbitrary workbooks in place (small workbooks can be edited with `ExcelPackage.Edit`; large ones are filled as templates), VBA, encryption and digital signatures. If an exporter depends on one of them, keep that exporter on EPPlus 4.5.3.3 (LGPL) or another library.

## After removing the EPPlus package

Removing EPPlus can break code that never called it. EPPlus 4.5.3.3 depends on `System.Text.Encoding.CodePages` and registers the code-page encoding provider when it loads. Other libraries in the same application may rely on that registration without declaring it; **ExcelDataReader** is the usual case, because legacy `.xls` files need code-page encodings. With EPPlus gone, those imports fail, sometimes silently with empty text.

Register the provider yourself at startup, before the first import:

```csharp
// Program.cs, before building the app. On .NET Framework or netstandard targets add the
// System.Text.Encoding.CodePages package; on .NET 8 it is part of the shared framework.
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
```

Then run one `.xls` import in a test to prove it. This was found while migrating a warehouse system whose exports moved to SheetStream but whose imports stayed on ExcelDataReader. For `.xlsx` imports, `ExcelReader` needs no provider.
