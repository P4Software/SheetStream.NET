# Editing existing workbooks

Open a small workbook, read or change some cells, save. This is the "open, change a cell, save" flow EPPlus users expect, built for budgets, settings sheets, price lists and other files a person maintains by hand.

```csharp
using SheetStream;

using var package = ExcelPackage.Edit("budget.xlsx");           // In place
var sheet = package.Workbook.Worksheets.Edit("Summary");

var current = sheet.Cells["B4"].Value;                            // 1250.0, a DateTime, "text", true...
sheet.Cells["B4"].Value = 1400;                                   // Keeps the cell's number format
sheet.Cells["C4"].Formula = "B4*1.21";

var next = sheet.Cells.Dimension.LastRow + 1;                     // Append below the data
sheet.Cells[next, 1].Value = "Added by code";
sheet.Cells[next, 2].Value = DateTime.Today;

package.Save();                                                   // Replaces budget.xlsx atomically
```

Three ways to open:

| Call | Writes to |
|---|---|
| `ExcelPackage.Edit("in.xlsx", "out.xlsx")` | A new file; the input is never touched |
| `ExcelPackage.Edit("book.xlsx")` | The same file. The copy is built next to it as a temporary file and swapped in only when the save succeeds; if saving throws, the original is untouched and the temporary file is removed |
| `ExcelPackage.Edit(inputStream, outputStream)` | A stream, for uploads and downloads |

## What happens on save

The whole workbook is re-written:

- A sheet opened with `Worksheets.Edit` is written from its cells. Cells you did not touch are written exactly as they were (formula text, cached result, rich text, metadata attributes). Row heights, hidden rows and outline levels are kept.
- Every other part is copied byte for byte: sheets you did not open, charts, pictures, tables, comments, defined names, print settings, document properties.
- Shared strings and styles are merged: the file's own entries keep their indexes, anything new is appended.
- The calculation chain is dropped and the workbook is set to recalculate on open, so formulas that read the cells you changed show correct values the moment Excel opens the file. SheetStream has no calculation engine; `Value` on a formula cell returns the result cached in the file.

## Reading and changing cells

| | |
|---|---|
| `sheet.Cells["B4"].Value` | The stored value: `string`, `double`, `bool`, `DateTime` (numbers with a date format), `TimeSpan` (elapsed-time formats), `ExcelErrorValue`, or `null`. A formula cell returns its cached result |
| `sheet.Cells["B4"].Formula` | The formula text without `=`, or `null` |
| `sheet.Cells["B4"].Style` | A `CellStyle` referencing the format the cell has in the file. Assign it to another cell to copy the look, or `sheet.TemplateStyle("B4")` for a cell you did not load |
| `sheet.Cells["B4"].Value = x` | Replaces the content, keeps the style. `null` empties the cell but keeps its format |
| `sheet.Cells["B4"].Clear()` | Removes the cell: value, formula and style |
| `sheet.Cells.Existing` | Every cell that exists, in row then column order |
| `sheet.Cells.Dimension` | `(LastRow, LastColumn)` of the data, `(0, 0)` on an empty sheet |

New styles come from `package.Workbook.Styles.Create(...)` as usual; they are appended after the file's own styles.

Rows appended below the original last row extend a table whose last row was the data's last row (and its filter), the same stretching template mode does. Nothing else moves.

## Limits

- **Size.** Editing holds the sheet in memory. A sheet with more than `PackageOptions.MaxEditableCells` cells (default 200,000) is refused with an `InvalidOperationException`; the rest of the package is unaffected, so you can still edit smaller sheets of the same file. For big sheets, fill a copy with [templates](templates.md) or write it from scratch with `Row()` streaming.
- **Layout comes from the file.** Column widths, frozen panes, merges, validation, conditional formatting, hyperlinks, comments, charts and protection cannot be added to an edited sheet; change them in Excel. New sheets cannot be added either.
- **Shared formulas.** Excel stores a formula filled down a column once, in a master cell the others point at. Changing or clearing that master cell throws `NotSupportedException` naming the cell, because the dependent cells would be left without a formula. Changing a dependent, or only the master's style, is fine. A dependent reports `Formula == null` (its text lives in the master).
- **Dates in 1904-system workbooks** are read and written with the shifted epoch; everything else about the date system is left as it was.
- **Defined names and workbook protection** cannot be added from code to an edited workbook (they are copied from the file).

## Edit, template or stream?

| You want to | Use | Memory |
|---|---|---|
| Change a few cells in a file someone maintains in Excel | `ExcelPackage.Edit` + `Worksheets.Edit` | The edited sheet |
| Pour thousands to millions of rows into a designed layout | `ExcelPackage.OpenTemplate` + `Worksheets.Open` | Flat |
| Produce a new file from data | `ExcelPackage.Create` + `Worksheets.Add` | Flat |

Edit and template modes share one pipeline: both copy the untouched parts of the source and rewrite the sheets you open. The difference is only in how the opened sheet's rows are produced: streamed around the original rows (template) or loaded and written back from memory (edit). You can mix them in one package, editing one sheet and streaming into another.
