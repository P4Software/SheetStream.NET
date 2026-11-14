---
title: "Fill Excel templates from code"
description: "Design the report in Excel, then let SheetStream.NET pour the data in. Everything you do not touch is copied from the template unchanged."
---

# Filling templates

Design the report in Excel (logo, colours, a styled header row, totals, charts, print settings) and let SheetStream pour the data in. Everything you do not touch is copied from the template unchanged.

```csharp
using SheetStream;

using var package = ExcelPackage.OpenTemplate("invoice-template.xlsx", "invoice-1042.xlsx");
var sheet = package.Workbook.Worksheets.Open("Lines");

// Reuse the formatting the designer gave the first data row.
var money = sheet.TemplateStyle("D5");

var row = 5;
foreach (var line in invoice.Lines)
    sheet.Row(row++).WriteCell(line.Sku).WriteCell(line.Description).WriteCell(line.Qty).WriteCell(line.Amount, money);

package.Save();
```

## How rows merge with the template

- Template rows above your first row are kept: titles, logos, the header row.
- Each row you write replaces the template row with the same number. A formatted but empty row (the usual "first data row") is simply replaced.
- Template rows below your last row are kept, so leave room for footers. `sheet.ReplacedTemplateRows` tells you how many template rows that held values were overwritten, so you can detect a footer that the data ran into.
- Rows stream, so a template can receive a million rows with flat memory.

## What stretches automatically

When a template range starts in the header row or the first data row and ends before your last row, SheetStream extends it to cover the data:

| In the template | After filling 100 rows from row 5 |
|---|---|
| Table `A4:D5` | `A4:D104` |
| Table filter `A4:D5` | `A4:D104` |
| Data validation on `C5` | `C5:C104` |
| Conditional format on `D5` | `D5:D104` |

Formulas such as `=SUM(Lines[Amount])` follow the table, and the workbook is set to recalculate when opened, so totals are correct immediately. Tables with a totals row are left as they are, because the data would run into it.

## What comes from the template

Column widths, frozen panes, sheet protection, merged cells, validation and formatting rules, charts and pictures belong to the template. Changing them on a template sheet throws `NotSupportedException` with a message pointing you back to Excel. New sheets cannot be added either; add them to the template.

New cell styles are fine: `package.Workbook.Styles.Create(...)` appends to the template's style table without disturbing the existing styles.

The filled copy appears at the output path when `Save()` completes. Until then, and if the package is disposed without saving, nothing is written there.

## Templates from people you do not know

`OpenTemplate` applies the same limits as `ExcelReader` to the template it loads (see [Reading workbooks](reading-workbooks.md#files-from-people-you-do-not-know)), so a workbook a user uploaded can be used as a template without it being able to exhaust memory. A template that goes over a limit is refused with an `InvalidDataException` naming the limit. The limits are on `PackageOptions.SourceLimits`:

```csharp
var options = new PackageOptions { SourceLimits = new ReaderOptions { MaxCellTextLength = 32_767, MaxLoadedPartBytes = 32 * 1024 * 1024 } };
using var package = ExcelPackage.OpenTemplate(uploadedTemplate, output, leaveOpen: false, options);
```

Two things to know. The packed file itself is held in memory while the package is open, so cap the upload size in your web framework as you would for any upload. And comments and processing instructions inside the XML of a sheet you fill are not carried over to the copy; Excel writes neither, so this only affects files hand-edited as XML.

Templates are parsed with DTD processing and external entities disabled, and the template file is never modified: `OpenTemplate` refuses to write to the same path it reads.
