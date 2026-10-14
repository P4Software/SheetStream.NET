# Analyzers

SheetStream streams each worksheet straight into the file. That is what makes it fast, and it is also why a few
things must happen in a fixed order: layout before rows, rows ascending, one write mode per sheet. The library
throws `InvalidOperationException` when the order is wrong, but an exception at run time, often on a server, is a
poor place to learn a rule.

The SheetStream package therefore ships a Roslyn analyzer. It runs inside the compiler and the IDE and shows the
three rules below as warnings while you type, with a message that says what to change. There is nothing to
install: `dotnet add package SheetStream` is enough, and the analyzer adds nothing to your application's output.

| Id | Title | What it catches |
|---|---|---|
| [SS0001](#ss0001) | Layout setting after rows were written | `Column()`, `Columns()`, `FreezePanes()`, `TabColor`, `ShowGridLines`, `Zoom`, `RightToLeft`, `PageSetup.FitToWidth` / `FitToHeight` used after `Row()` or `LoadFrom*()` |
| [SS0002](#ss0002) | Row() and Cells[] used on the same worksheet | A sheet written through both `Row()` streaming and `Cells[...]` random access |
| [SS0003](#ss0003) | Rows must ascend | `Row(3)` after `Row(5)` on the same sheet |

All three are warnings in the `SheetStream` category.

## How sure the analyzer is

The analyzer only reports what it can prove from the code in one method: the two calls use the same local
variable or parameter, and the earlier one is certain to have run before the later one (same block, or an
enclosing block, with no `if`, loop, `try`, lambda or `??` in between that could skip it). When the order depends on
a branch or a loop, or the sheet lives in a field, it stays quiet rather than guess. A missed warning still ends in
the usual run-time exception; a wrong warning would make you distrust the right ones, so the analyzer errs on the
side of silence.

## SS0001: Layout setting after rows were written

The OpenXML schema stores column widths, frozen panes, tab colour, gridlines, zoom, right-to-left and the
fit-to-page switch *before* the sheet data. SheetStream writes that header with the first row, so changing these
afterwards has nowhere to go.

Before:

```csharp
var sheet = workbook.Worksheets.Add("Orders");
sheet.LoadFromCollection(orders);
sheet.FreezePanes(1);              // SS0001
sheet.Column(1).SetWidth(20);      // SS0001
```

After, either set the layout first:

```csharp
var sheet = workbook.Worksheets.Add("Orders");
sheet.FreezePanes(1);
sheet.Column(1).SetWidth(20);
sheet.LoadFromCollection(orders);
```

or let `AutoFitColumns()` defer the header until the sheet is saved, which also lets the widths follow the data:

```csharp
var sheet = workbook.Worksheets.Add("Orders");
sheet.AutoFitColumns();
sheet.LoadFromCollection(orders);
sheet.FreezePanes(1);              // fine: the header is written at save time
```

Sheets written through `Cells[...]` buffer everything until save, so the rule does not apply to them either.
Settings that live after the data (`AutoFilter`, `MergeCells`, protection, validation, conditional formats, print
margins, orientation) can be set at any time and are never reported.

## SS0002: Row() and Cells[] used on the same worksheet

A sheet is either streamed forward-only through `Row()` and `LoadFrom*()`, or kept in memory and written in order
at save time through `Cells[row, column]`. The first `Cells` access on a streamed sheet, or the first `Row()` on a
buffered one, throws.

Before:

```csharp
var sheet = workbook.Worksheets.Add("Orders");
sheet.LoadFromCollection(orders);
sheet.Cells[1, 1].Value = "Total";   // SS0002
```

After, pick one mode per sheet. For a streamed sheet, write the extra cell as a row:

```csharp
var sheet = workbook.Worksheets.Add("Orders");
var result = sheet.LoadFromCollection(orders);
sheet.Row(result.LastRow + 1).WriteCell("Total").WriteFormula($"SUM(B2:B{result.LastRow})");
```

or, when the layout really is random access and the sheet is small, use `Cells` for all of it:

```csharp
var sheet = workbook.Worksheets.Add("Summary");
sheet.Cells[1, 1].Value = "Total";
sheet.Cells[1, 2].Formula = "SUM(Orders!B:B)";
```

## SS0003: Rows must ascend

Rows stream in ascending order; a row index that has already been passed cannot be written. The analyzer checks
integer literals only (`Row(5)` then `Row(3)`), since those are the cases it can be certain about.

Before:

```csharp
var sheet = workbook.Worksheets.Add("Report");
sheet.Row(5).WriteCell("Body");
sheet.Row(3).WriteCell("Title");   // SS0003
```

After, write in order:

```csharp
var sheet = workbook.Worksheets.Add("Report");
sheet.Row(3).WriteCell("Title");
sheet.Row(5).WriteCell("Body");
```

or switch that sheet to random access when the order is genuinely data-driven:

```csharp
var sheet = workbook.Worksheets.Add("Report");
sheet.Cells[5, 1].Value = "Body";
sheet.Cells[3, 1].Value = "Title";
```

## Suppressing a warning

The analyzer is meant to be right every time it speaks, so a warning usually means the code will throw. If you do
need to silence one, the standard mechanisms apply.

For one statement:

```csharp
#pragma warning disable SS0001 // header is deferred by an AutoFitColumns() call in another method
sheet.FreezePanes(1);
#pragma warning restore SS0001
```

For a member or type:

```csharp
[System.Diagnostics.CodeAnalysis.SuppressMessage("SheetStream", "SS0002:Row() and Cells[] used on the same worksheet", Justification = "The two calls never run in the same export")]
```

For a whole project, in `.editorconfig`:

```ini
[*.cs]
dotnet_diagnostic.SS0001.severity = none      # or suggestion, warning, error
dotnet_diagnostic.SS0002.severity = error
```

To turn the rules into build errors, set the severity to `error` the same way, or add `SS0001;SS0002;SS0003` to
`<WarningsAsErrors>` in the project file.

## Not what you expected?

If a warning fires on code that runs fine, or does not fire on code that throws, please open an issue at
<https://github.com/P4Software/SheetStream/issues> with the snippet. False positives are treated as bugs.
