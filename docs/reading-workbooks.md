---
title: "Read .xlsx files row by row in .NET"
description: "ExcelReader reads .xlsx files row by row in constant memory, with limits that make uploaded files safe to read."
---

# Reading workbooks

`ExcelReader` reads `.xlsx` files row by row. A sheet of any size reads in constant memory, apart from the shared string table, which is loaded once.

```csharp
using SheetStream.Reading;

using var reader = ExcelReader.Open("import.xlsx");

foreach (var name in reader.SheetNames)                  // tab order, hidden sheets included
    Console.WriteLine(name);

foreach (var row in reader.ReadRows("Products"))         // or ReadRows(0) by position
    Console.WriteLine($"{row.Index}: {row.GetString(0)} = {row.GetDouble(2)}");
```

Old binary `.xls` files are not supported; opening one fails with a message that says to save it as `.xlsx` first.

## What a row gives you

`ExcelRow.Index` is the 1-based row number in the sheet. `Values` is zero-based like a `DataRow` (column A is index 0) and runs up to the last cell that has a value; cells in between that are empty are `null`.

| Cell in the file | Value you get |
|---|---|
| Text, shared or inline, rich text flattened | `string` |
| Number | `double` |
| Number with a date or time format | `DateTime` (the 1904 date system is handled) |
| Number with an elapsed-time format such as `[h]:mm:ss` | `TimeSpan` |
| Boolean | `bool` |
| Error such as `#N/A` | `ExcelErrorValue` |
| Formula | Its cached result, typed as above; `null` when the file carries none |

`row[2]` returns the raw value, or `null` beyond the last cell. `GetString`, `GetDouble` and `GetDateTime` convert for you and throw a `FormatException` that names the row and column when the value cannot be converted.

## Rows into objects, or a DataTable

```csharp
List<Product> products = reader.Read<Product>("Products").ToList();
DataTable table = reader.ReadDataTable("Products");                 // first row is the header
```

`Read<T>` treats the first row as the header and matches each header, ignoring case, to a writable property or field by `[SheetColumn]` name, `[DisplayName]` or member name. Columns that match nothing are ignored, blank rows are skipped, and a value that cannot be converted throws a `FormatException` naming the row, the column and the header. For trimmed or Native AOT applications pass a generated mapper, `reader.Read("Products", Product.SheetMapper)`; see [Trimming, Native AOT and the source generator](aot-and-source-generator.md).

`ReadDataTable` names columns from the header row (a blank header gets its column letter, a duplicate gets a numeric suffix) and types every column as `object` so mixed columns survive. Pass `firstRowIsHeader: false` to get columns named A, B, C. A table holds the whole sheet, so its size is limited by `ReaderOptions.MaxLoadedPartBytes` (256 MB by default, which is about 400,000 rows of ten numeric columns); for a sheet larger than that, read the rows or raise the limit.

## Merged cells and hyperlinks

```csharp
foreach (var merge in reader.ReadMergedRanges("Report"))
    Console.WriteLine($"{merge.Range} covers B1: {merge.Contains(1, 2)}");
foreach (var link in reader.ReadHyperlinks("Report"))
    Console.WriteLine($"{link.Address}: {link.Url ?? link.Location}");
```

Both skip the rows without parsing them, so they are quick even on a very large sheet.

## Async (.NET 8 and later)

```csharp
using var reader = await ExcelReader.OpenAsync(Request.Body, cancellationToken: ct);
await foreach (var row in reader.ReadRowsAsync("Products", ct)) { /* ... */ }
await foreach (var product in reader.ReadAsync<Product>("Products", ct)) db.Products.Add(product);
var table = await reader.ReadDataTableAsync("Products", firstRowIsHeader: true, ct);
```

The sheet is read without blocking a thread, and cancellation is checked between rows.

## Streams that cannot seek

The directory of a zip file is at its end, so a stream that cannot seek (an HTTP request body, for example) is copied before reading. Up to 64 MB stays in memory; beyond that the copy goes to a temporary file that is deleted when the reader is disposed. This is the same for `Open` and `OpenAsync`.

## Files from people you do not know

An `.xlsx` is a zip archive, so a file of a few hundred kilobytes can unpack into gigabytes. A workbook built that way can be 192 KB on the wire and hold a single cell of 200 million characters. This library refuses such files instead of running out of memory: it stops with an `InvalidDataException` whose message names the limit that was exceeded. The limits apply wherever a workbook is opened: `ExcelReader`, and `ExcelPackage.OpenTemplate` and `ExcelPackage.Edit` through `PackageOptions.SourceLimits`.

| Limit (`ReaderOptions`) | Default | What it bounds |
|---|---|---|
| `MaxCellTextLength` | 1,048,576 characters | The text of one cell or one shared string. Excel's own limit is 32,767 |
| `MaxRowTextLength` | 16,777,216 characters | The text of one row across all its cells, not counting shared strings, which are stored once however many cells use them |
| `MaxLoadedPartBytes` | 256 MB | The memory one workbook may make the library keep: shared strings, styles, the workbook and its relationships, and any sheet loaded for editing. No such part may unpack to more than this, and what they all take in memory once parsed is added up and may not exceed it either |
| `MaxStreamedPartBytes` | 4 GB | The unpacked size of a part that is streamed and not kept: a sheet read row by row, a template sheet being filled, a part copied to the output. This one bounds time, not memory |

Some limits are fixed, because no real workbook comes near them:

- XML nested more than 256 levels deep.
- A single tag (an element with all its attributes) longer than `MaxCellTextLength` characters, or 65,536 if that is larger.
- More than 16,384 different element and attribute names in one part. A real part uses a few hundred.
- More than 16,384 cells in a row, Excel's last column.
- A row number outside 1 to 1,048,576 in a sheet that is edited or filled.

With the defaults, the most a hostile workbook can make the library hold is about `MaxLoadedPartBytes`, plus one row and the XML parser's buffers. In our tests, files built to exhaust memory, each about 100 MB unpacked, peaked at 65 MB of process memory or less through every way of opening them; the exception is a workbook whose shared strings are genuinely that large, which uses what the limit allows.

The defaults sit far above anything Excel writes, so that unusual but genuine files still open. An endpoint that takes uploads from the public and expects ordinary data should be much stricter:

```csharp
var options = new ReaderOptions
{
    MaxCellTextLength = 32_767,                 // Excel's own cell limit
    MaxLoadedPartBytes = 32 * 1024 * 1024,      // at most 32 MB held for one workbook
    MaxStreamedPartBytes = 256 * 1024 * 1024,   // at most 256 MB of sheet to process
};
using var reader = ExcelReader.Open(upload, leaveOpen: false, options);
```

If a genuine file is refused, the message tells you which property to raise.

### What a damaged file can throw

Reading a file that is damaged, truncated, over a limit or not a workbook at all ends in one of four exception types: `InvalidDataException`, `XmlException`, `FormatException` or `OverflowException`. Catch those around an import and report "this file could not be read" to the user. The same four cover `OpenTemplate` and `Edit`. Asking for a sheet by a name the workbook does not have throws `ArgumentException`, as it does for any workbook.

This is tested, not assumed. Every release runs about sixty kinds of hostile file through every way of opening a workbook, a second time in a process whose memory is capped, and feeds tens of thousands of randomly damaged files to the reader and to the template and edit code.

### What the limits do not cover

- **The file as uploaded.** The limits bound what unpacking can add. The packed file itself is your web framework's to cap, with its request size limit. `OpenTemplate` and `Edit` keep the packed file in memory while the package is open; `ExcelReader` reads a seekable stream in place and copies one that cannot seek (see above).
- **What you choose to collect.** `reader.ReadRows(...).ToList()` and `Read<T>(...).ToList()` hold every row, however carefully each one was read. Process rows as they come when the file is not yours. `ReadDataTable` is the exception: the table it builds is counted against `MaxLoadedPartBytes` and a sheet that would take more is refused.
- **The work of streaming.** A sheet within `MaxStreamedPartBytes` is read for as long as it is. A row with one cell in the last column still yields sixteen thousand values, so a sheet of such rows costs far more time than its size on disk suggests. Lower `MaxStreamedPartBytes`, or pass a cancellation token to `ReadRowsAsync`, to bound the time one file may take.
- **Text encodings.** Parts are read as UTF-8 or UTF-16, which is what the file format requires, or in the encoding their XML declaration names when the runtime provides it. A part in an encoding the runtime does not have is refused.