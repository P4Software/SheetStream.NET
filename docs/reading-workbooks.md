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

`ReadDataTable` names columns from the header row (a blank header gets its column letter, a duplicate gets a numeric suffix) and types every column as `object` so mixed columns survive. Pass `firstRowIsHeader: false` to get columns named A, B, C.

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

An `.xlsx` is a zip archive, so a file of a few hundred kilobytes can unpack into gigabytes. A workbook built that way can be 192 KB on the wire and hold a single cell of 200 million characters. `ExcelReader` refuses such files instead of running out of memory: it stops with an `InvalidDataException` whose message names the limit that was exceeded.

| Limit (`ReaderOptions`) | Default | What it bounds |
|---|---|---|
| `MaxCellTextLength` | 1,048,576 characters | The text of one cell or one shared string. Excel's own limit is 32,767 |
| `MaxRowTextLength` | 16,777,216 characters | The text of one row across all its cells, not counting shared strings, which are stored once however many cells use them |
| `MaxLoadedPartBytes` | 256 MB | The unpacked size of each part the reader holds in memory: shared strings, styles, the workbook and its relationships. Sheets are streamed and may be any size |

A tag or attribute far longer than any real workbook contains, and XML nested more than 256 levels deep, are rejected as well.

The defaults sit far above anything Excel writes. An import screen that expects ordinary data can be stricter:

```csharp
var options = new ReaderOptions { MaxCellTextLength = 32_767 };
using var reader = ExcelReader.Open(upload, leaveOpen: false, options);
```

If a genuine file is refused, the message tells you which property to raise.

### What a damaged file can throw

Reading a file that is damaged, truncated or not a workbook at all ends in one of four exception types: `InvalidDataException`, `XmlException`, `FormatException` or `OverflowException`. Catch those around an import and report "this file could not be read" to the user. The reader is fuzz-tested in CI against that contract.

### What the limits do not cover

`ExcelPackage.OpenTemplate` and `ExcelPackage.Edit` load the whole source workbook into memory and do not apply `ReaderOptions`. They are for workbooks your application owns. Do not pass them a file a user uploaded without checking it first; reading it with `ExcelReader` is the safe way to do that.

Collecting rows is your decision too: `reader.ReadRows(...).ToList()` or `ReadDataTable` on a very large sheet holds every row, however carefully each one was read.
