# SheetStream benchmarks

A reproducible comparison of SheetStream with the other .xlsx writers and readers .NET developers actually use. The suite measures speed, bytes allocated, peak memory and output size. Before anything is timed, it checks that every library really wrote the rows, and that every reader really read them.

| Library | Version | Licence | Why this version |
|---|---|---|---|
| SheetStream | this repo (ProjectReference) | MIT | |
| EPPlus | **4.5.3.3**, pinned `[4.5.3.3]` | LGPL | The last LGPL release. EPPlus 5 and later need a commercial licence for commercial use. |
| ClosedXML | 0.105.1 | MIT | Latest stable. |
| MiniExcel | 1.46.0 | Apache-2.0 | Latest stable. 2.0 is still in preview. |
| NPOI | **2.7.6**, pinned `[2.7.6]` | Apache-2.0 | The last release whose NuGet binaries are plain Apache-2.0. From 2.8.0 the package ships under an "OSMF EULA" with `requireLicenseAcceptance`: users above US$10,000 a year in revenue owe the project a maintenance fee for the binaries. The agreement says this fee is not a licence fee. 2.8.x still builds: it only prints a build warning until you set `AcceptNPOIOSMFLicense`. As with EPPlus 5+, this suite does not accept that agreement on anyone's behalf. |
| ExcelDataReader | 3.9.0 | MIT | Latest stable. Read benchmarks only; it does not write. |
| DocumentFormat.OpenXml | 3.5.1 | MIT | Used only by the correctness guard, to read every library's output back. |
| BenchmarkDotNet | 0.15.8 | MIT | |

NPOI is measured twice on the write side: once with **XSSF**, the in-memory user model, and once with **SXSSF**, the streaming port of Apache POI's SXSSF. On the read side it is also measured twice: **XSSF** (user model) and **SAX** (the event model, `XSSFReader` + `XSSFSheetXMLHandler`).

## How to run

All commands run from `benchmarks/SheetStream.Benchmarks`, and everything is Release.

```powershell
# The two original write scenarios (about 10 minutes on the machine below)
dotnet run -c Release -- --filter '*Long*' '*.WideBenchmarks.*'

# One scenario at a time
dotnet run -c Release -- --filter '*Long*'              # 100k rows x 10 columns
dotnet run -c Release -- --filter '*.WideBenchmarks.*'  # 1,000 x 1,000 ('*Wide*' alone would also match VeryWide)
dotnet run -c Release -- --filter '*Million*'           # 1,000,000 rows x 10 columns (needs 16 GB+ RAM, see below)
dotnet run -c Release -- --filter '*VeryWide*'          # 1,000 rows x 10,000 columns
dotnet run -c Release -- --filter '*Read*'              # read the 100k x 10 file with every reader (about 7 minutes)

# Everything, including the 10M-cell scenarios (well over an hour)
dotnet run -c Release -- --filter '*'

# Peak working set, output size and full-size verification, each library in a fresh process.
# With no scenario it runs all five (Long, Wide, Million, VeryWide, Read); name one or more to run just those.
dotnet run -c Release -- --peak                         # all scenarios
dotnet run -c Release -- --peak Long Wide               # about 3 minutes
dotnet run -c Release -- --peak Read                    # about 3.5 minutes
dotnet run -c Release -- --peak Million VeryWide        # tens of minutes; the in-memory libraries need 16 GB+ RAM

# Correctness guard only: shows which rows, cells and formats every writer actually wrote,
# and the checksum every reader arrived at
dotnet run -c Release -- --verify
```

In Git Bash, quote the filter (`'*'`). Otherwise the shell expands `*` to file names.

BenchmarkDotNet's `-j Dry` switch adds a Dry job next to the one declared on each class rather than replacing it, so a "quick" `--filter '*Read*' -j Dry` still runs the full job as well.

The **Million** scenario holds a million `LongRow` objects (about 80 MB) plus each library's working set. SheetStream and MiniExcel stay close to the baseline, but the in-memory models (EPPlus, ClosedXML, NPOI XSSF) build the whole workbook first and allocate ten times what they do on Long, so run Million on a machine with at least 16 GB of RAM.

EPPlus 4.5.3.3 depends on `System.Drawing.Common`, which only works on Windows under .NET 6+. The EPPlus rows therefore only run on Windows.

## Scenarios

Every library gets the same input. The input is built once, before timing, from deterministic code in `Data.cs`.

* **Long**: 100,000 rows x 10 columns plus a header row in bold. The columns are:
  * `int`
  * three strings, each from a small repeated set (8, 5 and 4 values), so shared-string de-duplication matters
  * three doubles formatted `#,##0.00`
  * a `DateTime` with no time part, formatted `yyyy-mm-dd`
  * a second `int`
* **Wide**: 1,000 rows x 1,000 columns of doubles, with no header and no formatting.
* **Million**: 1,000,000 rows x the same 10 columns as Long (10M cells). Same data generator, same `WriteLong` code; only the row count changes.
* **VeryWide**: 1,000 rows x 10,000 columns of doubles (10M cells, the same count as Million). Same `WriteWide` code as Wide; only the column count changes. Million and VeryWide together show whether a library cares about the shape of a sheet or only about the number of cells.
* **Read**: every reader reads the Long file (100,000 rows x 10 columns) that SheetStream wrote once, outside the timing, into a byte array. Each reader walks every data row and returns a checksum (the sum of the `Id` column plus the number of non-empty cells), so the work cannot be optimised away and the result can be checked against the generator.

## Methodology

* **One implementation per library** (`Writers.cs` for writing, `Readers.cs` for reading). The benchmark, the correctness guard and the peak mode all call the same `WriteLong` / `WriteWide` / `ReadLong` method. What is verified is exactly what is timed.
* **Output goes to a reused `MemoryStream`** (`ReusableOutput`) for every library. It is created once with enough capacity (64 MB for Long and Wide, 160 MB for Million and VeryWide, chosen from the largest output measured: MiniExcel's 76.6 MB on Million) and emptied before each write, so the buffer never grows while timed. As a result, "Allocated" counts only the library's own allocations. Its `Dispose` is a no-op, so a library that closes its output stream does no harm.
* **Each library is written the way its docs recommend:**
  * **SheetStream:** `new ExcelPackage(stream, leaveOpen: true)`, `Row(n).WriteCell(value, style)` with three styles created once, and default options (`CompressionLevel.Optimal`). A second row, **SheetStream (Fastest zip)**, sets only `PackageOptions.CompressionLevel = Fastest`. It is labelled separately because every other library runs at its default compression.
  * **EPPlus 4.5.3.3:** `Worksheet.SetValue(row, col, value)` in a loop. This avoids the `ExcelRange` that `Cells[r, c]` allocates on every call. Number formats are then set once per column range, and the package is written with `SaveAs(stream)`.
  * **ClosedXML:** `ws.Cell(r, c).Value = x` in a loop. `XLCellValue` is a struct, so this does not box. Formats are set once per column range, then `SaveAs(stream)`.
  * **MiniExcel:** `stream.SaveAs(...)`.
    * Long passes the typed `LongRow[]` with `[ExcelColumn(Format = ...)]` on the formatted columns.
    * Wide passes one `Dictionary<string, object>` per row, MiniExcel's documented shape for dynamic columns. This boxes every double.
    * Everything else is left at MiniExcel's defaults.
  * **NPOI:** `CreateRow` / `CreateCell` / `SetCellValue`, with the three `ICellStyle`s created once and `Write(stream, leaveOpen: true)`. SXSSF uses its default 100-row window. It flushes older rows to a temporary file on disk, so its time includes that disk I/O.
* **Each reader is written the way its docs recommend.** The input is the same byte array for all of them, handed over as a fresh read-only `MemoryStream` per invocation (created in `[IterationSetup]`, outside the timing), so nobody pays for disk I/O and a reader that disposes its stream does no harm. Every reader skips the header row and counts data rows only.
  * **SheetStream ReadRows:** `ExcelReader.Open(stream, leaveOpen: true)` and `ReadRows("Data")`, the untyped streaming path; numbers arrive as `double`, shared strings resolved, the date column as `DateTime`.
  * **SheetStream Read&lt;T&gt;:** `Read<LongRow>("Data")`, the typed path that maps the header row to properties by name.
  * **MiniExcel Query&lt;T&gt;:** `stream.Query<LongRow>(sheetName: "Data")`, with the same `LongRow` class the MiniExcel writer uses.
  * **MiniExcel Query (dynamic):** `stream.Query(useHeaderRow: true, sheetName: "Data")`, one `IDictionary<string, object>` per row keyed by header caption.
  * **ExcelDataReader:** `ExcelReaderFactory.CreateReader(stream)` and the `Read()` / `GetValue(i)` loop from its README. Its README also asks .NET Core applications to register `CodePagesEncodingProvider` at start-up; that is done once in a static constructor. It is only needed for binary .xls files, which this suite never opens, but it is the documented set-up step.
  * **ClosedXML:** `new XLWorkbook(stream)`, then `RowsUsed()` / `CellsUsed()` and `Cell(1).GetDouble()`.
  * **EPPlus 4.5.3.3:** `new ExcelPackage(stream)`, then a `Cells[r, c].Value` loop over `Dimension`. Each `Cells[r, c]` allocates an `ExcelRange`; there is no cheaper documented path on the read side.
  * **NPOI XSSF:** `new XSSFWorkbook(stream)`, then `GetRow` / `GetCell` over `LastRowNum`.
  * **NPOI SAX:** `OPCPackage.Open(stream)`, `XSSFReader`, `ReadOnlySharedStringsTable` and an `ISheetContentsHandler`, driven through `XSSFEventBasedExcelExtractor.ProcessSheet`, which is the only parser entry point NPOI exposes for its port of POI's event model. The handler delivers every cell as a *formatted string*, so `Id` is parsed back from text; that cost is inherent to the API. The package is closed with `Revert()` because `OPCPackage.Open(Stream)` opens read-write and `Close()` would try to save.
* **BenchmarkDotNet job:** `[MemoryDiagnoser]` with `SimpleJob(launchCount: 1, warmupCount: 3, iterationCount: 10)` for Long, Wide and Read, and a lighter `SimpleJob(launchCount: 1, warmupCount: 1, iterationCount: 3)` for Million and VeryWide, where the in-memory libraries take 20 to 60 seconds per write. One invocation per iteration throughout. `[IterationSetup]` empties the buffer (or opens the input stream), outside the measurement. Each benchmark runs in its own process, as BenchmarkDotNet does by default, on the default runtime settings (workstation GC, tiered compilation).
* **Correctness guard:** every write benchmark's `GlobalSetup` writes a small case (50 rows for the Long-shaped scenarios, 40 x 40 for the Wide-shaped ones) with every library. It reads each output back with the Open XML SDK's SAX reader and throws if the row count or the number of cells in the last row is wrong. The read benchmark's `GlobalSetup` writes a 50-row file and throws if any reader's checksum differs from the generator's figure (`Checksum.Expected`, computed from the data definition, not from any reader). The `--peak` children repeat both checks on the full-size data after their measurement.
* **Peak working set (`--peak`):** the executable starts itself again with `--peak-child <library> <scenario>`. Each child does the following:
  1. Builds the same input (for Read: writes the 100k-row file with SheetStream into memory).
  2. Writes (or reads) once, cold, including JIT.
  3. Reads `Process.PeakWorkingSet64`.
  4. Only then verifies the full-size output (or compares the checksum).

  Each case runs in 3 fresh processes and the median is reported. A **Baseline** child builds the input and calls no library. It shows the floor every library sits on: the .NET runtime, the benchmark executable and the input data (for Read, the input file as well). The Baseline child never calls a writer or reader, so the libraries' assemblies are not loaded or JIT-compiled in it. Each library's "above baseline" figure therefore includes loading and compiling its own code. The output buffer's pages count only once they are written, so each writer's figure includes roughly its output size.

## Machine

```
BenchmarkDotNet v0.15.8, Windows 11 (10.0.26340.9616)
Intel Core i7-8700 CPU 3.20GHz (Max: 3.19GHz) (Coffee Lake), 1 CPU, 12 logical and 6 physical cores
.NET SDK 10.0.401
  [Host]     : .NET 8.0.31 (8.0.31, 8.0.3126.42015), X64 RyuJIT x86-64-v3
  Job-IUYGTA : .NET 8.0.31 (8.0.31, 8.0.3126.42015), X64 RyuJIT x86-64-v3
```

This is a developer workstation, not an isolated benchmark box. The error columns show the noise this causes. SheetStream on Wide has the widest relative spread (±26% at 99.9%).

## Results

Measured on 2026-10-13. The tables are copied unchanged from `BenchmarkDotNet.Artifacts/results/*-report-github.md` and from the `--peak` output.

### Long: 100,000 rows x 10 columns (BenchmarkDotNet)

```
InvocationCount=1  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=3  
```

| Method                      | Mean       | Error     | StdDev    | Ratio | RatioSD | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|---------------------------- |-----------:|----------:|----------:|------:|--------:|------------:|-----------:|----------:|-----------:|------------:|
| SheetStream                 | 1,012.6 ms |  79.25 ms |  41.45 ms |  1.00 |    0.05 |   4000.0000 |          - |         - |   26.67 MB |        1.00 |
| &#39;SheetStream (Fastest zip)&#39; |   534.1 ms |  71.14 ms |  47.05 ms |  0.53 |    0.05 |   4000.0000 |          - |         - |   26.67 MB |        1.00 |
| &#39;EPPlus 4.5.3.3&#39;            | 2,286.7 ms | 286.31 ms | 189.38 ms |  2.26 |    0.20 |  43000.0000 |  6000.0000 | 2000.0000 |  390.19 MB |       14.63 |
| ClosedXML                   | 3,103.2 ms |  74.72 ms |  44.46 ms |  3.07 |    0.13 | 161000.0000 | 20000.0000 | 2000.0000 | 1104.16 MB |       41.41 |
| MiniExcel                   |   596.4 ms |  66.25 ms |  43.82 ms |  0.59 |    0.05 |  54000.0000 |          - |         - |  343.36 MB |       12.88 |
| &#39;NPOI XSSF&#39;                 | 4,658.4 ms | 478.16 ms | 316.27 ms |  4.61 |    0.35 | 164000.0000 | 53000.0000 | 3000.0000 | 1186.16 MB |       44.48 |
| &#39;NPOI SXSSF&#39;                | 2,039.7 ms | 370.02 ms | 220.19 ms |  2.02 |    0.22 | 117000.0000 | 39000.0000 |         - |  704.22 MB |       26.41 |

### Wide: 1,000 rows x 1,000 columns (BenchmarkDotNet)

```
InvocationCount=1  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=3  
```

| Method                      | Mean        | Error       | StdDev    | Ratio | RatioSD | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|---------------------------- |------------:|------------:|----------:|------:|--------:|------------:|-----------:|----------:|-----------:|------------:|
| SheetStream                 |    984.2 ms |   255.99 ms | 169.32 ms |  1.02 |    0.23 |   6000.0000 |          - |         - |   38.81 MB |        1.00 |
| &#39;SheetStream (Fastest zip)&#39; |    460.0 ms |   103.30 ms |  61.47 ms |  0.48 |    0.09 |   6000.0000 |          - |         - |   38.81 MB |        1.00 |
| &#39;EPPlus 4.5.3.3&#39;            |  1,863.9 ms |    69.74 ms |  41.50 ms |  1.94 |    0.29 |  59000.0000 |  7000.0000 | 2000.0000 |  470.67 MB |       12.13 |
| ClosedXML                   |  2,031.7 ms |   116.52 ms |  77.07 ms |  2.12 |    0.33 | 147000.0000 |  4000.0000 |         - |  976.04 MB |       25.15 |
| MiniExcel                   |    579.0 ms |    45.67 ms |  30.21 ms |  0.60 |    0.10 |  64000.0000 |  1000.0000 |         - |  400.56 MB |       10.32 |
| &#39;NPOI XSSF&#39;                 | 11,106.6 ms | 1,168.99 ms | 773.21 ms | 11.56 |    1.90 | 141000.0000 | 51000.0000 | 4000.0000 | 1040.94 MB |       26.82 |
| &#39;NPOI SXSSF&#39;                |  2,145.3 ms |   282.06 ms | 186.57 ms |  2.23 |    0.38 | 121000.0000 | 24000.0000 |         - |  704.68 MB |       18.16 |

### Peak working set, cold write time and output size (`--peak`)

Each figure is the median of 3 fresh processes. "Cold write time" is a single write that includes JIT, so it is higher than the BenchmarkDotNet means. "Above baseline" subtracts the Baseline child's peak.

**Long**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 36.0 | - | - | - | - |
| SheetStream | 65.6 | 29.6 | 1,203 | 5,798 | yes (100,001 rows x 10 cells) |
| SheetStream-Fastest | 67.0 | 31.0 | 737 | 7,006 | yes (100,001 rows x 10 cells) |
| EPPlus 4.5.3.3 | 219.9 | 183.8 | 3,060 | 5,601 | yes (100,001 rows x 10 cells) |
| ClosedXML | 267.2 | 231.1 | 4,416 | 6,030 | yes (100,001 rows x 10 cells) |
| MiniExcel | 73.3 | 37.3 | 1,016 | 7,421 | yes (100,001 rows x 10 cells) |
| NPOI XSSF | 581.2 | 545.1 | 5,351 | 5,304 | yes (100,001 rows x 10 cells) |
| NPOI SXSSF | 101.2 | 65.2 | 3,310 | 5,518 | yes (100,001 rows x 10 cells) |

**Wide**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 31.7 | - | - | - | - |
| SheetStream | 49.8 | 18.1 | 852 | 4,346 | yes (1,000 rows x 1,000 cells) |
| SheetStream-Fastest | 50.6 | 18.9 | 587 | 5,159 | yes (1,000 rows x 1,000 cells) |
| EPPlus 4.5.3.3 | 215.1 | 183.4 | 2,464 | 4,617 | yes (1,000 rows x 1,000 cells) |
| ClosedXML | 200.9 | 169.2 | 3,447 | 4,772 | yes (1,000 rows x 1,000 cells) |
| MiniExcel | 58.6 | 26.9 | 1,099 | 5,097 | yes (1,000 rows x 1,000 cells) |
| NPOI XSSF | 547.7 | 516.0 | 11,728 | 4,349 | yes (1,000 rows x 1,000 cells) |
| NPOI SXSSF | 136.4 | 104.8 | 3,669 | 4,353 | yes (1,000 rows x 1,000 cells) |

### Million: 1,000,000 rows x 10 columns (BenchmarkDotNet)

Measured on 2026-10-14. Job: `InvocationCount=1 IterationCount=3 LaunchCount=1 UnrollFactor=1 WarmupCount=1` (3 measured writes per library; the error column is wide because of that).

| Method                      | Mean     | Error      | StdDev   | Ratio | RatioSD | Gen0         | Gen1        | Gen2      | Allocated   | Alloc Ratio |
|---------------------------- |---------:|-----------:|---------:|------:|--------:|-------------:|------------:|----------:|------------:|------------:|
| SheetStream                 | 12.825 s |  14.5478 s | 0.7974 s |  1.00 |    0.08 |   45000.0000 |           - |         - |    274.7 MB |        1.00 |
| 'SheetStream (Fastest zip)' |  4.753 s |   1.0805 s | 0.0592 s |  0.37 |    0.02 |   45000.0000 |           - |         - |    274.7 MB |        1.00 |
| 'EPPlus 4.5.3.3'            | 26.228 s | 103.0873 s | 5.6506 s |  2.05 |    0.40 |  409000.0000 |  30000.0000 | 4000.0000 |  3934.59 MB |       14.32 |
| ClosedXML                   | 36.506 s |   9.3365 s | 0.5118 s |  2.85 |    0.16 | 1599000.0000 | 142000.0000 | 5000.0000 | 11559.78 MB |       42.08 |
| MiniExcel                   |  5.325 s |   0.6575 s | 0.0360 s |  0.42 |    0.02 |  553000.0000 |           - |         - |  3325.16 MB |       12.10 |
| 'NPOI XSSF'                 | 47.264 s |  17.7229 s | 0.9715 s |  3.70 |    0.22 | 1616000.0000 | 488000.0000 | 6000.0000 | 11596.94 MB |       42.22 |
| 'NPOI SXSSF'                | 20.738 s |   7.1444 s | 0.3916 s |  1.62 |    0.09 | 1178000.0000 | 393000.0000 |         - |  7049.39 MB |       25.66 |

At a million rows the zip compression level decides the time: SheetStream with `Fastest` is the quickest writer (4.8 s), MiniExcel is close (5.3 s), and SheetStream's default `Optimal` level costs 12.8 s for a file 17% smaller. Allocation scales linearly from the 100,000-row figures: 275 MB for SheetStream against 3.3 GB (MiniExcel), 3.9 GB (EPPlus) and 11.6 GB (ClosedXML, NPOI XSSF).

### VeryWide: 1,000 rows x 10,000 columns (BenchmarkDotNet)

Results pending: run `dotnet run -c Release -- --filter '*VeryWide*'` on the benchmark machine and paste `BenchmarkDotNet.Artifacts/results/SheetStream.Benchmarks.VeryWideBenchmarks-report-github.md` here.

### Read: 100,000 rows x 10 columns (BenchmarkDotNet)

Results pending: run `dotnet run -c Release -- --filter '*Read*'` on the benchmark machine and paste `BenchmarkDotNet.Artifacts/results/SheetStream.Benchmarks.ReadBenchmarks-report-github.md` here.

### Peak working set for Million, VeryWide and Read (`--peak`)

Measured on 2026-10-14 with `dotnet run -c Release -- --peak Million VeryWide Read` (the whole run took about two hours; NPOI XSSF alone needs 32 minutes per VeryWide write). Median of 3 fresh processes.

**Million: 1,000,000 rows x 10 columns**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 121.1 | - | - | - | - |
| SheetStream | 200.4 | 79.3 | 13,303 | 58,079 | yes (1,000,001 rows x 10 cells) |
| SheetStream-Fastest | 212.3 | 91.1 | 13,628 | 70,121 | yes (1,000,001 rows x 10 cells) |
| EPPlus 4.5.3.3 | 1548.1 | 1427.0 | 24,529 | 56,226 | yes (1,000,001 rows x 10 cells) |
| ClosedXML | 2378.4 | 2257.3 | 55,943 | 60,542 | yes (1,000,001 rows x 10 cells) |
| MiniExcel | 220.7 | 99.6 | 7,327 | 74,788 | yes (1,000,001 rows x 10 cells) |
| NPOI XSSF | 5025.6 | 4904.5 | 64,764 | 53,480 | yes (1,000,001 rows x 10 cells) |
| NPOI SXSSF | 225.2 | 104.1 | 29,035 | 55,180 | yes (1,000,001 rows x 10 cells) |

SheetStream's 79 MB above baseline is mostly the 58 MB output buffer: the writer itself holds a few MB. The in-memory libraries need 1.4 GB (EPPlus), 2.3 GB (ClosedXML) and 4.9 GB (NPOI XSSF) for the same file.

**VeryWide: 1,000 rows x 10,000 columns**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 100.3 | - | - | - | - |
| SheetStream | 160.1 | 59.8 | 16,680 | 45,908 | yes (1,000 rows x 10,000 cells) |
| SheetStream-Fastest | 166.0 | 65.7 | 9,942 | 52,144 | yes (1,000 rows x 10,000 cells) |
| EPPlus 4.5.3.3 | 1489.5 | 1389.2 | 72,930 | 49,690 | yes (1,000 rows x 10,000 cells) |
| ClosedXML | 1555.6 | 1455.3 | 48,753 | 48,254 | yes (1,000 rows x 10,000 cells) |
| MiniExcel | 180.4 | 80.1 | 14,691 | 49,460 | yes (1,000 rows x 10,000 cells) |
| NPOI XSSF | 4842.2 | 4741.9 | 1,918,212 | 47,748 | yes (1,000 rows x 10,000 cells) |

The NPOI SXSSF VeryWide row and the Read peak table did not survive the first run's log and are being re-measured; they will be added here.

### What each library actually wrote (`--verify`, small Long case)

| Library | Header bold | Price cell type / format | Date cell type / format |
|---|---|---|---|
| SheetStream (both) | yes | number / `#,##0.00` (built-in 4) | number / `yyyy-mm-dd` |
| EPPlus 4.5.3.3 | yes | number / `#,##0.00` (built-in 4) | number / `yyyy-mm-dd` |
| ClosedXML | yes | number / `#,##0.00` | number / `yyyy-mm-dd` |
| MiniExcel | **no** | number / `#,##0.00` | number / `yyyy-mm-dd` |
| NPOI XSSF / SXSSF | yes | number / `#,##0.00` (built-in 4) | number / `yyyy-mm-dd` |

Built-in number format 4 is `#,##0.00`, so every library stores the same values with the same number formats. The exceptions are in the header and sheet extras:

* **MiniExcel cannot make the header bold.** `OpenXmlHeaderStyle` offers only a background colour, wrap and alignment. Its default header style is a grey fill with borders.
* By default, MiniExcel also adds an auto-filter and freezes the first row on the Long sheet.

Both are small amounts of extra XML and were left at MiniExcel's defaults.

### What each reader read (`--verify`, 50-row Long file written by SheetStream)

Every reader returns the same checksum (1,775 for 50 rows: the Ids 1..50 sum to 1,275, plus 500 non-empty cells), which is also what the generator predicts. The full 100,000-row file gives 5,001,050,000 for every reader in the `--peak Read` children.

## Reading the results honestly

* **Speed at default settings: MiniExcel beats SheetStream in both scenarios.**
  * Long: 596 ms vs 1,013 ms. Wide: 579 ms vs 984 ms.
  * Most of the gap is zip compression. SheetStream defaults to `CompressionLevel.Optimal` and writes the smallest or near-smallest files. MiniExcel's files are 28% larger on Long (7,421 KB vs 5,798 KB) and 17% larger on Wide (5,097 KB vs 4,346 KB).
  * MiniExcel's file sizes are close to what SheetStream writes at `Fastest`, which suggests MiniExcel compresses at a faster level. That is an inference from the output sizes; we did not inspect MiniExcel's source.
  * With the same trade-off (`Fastest`), SheetStream has the lowest mean time in both scenarios: 534 ms on Long and 460 ms on Wide.
  * On Long, the gap between SheetStream (Fastest zip) and MiniExcel (534 vs 596 ms) is inside the error bars. Call it a tie. On Wide, SheetStream (Fastest zip) is ahead (460 vs 579 ms), but SheetStream's Wide runs are the noisiest in the suite.
* **Smallest file: not SheetStream on Long.**
  * NPOI XSSF writes the smallest Long file (5,304 KB), then NPOI SXSSF (5,518 KB) and EPPlus (5,601 KB). SheetStream comes fourth at 5,798 KB.
  * On Wide, SheetStream is the smallest (4,346 KB), with both NPOI variants within 0.2%.
  * For context, NPOI XSSF takes 4.6x (Long) and 11.6x (Wide) as long as SheetStream to write those files.
* **Allocations: SheetStream allocates the least by a wide margin.**
  * Long: 26.7 MB, against 343 MB for MiniExcel (next lowest) and 1.1 to 1.2 GB for ClosedXML and NPOI XSSF.
  * Wide: 38.8 MB, against 401 MB for MiniExcel.
  * SheetStream's allocation is identical at both compression levels. Compression changes time, not allocation.
* **Peak memory: SheetStream is lowest in both scenarios. MiniExcel is close; the in-memory models are not.**
  * Above the baseline, SheetStream holds about 30 MB on Long and 18 MB on Wide, and roughly the output size (4 to 6 MB) of that is the output buffer itself.
  * MiniExcel holds 37 and 27 MB. Both libraries stream.
  * EPPlus, ClosedXML and NPOI XSSF build the whole workbook in memory first: 170 to 545 MB above baseline.
  * NPOI SXSSF streams rows to a temp file but still reaches 65 to 105 MB, and it pays in disk I/O.
* **EPPlus 4.5.3.3** is the fastest of the in-memory models on both scenarios: 2.3x and 1.9x SheetStream's time.
* **NPOI XSSF on Wide** (11.1 s) is the one pathological case in the suite. On this shape, use SXSSF (2.1 s) instead.

## Caveats

* The read benchmarks read a file that SheetStream wrote. Every reader gets the same bytes, but a file from Excel or another writer could lay out shared strings, styles or cell references differently, and a reader tuned for those files might rank differently. The 10-column file is also small enough to fit in memory many times over; a reader's advantage on files that do not is not measured here.
* The read checksum needs the `Id` value and a null check per cell. For the typed readers (`Read<T>`, `Query<T>`) that means the mapping cost of all 10 properties, which is the price of using the typed API; the untyped readers only touch the values they are handed.
* The output stays in memory. Writing to a `FileStream` adds the same disk cost to every library, except NPOI SXSSF, which already does disk I/O for its temp file.
* Million and VeryWide use a lighter job (1 warm-up, 3 measured iterations) than Long and Wide, so their error bars are wider; treat differences of a few percent as noise.
* MiniExcel's Wide path boxes every double into a per-row dictionary. That is MiniExcel's documented shape for dynamic columns, but it explains much of its Wide allocation. An `IDataReader` source might allocate less; it was not measured.
* Every `GlobalSetup` runs the small correctness check for **all** libraries, so each benchmark process has every library's assemblies loaded and JIT-warmed. This affects neither the timed region nor the allocation figures.
