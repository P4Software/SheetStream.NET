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
# One scenario at a time. Each filter names exactly one benchmark class.
dotnet run -c Release -- --filter '*.LongBenchmarks.*'               # 100k rows x 10 mixed columns
dotnet run -c Release -- --filter '*.WideBenchmarks.*'               # 1,000 x 1,000 doubles
dotnet run -c Release -- --filter '*.SpreadCheetahOwnBenchmarks.*'   # 100k x 10 distinct strings: SpreadCheetah's own scenario
dotnet run -c Release -- --filter '*.StringHeavyBenchmarks.*'        # 100k x 10 strings, repeated and unique, with escaping
dotnet run -c Release -- --filter '*.StyledDatesBenchmarks.*'        # 100k x 8 dates, date-times, money and percentages
dotnet run -c Release -- --filter '*.MillionBenchmarks.*'            # 1,000,000 rows x 10 columns (needs 16 GB+ RAM, see below)
dotnet run -c Release -- --filter '*.VeryWideBenchmarks.*'           # 1,000 rows x 10,000 columns
dotnet run -c Release -- --filter '*.ReadBenchmarks.*'               # read the 100k x 10 file with every reader (about 7 minutes)

# Several at once (each class takes a few minutes; the three-launch classes take about three times as long as one launch)
dotnet run -c Release -- --filter '*.LongBenchmarks.*' '*.WideBenchmarks.*'

# Everything, including the 10M-cell scenarios (well over an hour)
dotnet run -c Release -- --filter '*'

# Peak working set, output size and full-size verification, each library in a fresh process.
# With no scenario it runs all of them (see Scenarios.All); name one or more to run just those.
dotnet run -c Release -- --peak                         # all scenarios
dotnet run -c Release -- --peak Long Wide               # about 3 minutes
dotnet run -c Release -- --peak SpreadCheetahOwn StringHeavy StyledDates
dotnet run -c Release -- --peak Read                    # about 3.5 minutes
dotnet run -c Release -- --peak Million VeryWide        # tens of minutes; the in-memory libraries need 16 GB+ RAM

# Correctness guard only: shows which rows, cells and formats every writer actually wrote
# (values are compared cell by cell for the string and styled scenarios), the data profile of the
# string scenarios, and the checksum every reader arrived at
dotnet run -c Release -- --verify
```

In Git Bash, quote the filter (`'*'`). Otherwise the shell expands `*` to file names.

**Filters match the full name of a benchmark, and the matching ignores letter case.** `'*Read*'` therefore also selects every `SpreadCheetah` benchmark in the suite (S-p-**read**-Cheetah), and `'*Wide*'` also selects VeryWide. Use the `'*.<Class>.*'` form above, which selects exactly one class; `--list flat --filter '<pattern>'` shows what a filter would run, without running anything.

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
* **SpreadCheetahOwn**: SpreadCheetah's own published benchmark shape, so every library is also measured on the scenario SpreadCheetah chose for itself. 100,000 rows x 10 columns, every cell a string, no header, no styles. Every value is distinct (`Value <row>-<column>`), so nothing can be de-duplicated and no value repeats in a row or a column. Per its README, SpreadCheetah's benchmark builds each string from its row and column position and writes to an in-memory stream; the text here is ours, the shape is theirs.
* **StringHeavy**: 100,000 rows x 10 string columns, no header, no styles, with a realistic mix. `--verify` prints this table (distinct values, longest value, cells that need XML escaping) for the data it generates:

  | Column | Content | Distinct values |
  |---:|---|---:|
  | 1 | order id (`SO-00000001`) | 100,000 |
  | 2 | status | 5 |
  | 3 | country | 32 |
  | 4 | category | 12 |
  | 5 | customer name (first + last, includes accented letters, an apostrophe and `Smith & Sons`) | 1,920 |
  | 6 | e-mail | 100,000 |
  | 7 | city | 50 |
  | 8 | priority | 3 |
  | 9 | SKU (hexadecimal, unique) | 100,000 |
  | 10 | free-text note, at most 60 characters, 41,666 of them containing `&`, `<`, `>` or `"` | 100,000 |
* **StyledDates**: 100,000 rows x 8 columns plus a bold header row. Three dates (two `yyyy-mm-dd`, one `yyyy-mm-dd hh:mm:ss` with whole seconds), three money columns (`#,##0.00`) and two percentage columns (`0.00%`). Every data cell carries a style, so each row uses four distinct styles (date, date-time, money, percentage), and the header makes five.
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
  * **SpreadCheetah** (all scenarios): one cell array created once and reused for every row (the reuse pattern SpreadCheetah's own benchmark and the code its source generator produces both use; `AddRowAsync` writes the row out before it returns). The cell type is the cheapest one that fits: `DataCell` where no cell has a style (Wide, SpreadCheetahOwn, StringHeavy), `StyledCell` where every cell does (StyledDates), `Cell` where only some do (Long). Styles are added once. It is measured at its own default zip level (Fastest) and at Optimal, next to SheetStream at both levels.
  * **SheetStream** (string scenarios): default options, which means `StringStorage.Auto`: SheetStream looks at the first 128 to 512 cells of every column and then either keeps sharing the column's text (the values repeat and are long or varied) or writes it inline (unique ids, e-mail addresses, free text, short values from a small set); a column that looked unique gets two more looks of 512 cells, at its 512th and 8,192nd cell, which share it if it repeats among thousands of values. Nothing is tuned for the benchmark. The Long, Million, SpreadCheetahOwn and StringHeavy scenarios also run two extra rows, **SheetStream (shared strings, as before)** and **(shared strings, Fastest zip)**, which set only `StringStorage = Shared`: every text through the shared string table, which is what SheetStream did before `StringStorage` existed. They let one run show before and after.
  * **MiniExcel** (string scenarios): one reused `Dictionary<string, object>` per sheet, overwritten for every row. MiniExcel's streaming writer consumes a row before asking for the next, so this is safe and is the cheapest way to feed it; the guard proves every row still arrives intact. StyledDates uses the typed `StyledRow[]` with `[ExcelColumn(Format = ...)]`, like Long.
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
* **BenchmarkDotNet job:** `[MemoryDiagnoser]` with `SimpleJob(launchCount: 3, warmupCount: 3, iterationCount: 10)` for Long, Wide, SpreadCheetahOwn, StringHeavy and StyledDates (30 measured writes per row, in three fresh processes: a single process can settle into a state a few milliseconds slower than the next one on a 55 ms write, and three launches average that), `SimpleJob(launchCount: 1, warmupCount: 3, iterationCount: 10)` for Read, and a lighter `SimpleJob(launchCount: 1, warmupCount: 1, iterationCount: 3)` for Million and VeryWide, where the in-memory libraries take 20 to 60 seconds per write. One invocation per iteration throughout. `[IterationSetup]` empties the buffer (or opens the input stream), outside the measurement. Each benchmark runs in its own process, as BenchmarkDotNet does by default, on the default runtime settings (workstation GC, tiered compilation).
* **Correctness guard:** every write benchmark's `GlobalSetup` writes a small case (50 rows for the Long-shaped scenarios, 40 x 40 for the Wide-shaped ones, 60 rows for the three new scenarios) with every library. It reads each output back with the Open XML SDK's SAX reader and throws if the row count or the number of cells in the last row is wrong. For SpreadCheetahOwn, StringHeavy and StyledDates it also reads every cell back and throws unless each text equals the generator's text (shared and inline strings are resolved first), each number equals the generator's number to within 1e-7 (dates are compared as their serial numbers), and, for StyledDates, each data cell carries the number format of its column. The header's bold font is reported, not enforced, because MiniExcel cannot make a header bold. The value check was itself tested by corrupting one cell in one writer and confirming the guard names that cell. The read benchmark's `GlobalSetup` writes a 50-row file and throws if any reader's checksum differs from the generator's figure (`Checksum.Expected`, computed from the data definition, not from any reader). The `--peak` children repeat both checks on the full-size data after their measurement.
* **Peak working set (`--peak`):** the executable starts itself again with `--peak-child <library> <scenario>`. Each child does the following:
  1. Builds the same input (for Read: writes the 100k-row file with SheetStream into memory).
  2. Writes (or reads) once, cold, including JIT.
  3. Reads `Process.PeakWorkingSet64`.
  4. Only then verifies the full-size output (or compares the checksum).

  Each case runs in 3 fresh processes and the median is reported. A **Baseline** child builds the input and calls no library. It shows the floor every library sits on: the .NET runtime, the benchmark executable and the input data (for Read, the input file as well). The Baseline child never calls a writer or reader, so the libraries' assemblies are not loaded or JIT-compiled in it. Each library's "above baseline" figure therefore includes loading and compiling its own code. The output buffer's pages count only once they are written, so each writer's figure includes roughly its output size.

## Reproduce these numbers

On a clean Windows machine (a new virtual machine is ideal) with only the .NET 10 SDK installed, copy or clone this repository, open the `benchmarks` folder and double-click `run-on-fresh-vm.cmd`, or run:

```powershell
.\run-on-fresh-vm.ps1 -DryRun                      # print the plan and the hardware, change nothing
.\run-on-fresh-vm.ps1                              # SpreadCheetahOwn, StringHeavy, StyledDates, Long and Wide, then --peak
.\run-on-fresh-vm.ps1 -Scenarios SpreadCheetahOwn  # one scenario (comma-separated list, or All)
```

The script does the following, in order, and stops with a plain-language message if a step fails:

1. Prints the CPU, cores, memory, operating system, .NET SDK, code version and the exact package versions.
2. Builds the benchmark project in Release.
3. Runs `--verify`. **If any library fails the correctness check the script stops and takes no timings.**
4. Runs each chosen scenario through BenchmarkDotNet.
5. Runs `--peak` for the same scenarios.
6. Copies the BenchmarkDotNet reports, the logs and a `hardware.txt` into `benchmarks\results\<date>-<time>-<machine>` (a folder that Git ignores, so a run is never committed by accident).

The run behind the tables below took 2 h 17 min for the BenchmarkDotNet stage (all eight scenarios, with three launches where the job says so) and 65 minutes for `--peak`. Leave the machine alone while it runs, and use a plugged-in, high-performance power setup. The tables below the "Results" heading are copied unchanged from the `*-report-github.md` files and the `peak.log` of such a run; each table says which machine and date it came from.

### Fairness rules

These are the rules every scenario follows. If a result here ever breaks one, the result is wrong, not the rule.

1. **Same data.** Every library writes the same in-memory input, generated by deterministic code in `Data.cs` and built once outside the timed region. `--verify` prints the data profile of the string scenarios so anyone can check what "unique" and "repeated" mean.
2. **Same zip level, paired.** SheetStream and SpreadCheetah each appear at both zip levels (default Optimal and Fastest for SheetStream; SpreadCheetah's default Fastest and Optimal), so every row has a row from the other library at the same level to be read against. A claim that SheetStream is faster must name the zip level it compares at.
3. **Correctness first.** Before any timing, every library must write exactly the data it was given: row count, cell count, and for the string and styled scenarios every cell's value and number format. A library that fails is fixed or removed; it is never timed. The same checks run again at full size in every `--peak` child.
4. **Each library the way its documentation recommends.** SpreadCheetah with a reused cell array and the cheapest cell type for the scenario (`DataCell`, `StyledCell` or `Cell`), so it is never handicapped by an avoidable allocation. SheetStream at its default options. Anything a library cannot express is listed in this file, not hidden (for example, MiniExcel cannot make a header bold).
5. **Same measuring.** Output goes to the same reused memory buffer for every library, so disk speed and buffer growth never enter the comparison; every benchmark runs in its own process with the same job settings; every scenario that SpreadCheetah chose for itself is included next to the ones we chose.
6. **Whole picture reported.** Time, bytes allocated, peak memory and output size are all reported, including the ones where SheetStream is not first. "Reading the results honestly" below is written from the measured tables, not before them.
7. **One machine per table.** Numbers from different machines are never mixed in one table; each table names its machine (from the hardware record, `hardware.txt` or `environment.txt`), date and code version.

## Machine

Every table below was measured on one machine, a virtual machine created for this purpose and used for nothing else:

```
BenchmarkDotNet v0.15.8, Windows 11 (10.0.26100.33730/24H2/2024Update/HudsonValley) (Hyper-V)
INTEL XEON PLATINUM 8573C 2.30GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 10.0.404
  [Host]     : .NET 10.0.15 (10.0.15, 10.0.1526.53008), X64 RyuJIT x86-64-v4
```

The operating system is Windows Server 2025 Datacenter: Azure Edition; BenchmarkDotNet names the build by its Windows 11 equivalent (both are build 26100).

* Azure `Standard_D8s_v6` in West US 3, 32 GB of memory, power plan "High performance".
* Windows Defender real-time protection was on, as it is on a default Windows Server install. The processor was 99.6% to 99.9% idle before the run started and after it finished.
* Date: 2026-11-18. Code: commit `3d7882a` of the `string-storage` branch (the library code is that of `3b6a8af`; `3d7882a` changes only the benchmark job settings).
* The benchmark runs the library through `ProjectReference`, in Release, on the default runtime settings (workstation GC, tiered compilation).

The raw output of this run (every BenchmarkDotNet report, `peak.log`, the correctness-guard log, the hardware record and the package list) is in [`published-results/2026-11-18-azure-d8s-v6`](published-results/2026-11-18-azure-d8s-v6), together with the two earlier full runs on the same machine that the consistency table below uses.

## Results

The tables are copied unchanged from BenchmarkDotNet's `*-report-github.md` files and from the `--peak` output of one run. The Long, Wide, SpreadCheetahOwn, StringHeavy and StyledDates jobs use **three launches** (three fresh processes per row, 10 measured writes in each), so every mean in those tables is the mean of 30 writes. A single launch can settle into a slower or faster state than the next; three launches average that out and show it in the StdDev column.

Before any timing, the correctness guard (`--verify`) wrote every scenario with every library and read each file back: **all libraries passed** (`verify.log`).

### Long: 100,000 rows x 10 columns (BenchmarkDotNet)

Mixed columns with a bold header; three of the text columns repeat (8, 5 and 4 values).

Job: `InvocationCount=1 IterationCount=10 LaunchCount=3 UnrollFactor=1 WarmupCount=3`

| Method                                      | Mean        | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated     | Alloc Ratio |
|-------------------------------------------- |------------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|--------------:|------------:|
| SheetStream                                 |   248.25 ms |  0.731 ms |  1.094 ms |  1.00 |    0.01 |          - |          - |         - |     164.18 KB |        1.00 |
| &#39;SheetStream (Fastest zip)&#39;                 |    79.95 ms |  0.463 ms |  0.692 ms |  0.32 |    0.00 |          - |          - |         - |     164.24 KB |        1.00 |
| &#39;SheetStream (shared strings, as before)&#39;   |   313.10 ms |  1.249 ms |  1.869 ms |  1.26 |    0.01 |          - |          - |         - |      163.6 KB |        1.00 |
| &#39;SheetStream (shared strings, Fastest zip)&#39; |    94.64 ms |  0.434 ms |  0.636 ms |  0.38 |    0.00 |          - |          - |         - |     163.68 KB |        1.00 |
| &#39;SpreadCheetah (default, Fastest zip)&#39;      |   106.82 ms |  0.578 ms |  0.865 ms |  0.43 |    0.00 |          - |          - |         - |      12.45 KB |        0.08 |
| &#39;SpreadCheetah (Optimal zip)&#39;               |   274.00 ms |  0.920 ms |  1.378 ms |  1.10 |    0.01 |          - |          - |         - |      12.45 KB |        0.08 |
| &#39;EPPlus 4.5.3.3&#39;                            | 2,172.33 ms | 10.270 ms | 15.053 ms |  8.75 |    0.07 |  4000.0000 |  2000.0000 | 1000.0000 |  399505.88 KB |    2,433.35 |
| ClosedXML                                   | 2,441.78 ms | 16.758 ms | 24.033 ms |  9.84 |    0.10 | 11000.0000 |  9000.0000 | 1000.0000 | 1036917.61 KB |    6,315.75 |
| MiniExcel                                   |   335.52 ms |  2.759 ms |  4.044 ms |  1.35 |    0.02 |  3000.0000 |          - |         - |  316404.79 KB |    1,927.19 |
| &#39;NPOI XSSF&#39;                                 | 4,354.03 ms | 55.997 ms | 80.309 ms | 17.54 |    0.33 | 13000.0000 | 12000.0000 | 2000.0000 | 1176735.38 KB |    7,167.36 |
| &#39;NPOI SXSSF&#39;                                | 1,582.29 ms | 59.085 ms | 86.606 ms |  6.37 |    0.34 |  7000.0000 |  2000.0000 |         - |  651584.13 KB |    3,968.73 |

### Wide: 1,000 rows x 1,000 columns (BenchmarkDotNet)

Doubles only, no styles.

Job: `InvocationCount=1 IterationCount=10 LaunchCount=3 UnrollFactor=1 WarmupCount=3`

| Method                                 | Mean        | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated     | Alloc Ratio |
|--------------------------------------- |------------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|--------------:|------------:|
| SheetStream                            |   264.36 ms |  0.987 ms |  1.383 ms |  1.00 |    0.01 |          - |          - |         - |     143.55 KB |        1.00 |
| &#39;SheetStream (Fastest zip)&#39;            |    58.47 ms |  0.348 ms |  0.500 ms |  0.22 |    0.00 |          - |          - |         - |     143.61 KB |        1.00 |
| &#39;SpreadCheetah (default, Fastest zip)&#39; |   101.26 ms |  0.525 ms |  0.785 ms |  0.38 |    0.00 |          - |          - |         - |       34.3 KB |        0.24 |
| &#39;SpreadCheetah (Optimal zip)&#39;          |   306.97 ms |  0.713 ms |  1.045 ms |  1.16 |    0.01 |          - |          - |         - |       34.3 KB |        0.24 |
| &#39;EPPlus 4.5.3.3&#39;                       | 1,401.13 ms |  7.245 ms | 10.391 ms |  5.30 |    0.05 |  5000.0000 |  2000.0000 | 1000.0000 |  481930.07 KB |    3,357.30 |
| ClosedXML                              | 1,452.91 ms | 12.335 ms | 18.463 ms |  5.50 |    0.07 | 10000.0000 |  2000.0000 |         - |  952543.78 KB |    6,635.77 |
| MiniExcel                              |   331.68 ms |  2.613 ms |  3.747 ms |  1.25 |    0.02 |  4000.0000 |  1000.0000 |         - |  371092.91 KB |    2,585.17 |
| &#39;NPOI XSSF&#39;                            | 8,086.80 ms | 25.857 ms | 37.901 ms | 30.59 |    0.21 | 12000.0000 | 11000.0000 | 2000.0000 | 1065900.84 KB |    7,425.45 |
| &#39;NPOI SXSSF&#39;                           | 1,609.69 ms | 29.827 ms | 40.828 ms |  6.09 |    0.15 |  8000.0000 |  7000.0000 |         - |  674066.31 KB |    4,695.79 |

### SpreadCheetahOwn: 100,000 rows x 10 text columns, every value distinct (BenchmarkDotNet)

SpreadCheetah's own benchmark shape. Nothing can be shared, so this is the hardest case for SheetStream's `StringStorage.Auto`. Read "Start-up writes" below before using the SheetStream (Fastest zip) mean.

Job: `InvocationCount=1 IterationCount=10 LaunchCount=3 UnrollFactor=1 WarmupCount=3`

| Method                                      | Mean        | Error     | StdDev    | Median      | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated     | Alloc Ratio |
|-------------------------------------------- |------------:|----------:|----------:|------------:|------:|--------:|-----------:|-----------:|----------:|--------------:|------------:|
| SheetStream                                 |   172.20 ms |  0.697 ms |  0.954 ms |   172.25 ms |  1.00 |    0.01 |          - |          - |         - |    1336.41 KB |       1.000 |
| &#39;SheetStream (Fastest zip)&#39;                 |    81.79 ms | 26.389 ms | 39.497 ms |    56.28 ms |  0.47 |    0.23 |          - |          - |         - |    1336.47 KB |       1.000 |
| &#39;SheetStream (shared strings, as before)&#39;   |   351.29 ms |  1.685 ms |  2.417 ms |   351.56 ms |  2.04 |    0.02 |          - |          - |         - |   43875.64 KB |      32.831 |
| &#39;SheetStream (shared strings, Fastest zip)&#39; |   162.34 ms |  1.443 ms |  2.159 ms |   161.98 ms |  0.94 |    0.01 |          - |          - |         - |   43875.72 KB |      32.831 |
| &#39;SpreadCheetah (default, Fastest zip)&#39;      |    54.95 ms |  0.510 ms |  0.663 ms |    55.01 ms |  0.32 |    0.00 |          - |          - |         - |      11.13 KB |       0.008 |
| &#39;SpreadCheetah (Optimal zip)&#39;               |   168.72 ms |  0.596 ms |  0.874 ms |   168.77 ms |  0.98 |    0.01 |          - |          - |         - |      11.13 KB |       0.008 |
| &#39;EPPlus 4.5.3.3&#39;                            | 2,202.55 ms | 11.142 ms | 16.332 ms | 2,199.64 ms | 12.79 |    0.12 |  5000.0000 |  4000.0000 | 2000.0000 |  534375.66 KB |     399.860 |
| ClosedXML                                   | 2,673.73 ms | 15.736 ms | 23.552 ms | 2,670.22 ms | 15.53 |    0.16 | 13000.0000 |  4000.0000 | 1000.0000 | 1481312.91 KB |   1,108.430 |
| MiniExcel                                   |   353.60 ms |  3.708 ms |  5.435 ms |   351.75 ms |  2.05 |    0.03 |  4000.0000 |          - |         - |  420807.95 KB |     314.880 |
| &#39;NPOI XSSF&#39;                                 | 6,857.18 ms | 28.378 ms | 42.474 ms | 6,860.84 ms | 39.82 |    0.32 | 26000.0000 | 23000.0000 | 3000.0000 | 2473359.71 KB |   1,850.754 |
| &#39;NPOI SXSSF&#39;                                | 1,559.35 ms | 55.658 ms | 76.185 ms | 1,526.31 ms |  9.06 |    0.44 |  8000.0000 |  2000.0000 |         - |   697414.7 KB |     521.858 |

### StringHeavy: 100,000 rows x 10 text columns, a realistic mix (BenchmarkDotNet)

Repeated and unique text, accented letters and XML special characters. Read "Start-up writes" below before using the SheetStream (Fastest zip) mean.

Job: `InvocationCount=1 IterationCount=10 LaunchCount=3 UnrollFactor=1 WarmupCount=3`

| Method                                      | Mean        | Error     | StdDev    | Median      | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated     | Alloc Ratio |
|-------------------------------------------- |------------:|----------:|----------:|------------:|------:|--------:|-----------:|-----------:|----------:|--------------:|------------:|
| SheetStream                                 |   196.05 ms |  0.996 ms |  1.428 ms |   196.15 ms |  1.00 |    0.01 |          - |          - |         - |      729.3 KB |        1.00 |
| &#39;SheetStream (Fastest zip)&#39;                 |    92.03 ms | 22.165 ms | 33.175 ms |    78.62 ms |  0.47 |    0.17 |          - |          - |         - |     729.36 KB |        1.00 |
| &#39;SheetStream (shared strings, as before)&#39;   |   406.66 ms |  1.814 ms |  2.659 ms |   407.01 ms |  2.07 |    0.02 |          - |          - |         - |   25477.98 KB |       34.93 |
| &#39;SheetStream (shared strings, Fastest zip)&#39; |   193.53 ms |  1.229 ms |  1.840 ms |   193.67 ms |  0.99 |    0.01 |          - |          - |         - |   25478.05 KB |       34.94 |
| &#39;SpreadCheetah (default, Fastest zip)&#39;      |    79.68 ms |  0.406 ms |  0.607 ms |    79.73 ms |  0.41 |    0.00 |          - |          - |         - |      11.13 KB |        0.02 |
| &#39;SpreadCheetah (Optimal zip)&#39;               |   196.27 ms |  0.914 ms |  1.311 ms |   196.26 ms |  1.00 |    0.01 |          - |          - |         - |      11.13 KB |        0.02 |
| &#39;EPPlus 4.5.3.3&#39;                            | 1,982.07 ms |  8.011 ms | 11.991 ms | 1,977.62 ms | 10.11 |    0.09 |  3000.0000 |  2000.0000 | 1000.0000 |  430956.65 KB |      590.92 |
| ClosedXML                                   | 2,274.98 ms | 13.626 ms | 20.395 ms | 2,274.31 ms | 11.60 |    0.13 | 12000.0000 |  4000.0000 | 1000.0000 | 1260362.28 KB |    1,728.19 |
| MiniExcel                                   |   417.17 ms |  1.925 ms |  2.822 ms |   416.70 ms |  2.13 |    0.02 |  5000.0000 |          - |         - |  437170.77 KB |      599.44 |
| &#39;NPOI XSSF&#39;                                 | 5,750.45 ms | 34.948 ms | 52.308 ms | 5,745.28 ms | 29.33 |    0.34 | 22000.0000 | 20000.0000 | 3000.0000 | 2030618.88 KB |    2,784.35 |
| &#39;NPOI SXSSF&#39;                                | 1,514.09 ms | 62.735 ms | 89.972 ms | 1,466.70 ms |  7.72 |    0.45 |  8000.0000 |  2000.0000 |         - |  697830.42 KB |      956.85 |

### StyledDates: 100,000 rows x 8 styled columns (BenchmarkDotNet)

Dates, date-times, money and percentages; every data cell carries a style.

Job: `InvocationCount=1 IterationCount=10 LaunchCount=3 UnrollFactor=1 WarmupCount=3`

| Method                                 | Mean       | Error    | StdDev   | Ratio | RatioSD | Gen0       | Gen1      | Gen2      | Allocated    | Alloc Ratio |
|--------------------------------------- |-----------:|---------:|---------:|------:|--------:|-----------:|----------:|----------:|-------------:|------------:|
| SheetStream                            |   341.4 ms |  1.33 ms |  1.95 ms |  1.00 |    0.01 |          - |         - |         - |     163.5 KB |        1.00 |
| &#39;SheetStream (Fastest zip)&#39;            |   104.7 ms |  0.60 ms |  0.89 ms |  0.31 |    0.00 |          - |         - |         - |    163.56 KB |        1.00 |
| &#39;SpreadCheetah (default, Fastest zip)&#39; |   121.9 ms |  0.76 ms |  1.13 ms |  0.36 |    0.00 |          - |         - |         - |     14.02 KB |        0.09 |
| &#39;SpreadCheetah (Optimal zip)&#39;          |   351.7 ms |  1.10 ms |  1.58 ms |  1.03 |    0.01 |          - |         - |         - |     13.95 KB |        0.09 |
| &#39;EPPlus 4.5.3.3&#39;                       | 1,964.5 ms |  4.12 ms |  5.77 ms |  5.75 |    0.04 |  4000.0000 | 2000.0000 | 1000.0000 | 387427.76 KB |    2,369.59 |
| ClosedXML                              | 2,511.4 ms | 18.55 ms | 27.19 ms |  7.36 |    0.09 | 10000.0000 | 8000.0000 | 1000.0000 | 934491.15 KB |    5,715.54 |
| MiniExcel                              |   300.7 ms |  1.60 ms |  2.30 ms |  0.88 |    0.01 |  2000.0000 |         - |         - | 233493.67 KB |    1,428.10 |
| &#39;NPOI XSSF&#39;                            | 3,994.5 ms | 16.15 ms | 24.18 ms | 11.70 |    0.10 | 10000.0000 | 9000.0000 | 2000.0000 | 814260.51 KB |    4,980.19 |
| &#39;NPOI SXSSF&#39;                           | 1,585.8 ms | 38.22 ms | 54.81 ms |  4.65 |    0.16 |  6000.0000 | 1000.0000 |         - |  518866.8 KB |    3,173.50 |

### Million: 1,000,000 rows x 10 columns (BenchmarkDotNet)

The Long data with ten times the rows (10 million cells). Light job: three measured writes per library, so read differences of a few percent as noise.

Job: `InvocationCount=1 IterationCount=3 LaunchCount=1 UnrollFactor=1 WarmupCount=1`

| Method                                      | Mean        | Error       | StdDev    | Ratio | RatioSD | Gen0        | Gen1        | Gen2      | Allocated      | Alloc Ratio |
|-------------------------------------------- |------------:|------------:|----------:|------:|--------:|------------:|------------:|----------:|---------------:|------------:|
| SheetStream                                 |  2,542.8 ms |   172.84 ms |   9.47 ms |  1.00 |    0.00 |           - |           - |         - |      164.18 KB |        1.00 |
| &#39;SheetStream (Fastest zip)&#39;                 |    813.4 ms |    38.84 ms |   2.13 ms |  0.32 |    0.00 |           - |           - |         - |      164.24 KB |        1.00 |
| &#39;SheetStream (shared strings, as before)&#39;   |  3,238.9 ms |    26.44 ms |   1.45 ms |  1.27 |    0.00 |           - |           - |         - |       163.6 KB |        1.00 |
| &#39;SheetStream (shared strings, Fastest zip)&#39; |    956.3 ms |    52.28 ms |   2.87 ms |  0.38 |    0.00 |           - |           - |         - |      163.68 KB |        1.00 |
| &#39;SpreadCheetah (default, Fastest zip)&#39;      |  1,054.2 ms |    90.04 ms |   4.94 ms |  0.41 |    0.00 |           - |           - |         - |       13.88 KB |        0.08 |
| &#39;SpreadCheetah (Optimal zip)&#39;               |  2,934.3 ms |    69.44 ms |   3.81 ms |  1.15 |    0.00 |           - |           - |         - |       13.85 KB |        0.08 |
| &#39;EPPlus 4.5.3.3&#39;                            | 22,087.0 ms | 1,697.90 ms |  93.07 ms |  8.69 |    0.04 |  34000.0000 |   7000.0000 | 4000.0000 |     4028665 KB |   24,538.15 |
| ClosedXML                                   | 32,798.5 ms | 4,047.96 ms | 221.88 ms | 12.90 |    0.09 | 113000.0000 |  82000.0000 | 5000.0000 | 10899731.45 KB |   66,389.04 |
| MiniExcel                                   |  3,519.0 ms |   157.44 ms |   8.63 ms |  1.38 |    0.01 |  37000.0000 |           - |         - |  3052944.63 KB |   18,595.14 |
| &#39;NPOI XSSF&#39;                                 | 45,740.9 ms | 1,418.03 ms |  77.73 ms | 17.99 |    0.06 | 122000.0000 | 111000.0000 | 6000.0000 | 11492840.85 KB |   70,001.60 |
| &#39;NPOI SXSSF&#39;                                | 14,983.3 ms | 3,459.38 ms | 189.62 ms |  5.89 |    0.07 |  79000.0000 |  26000.0000 |         - |  6523266.55 KB |   39,732.48 |

### VeryWide: 1,000 rows x 10,000 columns (BenchmarkDotNet)

The Wide data with ten times the columns (10 million cells). Light job, as Million. NPOI XSSF takes about nine minutes per write on this shape.

Job: `InvocationCount=1 IterationCount=3 LaunchCount=1 UnrollFactor=1 WarmupCount=1`

| Method                                 | Mean         | Error       | StdDev    | Ratio  | RatioSD | Gen0        | Gen1        | Gen2       | Allocated      | Alloc Ratio |
|--------------------------------------- |-------------:|------------:|----------:|-------:|--------:|------------:|------------:|-----------:|---------------:|------------:|
| SheetStream                            |   2,621.1 ms |   118.52 ms |   6.50 ms |   1.00 |    0.00 |           - |           - |          - |      143.55 KB |        1.00 |
| &#39;SheetStream (Fastest zip)&#39;            |     597.6 ms |    30.42 ms |   1.67 ms |   0.23 |    0.00 |           - |           - |          - |      143.61 KB |        1.00 |
| &#39;SpreadCheetah (default, Fastest zip)&#39; |   1,041.4 ms |    79.10 ms |   4.34 ms |   0.40 |    0.00 |           - |           - |          - |      245.24 KB |        1.71 |
| &#39;SpreadCheetah (Optimal zip)&#39;          |   3,084.4 ms |   158.37 ms |   8.68 ms |   1.18 |    0.00 |           - |           - |          - |      245.24 KB |        1.71 |
| &#39;EPPlus 4.5.3.3&#39;                       |  17,661.9 ms | 5,187.69 ms | 284.35 ms |   6.74 |    0.10 |  50000.0000 |   7000.0000 |  3000.0000 |  5385708.95 KB |   37,518.82 |
| ClosedXML                              |  15,107.8 ms |   275.46 ms |  15.10 ms |   5.76 |    0.01 | 109000.0000 |  19000.0000 |  4000.0000 | 10178297.51 KB |   70,905.74 |
| MiniExcel                              |   4,956.8 ms |   246.00 ms |  13.48 ms |   1.89 |    0.01 |  51000.0000 |  36000.0000 |  7000.0000 |   3907738.7 KB |   27,222.74 |
| &#39;NPOI XSSF&#39;                            | 526,084.8 ms | 4,517.72 ms | 247.63 ms | 200.71 |    0.44 | 125000.0000 | 124000.0000 |  7000.0000 | 11740988.42 KB |   81,792.02 |
| &#39;NPOI SXSSF&#39;                           |  17,053.7 ms | 2,520.44 ms | 138.15 ms |   6.51 |    0.05 |  92000.0000 |  86000.0000 | 11000.0000 |  7394866.21 KB |   51,515.34 |

### Read: 100,000 rows x 10 columns (BenchmarkDotNet)

Every reader reads the same 5.8 MB file that SheetStream wrote. This job still uses one launch.

Job: `InvocationCount=1 IterationCount=10 LaunchCount=1 UnrollFactor=1 WarmupCount=3`

| Method                      | Mean       | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|---------------------------- |-----------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|-----------:|------------:|
| &#39;SheetStream ReadRows&#39;      |   350.8 ms |   1.72 ms |   1.14 ms |  1.00 |    0.00 |  1000.0000 |          - |         - |  124.39 MB |        1.00 |
| &#39;SheetStream Read&lt;T&gt;&#39;       |   382.2 ms |   1.66 ms |   1.10 ms |  1.09 |    0.00 |  1000.0000 |          - |         - |  139.65 MB |        1.12 |
| &#39;MiniExcel Query&lt;T&gt;&#39;        | 1,458.7 ms |   5.93 ms |   3.92 ms |  4.16 |    0.02 | 27000.0000 |          - |         - | 2156.58 MB |       17.34 |
| &#39;MiniExcel Query (dynamic)&#39; | 1,431.4 ms |   5.20 ms |   3.44 ms |  4.08 |    0.02 | 27000.0000 |          - |         - | 2182.21 MB |       17.54 |
| ExcelDataReader             |   693.4 ms |   3.03 ms |   2.00 ms |  1.98 |    0.01 |  3000.0000 |          - |         - |  249.39 MB |        2.00 |
| ClosedXML                   | 2,663.2 ms |  16.17 ms |   9.62 ms |  7.59 |    0.04 | 20000.0000 | 12000.0000 | 2000.0000 | 1508.51 MB |       12.13 |
| &#39;EPPlus 4.5.3.3&#39;            | 1,162.9 ms |  11.87 ms |   7.06 ms |  3.31 |    0.02 | 11000.0000 |  7000.0000 | 3000.0000 |  804.63 MB |        6.47 |
| &#39;NPOI XSSF&#39;                 | 3,526.4 ms | 431.90 ms | 285.68 ms | 10.05 |    0.78 | 17000.0000 | 16000.0000 | 3000.0000 | 1277.18 MB |       10.27 |
| &#39;NPOI SAX&#39;                  | 2,028.5 ms |   9.06 ms |   4.74 ms |  5.78 |    0.02 | 23000.0000 |  1000.0000 | 1000.0000 | 1838.81 MB |       14.78 |

### Peak working set, cold write time and output size (`--peak`)

Each figure is the median of 3 fresh processes. "Cold write time" is a single write that includes JIT, so it is higher than the BenchmarkDotNet means. "Above baseline" subtracts the Baseline child's peak. `SheetStream-Fastest` is SheetStream with `CompressionLevel.Fastest`; `SpreadCheetah` is SpreadCheetah at its default (Fastest) level.


**Peak working set: Long (3 fresh processes each, median shown)**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 34.8 | - | - | - | - |
| SheetStream | 49.2 | 14.4 | 392 | 2,917 | yes (100,001 rows x 10 cells) |
| SheetStream-Fastest | 51.5 | 16.7 | 256 | 5,304 | yes (100,001 rows x 10 cells) |
| SheetStream-Shared | 49.2 | 14.4 | 495 | 2,927 | yes (100,001 rows x 10 cells) |
| SheetStream-Shared-Fastest | 51.6 | 16.8 | 345 | 5,609 | yes (100,001 rows x 10 cells) |
| SpreadCheetah | 49.7 | 15.0 | 375 | 5,296 | yes (100,001 rows x 10 cells) |
| SpreadCheetah-Optimal | 47.5 | 12.8 | 506 | 2,916 | yes (100,001 rows x 10 cells) |
| EPPlus 4.5.3.3 | 321.7 | 287.0 | 3,249 | 5,601 | yes (100,001 rows x 10 cells) |
| ClosedXML | 386.2 | 351.5 | 4,488 | 5,955 | yes (100,001 rows x 10 cells) |
| MiniExcel | 149.4 | 114.6 | 869 | 10,420 | yes (100,001 rows x 10 cells) |
| NPOI XSSF | 654.6 | 619.9 | 5,965 | 5,304 | yes (100,001 rows x 10 cells) |
| NPOI SXSSF | 165.7 | 130.9 | 3,237 | 5,518 | yes (100,001 rows x 10 cells) |

**Peak working set: Wide (3 fresh processes each, median shown)**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 32.8 | - | - | - | - |
| SheetStream | 44.2 | 11.3 | 360 | 2,050 | yes (1,000 rows x 1,000 cells) |
| SheetStream-Fastest | 45.1 | 12.2 | 186 | 3,673 | yes (1,000 rows x 1,000 cells) |
| SheetStream-Shared | 44.1 | 11.3 | 358 | 2,050 | yes (1,000 rows x 1,000 cells) |
| SheetStream-Shared-Fastest | 45.2 | 12.3 | 187 | 3,673 | yes (1,000 rows x 1,000 cells) |
| SpreadCheetah | 45.7 | 12.9 | 301 | 3,673 | yes (1,000 rows x 1,000 cells) |
| SpreadCheetah-Optimal | 44.3 | 11.5 | 461 | 2,050 | yes (1,000 rows x 1,000 cells) |
| EPPlus 4.5.3.3 | 309.9 | 277.0 | 2,350 | 4,617 | yes (1,000 rows x 1,000 cells) |
| ClosedXML | 269.0 | 236.2 | 3,204 | 4,779 | yes (1,000 rows x 1,000 cells) |
| MiniExcel | 137.2 | 104.3 | 924 | 7,394 | yes (1,000 rows x 1,000 cells) |
| NPOI XSSF | 580.7 | 547.9 | 9,592 | 4,349 | yes (1,000 rows x 1,000 cells) |
| NPOI SXSSF | 172.1 | 139.2 | 3,229 | 4,353 | yes (1,000 rows x 1,000 cells) |

**Peak working set: SpreadCheetahOwn (3 fresh processes each, median shown)**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 114.2 | - | - | - | - |
| SheetStream | 125.6 | 11.4 | 322 | 2,894 | yes (100,000 rows x 10 cells) |
| SheetStream-Fastest | 126.5 | 12.3 | 221 | 4,681 | yes (100,000 rows x 10 cells) |
| SheetStream-Shared | 159.4 | 45.2 | 667 | 3,678 | yes (100,000 rows x 10 cells) |
| SheetStream-Shared-Fastest | 161.8 | 47.6 | 514 | 6,022 | yes (100,000 rows x 10 cells) |
| SpreadCheetah | 125.8 | 11.6 | 279 | 4,637 | yes (100,000 rows x 10 cells) |
| SpreadCheetah-Optimal | 124.5 | 10.3 | 365 | 2,860 | yes (100,000 rows x 10 cells) |
| EPPlus 4.5.3.3 | 413.3 | 299.1 | 3,486 | 5,729 | yes (100,000 rows x 10 cells) |
| ClosedXML | 541.7 | 427.5 | 4,806 | 5,513 | yes (100,000 rows x 10 cells) |
| MiniExcel | 208.3 | 94.1 | 1,162 | 5,603 | yes (100,000 rows x 10 cells) |
| NPOI XSSF | 1081.4 | 967.2 | 8,248 | 5,593 | yes (100,000 rows x 10 cells) |
| NPOI SXSSF | 209.0 | 94.7 | 3,383 | 3,423 | yes (100,000 rows x 10 cells) |

**Peak working set: StringHeavy (3 fresh processes each, median shown)**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 91.8 | - | - | - | - |
| SheetStream | 102.6 | 10.7 | 341 | 3,146 | yes (100,000 rows x 10 cells) |
| SheetStream-Fastest | 104.5 | 12.7 | 253 | 5,412 | yes (100,000 rows x 10 cells) |
| SheetStream-Shared | 125.0 | 33.1 | 692 | 3,648 | yes (100,000 rows x 10 cells) |
| SheetStream-Shared-Fastest | 127.4 | 35.5 | 503 | 6,143 | yes (100,000 rows x 10 cells) |
| SpreadCheetah | 103.3 | 11.5 | 322 | 5,393 | yes (100,000 rows x 10 cells) |
| SpreadCheetah-Optimal | 101.2 | 9.3 | 397 | 3,122 | yes (100,000 rows x 10 cells) |
| EPPlus 4.5.3.3 | 396.3 | 304.4 | 3,052 | 5,711 | yes (100,000 rows x 10 cells) |
| ClosedXML | 469.1 | 377.3 | 4,290 | 5,888 | yes (100,000 rows x 10 cells) |
| MiniExcel | 193.8 | 101.9 | 1,002 | 8,937 | yes (100,000 rows x 10 cells) |
| NPOI XSSF | 840.3 | 748.4 | 7,243 | 5,609 | yes (100,000 rows x 10 cells) |
| NPOI SXSSF | 209.6 | 117.8 | 3,355 | 5,807 | yes (100,000 rows x 10 cells) |

**Peak working set: StyledDates (3 fresh processes each, median shown)**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 34.0 | - | - | - | - |
| SheetStream | 46.9 | 12.9 | 480 | 3,320 | yes (100,001 rows x 8 cells) |
| SheetStream-Fastest | 49.3 | 15.4 | 282 | 5,840 | yes (100,001 rows x 8 cells) |
| SheetStream-Shared | 46.9 | 13.0 | 471 | 3,320 | yes (100,001 rows x 8 cells) |
| SheetStream-Shared-Fastest | 49.3 | 15.4 | 281 | 5,840 | yes (100,001 rows x 8 cells) |
| SpreadCheetah | 49.3 | 15.3 | 340 | 5,780 | yes (100,001 rows x 8 cells) |
| SpreadCheetah-Optimal | 46.9 | 12.9 | 532 | 3,281 | yes (100,001 rows x 8 cells) |
| EPPlus 4.5.3.3 | 308.7 | 274.8 | 3,152 | 5,221 | yes (100,001 rows x 8 cells) |
| ClosedXML | 350.3 | 316.4 | 4,658 | 5,567 | yes (100,001 rows x 8 cells) |
| MiniExcel | 147.8 | 113.8 | 777 | 9,205 | yes (100,001 rows x 8 cells) |
| NPOI XSSF | 569.6 | 535.6 | 5,153 | 5,065 | yes (100,001 rows x 8 cells) |
| NPOI SXSSF | 165.2 | 131.3 | 3,178 | 4,965 | yes (100,001 rows x 8 cells) |

**Peak working set: Million (3 fresh processes each, median shown)**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 118.3 | - | - | - | - |
| SheetStream | 158.9 | 40.6 | 2,683 | 29,220 | yes (1,000,001 rows x 10 cells) |
| SheetStream-Fastest | 181.7 | 63.5 | 995 | 53,015 | yes (1,000,001 rows x 10 cells) |
| SheetStream-Shared | 158.9 | 40.7 | 3,383 | 29,386 | yes (1,000,001 rows x 10 cells) |
| SheetStream-Shared-Fastest | 184.8 | 66.5 | 1,217 | 56,135 | yes (1,000,001 rows x 10 cells) |
| SpreadCheetah | 180.1 | 61.8 | 1,408 | 52,962 | yes (1,000,001 rows x 10 cells) |
| SpreadCheetah-Optimal | 157.1 | 38.8 | 3,100 | 29,220 | yes (1,000,001 rows x 10 cells) |
| EPPlus 4.5.3.3 | 1540.0 | 1421.7 | 23,495 | 56,226 | yes (1,000,001 rows x 10 cells) |
| ClosedXML | 2720.8 | 2602.6 | 35,765 | 59,796 | yes (1,000,001 rows x 10 cells) |
| MiniExcel | 327.4 | 209.2 | 4,093 | 104,941 | yes (1,000,001 rows x 10 cells) |
| NPOI XSSF | 5201.7 | 5083.4 | 47,221 | 53,480 | yes (1,000,001 rows x 10 cells) |
| NPOI SXSSF | 300.0 | 181.7 | 17,363 | 55,180 | yes (1,000,001 rows x 10 cells) |

**Peak working set: VeryWide (3 fresh processes each, median shown)**

| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |
|---|---:|---:|---:|---:|---|
| Baseline | 101.5 | - | - | - | - |
| SheetStream | 131.7 | 30.1 | 2,723 | 21,275 | yes (1,000 rows x 10,000 cells) |
| SheetStream-Fastest | 146.6 | 45.0 | 738 | 36,719 | yes (1,000 rows x 10,000 cells) |
| SheetStream-Shared | 131.6 | 30.0 | 2,727 | 21,275 | yes (1,000 rows x 10,000 cells) |
| SheetStream-Shared-Fastest | 146.7 | 45.2 | 734 | 36,719 | yes (1,000 rows x 10,000 cells) |
| SpreadCheetah | 146.3 | 44.8 | 1,297 | 36,719 | yes (1,000 rows x 10,000 cells) |
| SpreadCheetah-Optimal | 131.1 | 29.6 | 3,246 | 21,274 | yes (1,000 rows x 10,000 cells) |
| EPPlus 4.5.3.3 | 1715.2 | 1613.7 | 18,407 | 49,690 | yes (1,000 rows x 10,000 cells) |
| ClosedXML | 1626.7 | 1525.2 | 17,379 | 48,326 | yes (1,000 rows x 10,000 cells) |
| MiniExcel | 305.2 | 203.7 | 5,612 | 73,448 | yes (1,000 rows x 10,000 cells) |
| NPOI XSSF | 4871.6 | 4770.1 | 557,574 | 47,748 | yes (1,000 rows x 10,000 cells) |
| NPOI SXSSF | 508.4 | 406.8 | 18,931 | 47,753 | yes (1,000 rows x 10,000 cells) |

**Peak working set: Read (3 fresh processes each, median shown)**

| Library | Peak working set (MB) | Above baseline (MB) | Cold read time (ms) | Input size (KB) | Full-size checksum verified |
|---|---:|---:|---:|---:|---|
| Baseline | 57.8 | - | - | 6,026 | - |
| SheetStream-Rows | 135.6 | 77.8 | 1,154 | 6,026 | yes (checksum 5,001,050,000) |
| SheetStream-Typed | 136.4 | 78.6 | 1,278 | 6,026 | yes (checksum 5,001,050,000) |
| MiniExcel-Typed | 154.5 | 96.7 | 2,563 | 6,026 | yes (checksum 5,001,050,000) |
| MiniExcel-Dynamic | 155.5 | 97.7 | 2,522 | 6,026 | yes (checksum 5,001,050,000) |
| ExcelDataReader | 149.5 | 91.7 | 1,822 | 6,026 | yes (checksum 5,001,050,000) |
| ClosedXML | 253.0 | 195.2 | 4,651 | 6,026 | yes (checksum 5,001,050,000) |
| EPPlus 4.5.3.3 | 298.0 | 240.2 | 2,090 | 6,026 | yes (checksum 5,001,050,000) |
| NPOI XSSF | 1112.2 | 1054.5 | 5,606 | 6,026 | yes (checksum 5,001,050,000) |
| NPOI SAX | 245.7 | 188.0 | 3,571 | 6,026 | yes (checksum 5,001,050,000) |
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

### Start-up writes in two SheetStream (Fastest zip) rows

The SheetStream (Fastest zip) rows of **SpreadCheetahOwn** and **StringHeavy** have a StdDev of 39.5 and 33.2 ms on a write of 55 to 80 ms, and their means (81.8 and 92.0 ms) are well above their medians (56.3 and 78.6 ms). The per-iteration times show why (`per-iteration-fastest-rows.txt` in the published results):

* **SpreadCheetahOwn:** in every launch, SheetStream's first three measured writes took 128 to 150 ms, the next one or two took 56 to 78 ms, and from the sixth write on every write took 54 to 56 ms. SpreadCheetah's launches show the same kind of start-up in a milder form (its first two writes took 78 and 59 ms, 259 and 194 ms, and 82 and 65 ms). In the last five writes of each launch, SheetStream took 54.3 to 56.3 ms and SpreadCheetah 54.0 to 55.9 ms.
* **StringHeavy:** the first and third launches were steady from the first write (77 to 83 ms). In the second, four writes took 163 to 188 ms before the writer settled at 78 ms.

So this is a start-up effect, not a different steady state. We have not established its cause (JIT tiering is the likely one, but that is unverified), and these two rows are the only ones with it. A single cold write in a fresh process is not slower: the `--peak` tables above show 221 ms for SheetStream (Fastest zip) against 279 ms for SpreadCheetah on SpreadCheetahOwn, and 322 against 365 ms at the Optimal level. For these two rows we compare **medians** (SheetStream 56.3 against SpreadCheetah 55.0 ms on SpreadCheetahOwn; 78.6 against 79.7 ms on StringHeavy).

### The same comparison in three full runs

Three complete runs on the same machine, the two earlier ones with one launch per row (`3b6a8af`) and the last one as above. Each cell is SheetStream's time divided by SpreadCheetah's at the same zip level: **below 1.00 means SheetStream is faster**. A `*` marks a row whose StdDev is above 5% of its mean (the start-up effect above); there the median is used.

| Scenario | Zip level | Run A | Run B | Run C |
|---|---|---:|---:|---:|
| Long | Fastest zip | 0.74 | 0.75 | 0.75 |
| Long | Optimal zip | 0.90 | 0.91 | 0.91 |
| Wide | Fastest zip | 0.60 | 0.59 | 0.58 |
| Wide | Optimal zip | 0.84 | 0.86 | 0.86 |
| SpreadCheetahOwn | Fastest zip | 1.07* | 1.01* | 1.02* |
| SpreadCheetahOwn | Optimal zip | 1.03 | 1.02 | 1.02 |
| StringHeavy | Fastest zip | 0.98 | 1.07* | 0.99* |
| StringHeavy | Optimal zip | 1.00 | 0.99 | 1.00 |
| StyledDates | Fastest zip | 0.85 | 0.85 | 0.86 |
| StyledDates | Optimal zip | 0.96 | 0.96 | 0.97 |
| Million | Fastest zip | 0.75 | 0.75 | 0.77 |
| Million | Optimal zip | 0.90 | 0.90 | 0.87 |
| VeryWide | Fastest zip | 0.57 | 0.57 | 0.57 |
| VeryWide | Optimal zip | 0.85 | 0.86 | 0.85 |

The first two runs used one launch, which is why StringHeavy at Fastest zip reads 1.07 in run B: that launch happened to hit the start-up effect. The 3-launch run C is the reference.

## Reading the results honestly

* **Against SpreadCheetah, at the same zip level (the fairness pairing, rule 2).**
  * **SheetStream is faster on five of the seven write scenarios, in all three runs.** At `Fastest`: Long 25% less time, Wide 40 to 42%, VeryWide 43%, Million 23 to 25%, StyledDates 14 to 15%. At `Optimal` the gap is smaller: Long 9 to 10%, Wide 14 to 16%, VeryWide 14 to 15%, Million 10 to 13%, StyledDates 3 to 4%.
  * **StringHeavy is level.** At `Optimal` the ratio is 0.99 to 1.00 in every run; at `Fastest` the median ratio is 0.99 (the start-up effect above affects the mean).
  * **SpreadCheetah is ahead on its own scenario, SpreadCheetahOwn (every value distinct), by about 2% at `Optimal`** (SheetStream 172.2 ms against 168.7 ms; the ratio is 1.02 to 1.03 in all three runs). At `Fastest`, steady-state writes are level and the median is 2% behind. We have not closed this gap. An experiment that skipped the looks `StringStorage.Auto` takes at unique columns closed 1 to 3 ms of it, but it kept unbounded text in memory and was rejected. `StringStorage.Inline` (no looks, no de-duplication) did beat SpreadCheetah on this shape in our paired in-process test, but it is not the default because it gives up de-duplication for the columns that repeat.
  * **We do not claim "fastest on every shape".** The claim these numbers support: faster than SpreadCheetah on the long, wide, very wide, million-row and styled shapes, level on mixed text, and within about 2% on all-unique text.
* **Against MiniExcel.** SheetStream at its default settings is faster in every scenario except StyledDates: Long 248 against 336 ms, Wide 264 against 332 ms, SpreadCheetahOwn 172 against 354 ms, StringHeavy 196 against 417 ms, Million 2.5 against 3.5 s, VeryWide 2.6 against 5.0 s. On **StyledDates MiniExcel is faster than SheetStream's default** (301 against 341 ms); SheetStream at `Fastest` takes 105 ms. MiniExcel's files are larger (Long 10,420 KB against 2,917 KB); its zip level was not inspected.
* **Against EPPlus 4.5.3.3, ClosedXML and NPOI.** On the 100,000-row and 1,000 x 1,000 scenarios they take 1.4 to 8.1 s, against 0.06 to 0.34 s for SheetStream. At 10 million cells they take 15 to 46 s, and NPOI XSSF needs 526 s on VeryWide.
* **Reading.** SheetStream reads the 100,000 x 10 file in 351 ms (382 ms mapped to objects), against 693 ms for ExcelDataReader, 1.16 s for EPPlus, 1.4 to 1.5 s for MiniExcel, 2.0 s for NPOI SAX, 2.7 s for ClosedXML and 3.5 s for NPOI XSSF. This job uses one launch.
* **Output size.** Within about 1% of SpreadCheetah at the same level, and within 0.1% on Long, Wide and Million: for example Long 2,917 against 2,916 KB at `Optimal`. Where text is written inline SheetStream is slightly larger: SpreadCheetahOwn +1.2%, StringHeavy +0.8%, StyledDates +1.2% at `Optimal`.
* **Allocation.** SpreadCheetah allocates less on every shape except VeryWide: 11 to 34 KB per write, against SheetStream's 144 to 164 KB on Long, Wide, StyledDates and Million, 729 KB on StringHeavy and 1,336 KB on SpreadCheetahOwn (the looks `StringStorage.Auto` takes at unique columns). SheetStream's figure does not grow with the number of rows (164 KB on Long and Million alike). On VeryWide SpreadCheetah allocates 245 KB against SheetStream's 144 KB. EPPlus, ClosedXML and NPOI allocate 0.4 to 2.5 GB on the 100,000-row scenarios.
* **Peak memory.** SheetStream and SpreadCheetah are within 2 MB of each other at the same zip level, above the baseline. At `Optimal`: Long 14.4 against 12.8 MB, Wide 11.3 against 11.5 MB, Million 40.6 against 38.8 MB. At `Fastest`: 16.7 against 15.0 MB, 12.2 against 12.9 MB, 63.5 against 61.8 MB. On the 100,000-row scenarios MiniExcel needs 94 to 115 MB; on Long, Wide and SpreadCheetahOwn EPPlus needs 277 to 299 MB, ClosedXML 236 to 428 MB and NPOI XSSF 548 to 967 MB.
* **Speed gap from compression.** Most of the difference between the `Fastest` and `Optimal` rows is zip compression: `Optimal` costs about three times the time for a file 45% smaller. Compare libraries at the same level, as the pairing does.

## Caveats

* **One machine.** These are the numbers of one Azure virtual machine. The relative results (the ratios above) were the same in all three runs; the absolute times will differ on your hardware, so rerun the suite there before relying on a close result.
* **Cells without `r`.** SheetStream by default omits the `r` (cell reference) attribute on a cell that directly follows the previous one in its row. That is valid under ECMA-376, and Excel, SheetStream's own reader, ExcelDataReader, EPPlus, ClosedXML and NPOI XSSF read it correctly, but **NPOI 2.7.6's SAX reader, the Excel OleDb ACE 12.0 and 16.0 providers and MiniExcel 1.46.0 misread such files**. `PackageOptions.WriteCellReferences = true` writes every address and removes the speed this saves. The benchmark uses the default.
* The read benchmarks read a file that SheetStream wrote. Every reader gets the same bytes, but a file from Excel or another writer could lay out shared strings, styles or cell references differently, and a reader tuned for those files might rank differently. The 10-column file is also small enough to fit in memory many times over.
* The read checksum needs the `Id` value and a null check per cell. For the typed readers (`Read<T>`, `Query<T>`) that includes the mapping cost of all 10 properties, which is the price of using the typed API.
* The output stays in memory. Writing to a `FileStream` adds the same disk cost to every library, except NPOI SXSSF, which already does disk I/O for its temp file.
* Million and VeryWide use a lighter job (1 launch, 1 warm-up, 3 measured iterations), and Read uses one launch, so differences of a few percent there are noise.
* MiniExcel's Wide path boxes every double into a per-row dictionary. That is MiniExcel's documented shape for dynamic columns, but it explains much of its Wide allocation. An `IDataReader` source might allocate less; it was not measured.
* Every `GlobalSetup` runs the small correctness check for **all** libraries, so each benchmark process has every library's assemblies loaded and JIT-warmed. This affects neither the timed region nor the allocation figures.
* The `--peak` tables are one cold write each (including JIT), so they mix start-up cost into the comparison; use them for memory and output size, and the BenchmarkDotNet tables for speed.

