# SheetStream benchmarks

A reproducible comparison of SheetStream with the other .xlsx writers .NET developers actually use. The suite measures speed, bytes allocated, peak memory and output size. Before anything is timed, it checks that every library really wrote the rows.

| Library | Version | Licence | Why this version |
|---|---|---|---|
| SheetStream | this repo (ProjectReference) | MIT | |
| EPPlus | **4.5.3.3**, pinned `[4.5.3.3]` | LGPL | The last LGPL release. EPPlus 5 and later need a commercial licence for commercial use. |
| ClosedXML | 0.105.1 | MIT | Latest stable. |
| MiniExcel | 1.46.0 | Apache-2.0 | Latest stable. 2.0 is still in preview. |
| NPOI | **2.7.6**, pinned `[2.7.6]` | Apache-2.0 | The last release whose NuGet binaries are plain Apache-2.0. From 2.8.0 the package ships under an "OSMF EULA" with `requireLicenseAcceptance`: users above US$10,000 a year in revenue owe the project a maintenance fee for the binaries. The agreement says this fee is not a licence fee. 2.8.x still builds: it only prints a build warning until you set `AcceptNPOIOSMFLicense`. As with EPPlus 5+, this suite does not accept that agreement on anyone's behalf. |
| DocumentFormat.OpenXml | 3.5.1 | MIT | Used only by the correctness guard, to read every library's output back. |
| BenchmarkDotNet | 0.15.8 | MIT | |

NPOI is measured twice: once with **XSSF**, the in-memory user model, and once with **SXSSF**, the streaming port of Apache POI's SXSSF.

## How to run

All commands run from `benchmarks/SheetStream.Benchmarks`, and everything is Release.

```powershell
# Every benchmark (about 10 minutes on the machine below)
dotnet run -c Release -- --filter '*'

# One scenario
dotnet run -c Release -- --filter '*Wide*'

# Peak working set, output size and full-size verification, each library in a fresh process (about 3 minutes)
dotnet run -c Release -- --peak

# Correctness guard only: shows which rows, cells and formats every library actually wrote
dotnet run -c Release -- --verify
```

In Git Bash, quote the filter (`'*'`). Otherwise the shell expands `*` to file names.

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

## Methodology

* **One implementation per library** (`Writers.cs`). The benchmark, the correctness guard and the peak mode all call the same `WriteLong` / `WriteWide` method. What is verified is exactly what is timed.
* **Output goes to a reused `MemoryStream`** (`ReusableOutput`) for every library. It is created once with 64 MB of capacity and emptied before each write, so the buffer never grows while timed. As a result, "Allocated" counts only the library's own allocations. Its `Dispose` is a no-op, so a library that closes its output stream does no harm.
* **Each library is written the way its docs recommend:**
  * **SheetStream:** `new ExcelPackage(stream, leaveOpen: true)`, `Row(n).WriteCell(value, style)` with three styles created once, and default options (`CompressionLevel.Optimal`). A second row, **SheetStream (Fastest zip)**, sets only `PackageOptions.CompressionLevel = Fastest`. It is labelled separately because every other library runs at its default compression.
  * **EPPlus 4.5.3.3:** `Worksheet.SetValue(row, col, value)` in a loop. This avoids the `ExcelRange` that `Cells[r, c]` allocates on every call. Number formats are then set once per column range, and the package is written with `SaveAs(stream)`.
  * **ClosedXML:** `ws.Cell(r, c).Value = x` in a loop. `XLCellValue` is a struct, so this does not box. Formats are set once per column range, then `SaveAs(stream)`.
  * **MiniExcel:** `stream.SaveAs(...)`.
    * Long passes the typed `LongRow[]` with `[ExcelColumn(Format = ...)]` on the formatted columns.
    * Wide passes one `Dictionary<string, object>` per row, MiniExcel's documented shape for dynamic columns. This boxes every double.
    * Everything else is left at MiniExcel's defaults.
  * **NPOI:** `CreateRow` / `CreateCell` / `SetCellValue`, with the three `ICellStyle`s created once and `Write(stream, leaveOpen: true)`. SXSSF uses its default 100-row window. It flushes older rows to a temporary file on disk, so its time includes that disk I/O.
* **BenchmarkDotNet job:** `[MemoryDiagnoser]` with `SimpleJob(launchCount: 1, warmupCount: 3, iterationCount: 10)` and one invocation per iteration. `[IterationSetup]` empties the buffer, outside the measurement. Each benchmark runs in its own process, as BenchmarkDotNet does by default, on the default runtime settings (workstation GC, tiered compilation).
* **Correctness guard:** every benchmark's `GlobalSetup` writes a small case (50 rows for Long, 40 x 40 for Wide) with every library. It reads each output back with the Open XML SDK's SAX reader and throws if the row count or the number of cells in the last row is wrong. The `--peak` children repeat the check on the full-size output after their measurement.
* **Peak working set (`--peak`):** the executable starts itself again with `--peak-child <library> <scenario>`. Each child does the following:
  1. Builds the same input.
  2. Writes once, cold, including JIT.
  3. Reads `Process.PeakWorkingSet64`.
  4. Only then verifies the full-size output.

  Each case runs in 3 fresh processes and the median is reported. A **Baseline** child builds the input and writes nothing. It shows the floor every library sits on: the .NET runtime, the benchmark executable and the input data. The Baseline child never calls a writer, so the libraries' assemblies are not loaded or JIT-compiled in it. Each library's "above baseline" figure therefore includes loading and compiling its own code. The output buffer's pages count only once they are written, so each library's figure includes roughly its output size.

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

* These are write benchmarks only; reading is not measured.
* The output stays in memory. Writing to a `FileStream` adds the same disk cost to every library, except NPOI SXSSF, which already does disk I/O for its temp file.
* MiniExcel's Wide path boxes every double into a per-row dictionary. That is MiniExcel's documented shape for dynamic columns, but it explains much of its Wide allocation. An `IDataReader` source might allocate less; it was not measured.
* Every `GlobalSetup` runs the small correctness check for **all** libraries, so each benchmark process has every library's assemblies loaded and JIT-warmed. This affects neither the timed region nor the allocation figures.
