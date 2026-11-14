# Contributing to SheetStream.NET

Thank you for helping. Bug reports, fixes, tests and documentation are all welcome.

## Ground rules

1. Zero external dependencies. Library code uses only APIs available on `netstandard2.0` and `net8.0` (`System.IO.Compression`, `System.Xml` and so on). The test project may reference `DocumentFormat.OpenXml` to validate output; the library never does.
2. The hot path stays allocation-free. `RowWriter.WriteCell` must not allocate, and the writer must not build a whole-document model in memory.
3. Output follows ECMA-376. Every new element must be covered by the schema-validation test, so Excel, Google Sheets and LibreOffice open the file.
4. Public types and members carry XML doc comments, and non-obvious logic has a comment that explains why.
5. Anything that opens a workbook this library did not write is bounded in memory and in time, and fails with `InvalidDataException`, `XmlException`, `FormatException` or `OverflowException` only. That covers `ExcelReader`, `ExcelPackage.Edit` and `ExcelPackage.OpenTemplate`. See "Code that reads a workbook" below.

## Code that reads a workbook

A workbook is a zip archive of XML, so a file of a few kilobytes can unpack to gigabytes, name a part a thousand times, or nest a formula a hundred thousand levels deep. Two releases shipped with holes of this kind, so the checks below are not optional.

- **Read parts through the guards.** `ExcelReader.Guard(...)` and `TemplateSource.OpenXml(...)` decode a part and enforce the limits on it. Do not create an `XmlReader` over an entry stream directly.
- **Charge what you keep.** Anything that outlives the read of one row goes against a `MemoryBudget`. A count limit is not enough: charge the bytes, including strings.
- **Mind the time.** Look for anything that costs more than the size of what it reads: a lookup inside a loop, a list shortened from the front, a string inspected once per cell. Cache per key.
- **Add your shape to the table.** `tests/Shared/HostileWorkbooks.cs` lists hostile files and every way of opening a workbook. A new way in goes into `Entries`; a new kind of hostile file into `Shapes`. Every shape then runs through every entry point, in-process and again in a process whose heap is capped at 64 MB (`tests/SheetStream.MemoryProbe`). The outcome of each pair is recorded in `HostileWorkbookOutcomes.txt`; after an intended change, run the tests with `SHEETSTREAM_RECORD_OUTCOMES=1` and review the difference.
- **Fuzz deep before a release.** The fuzz tests run a few hundred damaged files by default. Set `SHEETSTREAM_FUZZ_ITERATIONS=3000` and a few values of `SHEETSTREAM_FUZZ_SEED` for a run that takes minutes and finds what the quick one does not.
- **Compare with Excel.** `tests/corpus/excel` holds workbooks Excel wrote, with Excel's own account of every cell; `ExcelCorpusTests` compares the reader with it. `tests/corpus/build-excel-corpus.ps1` regenerates it (needs Excel). `tests/excel-check.ps1` opens workbooks in Excel and fails on any it has to repair; run it on the samples' output and on round-tripped corpus files when you change what is written.

## Reporting a bug or requesting a feature

Search the open issues first, then open a new one with the steps to reproduce it or the OpenXML use case you need. For a security problem, follow [SECURITY.md](SECURITY.md) instead of opening a public issue.

## Sending a pull request

1. Fork the repository and branch from `main`, for example `fix/sst-memory-leak`.
2. Add a test with every feature or fix. For output changes, the test confirms that `OpenXmlValidator` reports zero errors.
3. Run the suite before you push:

   ```bash
   dotnet build
   dotnet test
   dotnet test tests/SheetStream.NetFx.Tests   # .NET Framework 4.8 run of the netstandard2.0 build (Windows only)
   ```

4. Describe what changed and why in the pull request.

## Licensing

By contributing you agree that your work is licensed under the project's [MIT licence](LICENSE), and you confirm that you wrote it or have the right to submit it.
