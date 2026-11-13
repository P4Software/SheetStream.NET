# Contributing to SheetStream.NET

Thank you for helping. Bug reports, fixes, tests and documentation are all welcome.

## Ground rules

1. Zero external dependencies. Library code uses only APIs available on `netstandard2.0` and `net8.0` (`System.IO.Compression`, `System.Xml` and so on). The test project may reference `DocumentFormat.OpenXml` to validate output; the library never does.
2. The hot path stays allocation-free. `RowWriter.WriteCell` must not allocate, and the writer must not build a whole-document model in memory.
3. Output follows ECMA-376. Every new element must be covered by the schema-validation test, so Excel, Google Sheets and LibreOffice open the file.
4. Public types and members carry XML doc comments, and non-obvious logic has a comment that explains why.

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
