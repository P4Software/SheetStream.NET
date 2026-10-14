# Security Policy

## Supported versions

Security fixes are released for the latest minor version of SheetStream and SheetStream.AspNetCore on NuGet.

## Reporting a vulnerability

Please do not open a public issue for security problems.

Report them privately through GitHub: open the repository's **Security** tab and choose **Report a vulnerability**. Include the SheetStream version, a minimal reproduction (code or a sample file), and the impact you observed.

We aim to acknowledge reports within 3 business days and to publish a fix or mitigation within 30 days for confirmed issues. We credit reporters in the release notes unless you ask us not to.

## Scope notes

- **Reading untrusted files.** `ExcelReader` and `ExcelPackage.OpenTemplate` parse files that may come from users. DTD processing and external entity resolution are disabled, cell references beyond Excel's limits are rejected, and the reader is fuzz-tested in CI. Reports of crashes, hangs or excessive memory use on crafted files are in scope.
- **Sheet and workbook protection** stores a salted SHA-512 password verifier, as Excel does. It deters accidental edits; it is not encryption, and files can be opened by tools that ignore it. That behaviour is by design and not a vulnerability.
