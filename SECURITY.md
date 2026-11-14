# Security Policy

## Supported versions

Security fixes are released for the latest minor version of SheetStream and SheetStream.AspNetCore on NuGet.

## Reporting a vulnerability

Please do not open a public issue for security problems.

Report them privately through GitHub's private vulnerability reporting: https://github.com/P4Software/SheetStream.NET/security/advisories/new (the **Report a vulnerability** button on the repository's **Security** tab). Include the SheetStream version, a minimal reproduction (code or a sample file), and the impact you observed.

What happens next:

1. We acknowledge the report within 3 business days.
2. We confirm the issue, agree a severity with you, and work on a fix in a private fork.
3. For confirmed issues we publish a fixed version on NuGet within 30 days, together with a GitHub security advisory that credits the reporter (unless you ask us not to) and describes the affected versions and the mitigation.

Please give us that window before disclosing publicly. If you have not heard back within 3 business days, open a plain issue saying only that you sent a security report, without details, and we will respond there.

## Verifying what you install

Every release is built by the public `Release` workflow in this repository. Each package carries a signed build-provenance attestation (`gh attestation verify SheetStream.NET.<version>.nupkg --owner P4Software`), and the GitHub release includes an SPDX software bill of materials. The packages have no runtime dependencies.

## Scope notes

- **Opening untrusted files.** Every way of opening a workbook the library did not write is in scope: `ExcelReader`, `ExcelPackage.OpenTemplate` and `ExcelPackage.Edit`. DTD processing and external entity resolution are disabled. From 0.6.0 all three bound what a crafted file can make them hold in memory and how much work it can cause (`ReaderOptions`, and `PackageOptions.SourceLimits` for templates and edits), and end in `InvalidDataException`, `XmlException`, `FormatException` or `OverflowException` for a file they cannot process. Reports of crashes, hangs, other exception types or memory use beyond the configured limits on a crafted file are in scope; please include the options you used. Versions before 0.6.0 are known to be affected by memory exhaustion (see the changelog) and should not be used on untrusted files.
- **What is not a vulnerability.** The packed file as uploaded is not bounded by the library; cap the request size in your web framework. `ReadDataTable`, `Read<T>().ToList()` and similar hold every row because that is what they are asked to do. Memory use up to `ReaderOptions.MaxLoadedPartBytes` (256 MB by default) is within the limit, not past it; lower the limit if that is more than you want one file to cost.
- **Sheet and workbook protection** stores a salted SHA-512 password verifier, as Excel does. It deters accidental edits; it is not encryption, and files can be opened by tools that ignore it. That behaviour is by design and not a vulnerability.
