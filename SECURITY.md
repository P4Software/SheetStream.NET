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

- **Reading untrusted files.** `ExcelReader` and `ExcelPackage.OpenTemplate` parse files that may come from users. DTD processing and external entity resolution are disabled, cell references beyond Excel's limits are rejected, and the reader is fuzz-tested in CI. Reports of crashes, hangs or excessive memory use on crafted files are in scope.
- **Sheet and workbook protection** stores a salted SHA-512 password verifier, as Excel does. It deters accidental edits; it is not encryption, and files can be opened by tools that ignore it. That behaviour is by design and not a vulnerability.
