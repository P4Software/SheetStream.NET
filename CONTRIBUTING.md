# Contributing to SheetStream 🚀

First off, thank you for considering contributing to SheetStream! We are on a mission to build a completely free, lightning-fast, and open-source alternative to bloated commercial .NET Excel libraries—and we can't do it without the community.

By contributing, you help make software licensing fair and accessible for developers everywhere.

## 📐 Ground Rules & Principles

Before writing any code, please keep our core architectural pillars in mind:

1. **Zero External Dependencies:** Every line of library code must rely purely on native .NET APIs available on `netstandard2.0` and `net8.0` (`System.IO.Compression`, `System.Xml`, etc.). Do not pull in third-party packages. The test project may reference `DocumentFormat.OpenXml` to validate output; the library itself never does.
2. **Performance First:** SheetStream is built around low memory overhead and sequential streaming XML writers. Avoid loading full documents or heavy tracking tables into memory (`DOM` patterns). The hot path (`RowWriter.WriteCell`) must stay allocation-free.
3. **Strict Compliance:** All outputs must stringently follow the **ECMA-376 OpenXML standard** to guarantee compatibility with native Microsoft Excel, Google Sheets and LibreOffice. Every new element must be covered by the schema-validation test.
4. **Documented as built:** Public types and members carry XML doc comments; non-obvious logic explains *why* in a comment.

## 🛠️ How to Contribute

### 1. Reporting Bugs or Requesting Features

* Check the **Issues** tab to make sure your topic hasn't already been reported.
* Open a new issue using our structured templates, detailing reproduction steps or specific OpenXML implementation use cases.

### 2. Submitting a Pull Request (PR)

* Fork the repository and create your branch from `main` (e.g., `feature/add-borders` or `fix/sst-memory-leak`).
* Write clean, commented C# adhering to modern standard conventions (PascalCase for methods, camelCase for local variables).
* **Add Tests:** If you are adding a feature or fixing a bug, please include a test confirming that the generated Excel output remains valid (`OpenXmlValidator` reports zero errors) and doesn't get corrupted.
* Run the full suite before pushing:

  ```bash
  dotnet build
  dotnet test
  ```

* Submit your PR with a clear summary of what changed and why.

## 📜 Legal Note (DCO)

By contributing to this repository, you explicitly agree that your code will be licensed under the project's **MIT License**. You certify that you wrote the code yourself or have the legal right to submit it as open source.
