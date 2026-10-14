# Shipping the analyzer inside the SheetStream package

`SheetStream.Analyzers` is not a NuGet package of its own. The `SheetStream` package carries it under
`analyzers/dotnet/cs`, so every project that installs SheetStream gets the SS0001-SS0003 warnings with no extra
install step.

## Lines to add to `src/SheetStream/SheetStream.csproj`

Add this `ItemGroup` next to the existing one that packs `README.md` and `icon.png`:

```xml
  <ItemGroup>
    <!-- Design-time analyzer (SS0001-SS0003) shipped inside this package. ReferenceOutputAssembly=false only orders
         the build; PrivateAssets=all keeps it out of the package's dependency list. -->
    <ProjectReference Include="../SheetStream.Analyzers/SheetStream.Analyzers.csproj" ReferenceOutputAssembly="false" PrivateAssets="all" />
    <None Include="../SheetStream.Analyzers/bin/$(Configuration)/netstandard2.0/SheetStream.Analyzers.dll" Pack="true" PackagePath="analyzers/dotnet/cs" Visible="false" />
  </ItemGroup>
```

Why these two lines and not `OutputItemType="Analyzer"`: that attribute would also run the analyzer on the
SheetStream library itself, which is pointless. The plain `ProjectReference` with `ReferenceOutputAssembly="false"`
only guarantees the analyzer is built first; the `None` item copies the built dll into the package.

## What `dotnet pack -c Release` must produce

```
analyzers/dotnet/cs/SheetStream.Analyzers.dll
lib/net8.0/SheetStream.dll
lib/netstandard2.0/SheetStream.dll
```

and the nuspec `<dependencies>` groups must stay empty (the analyzer is not a dependency, it is content).

Verified on 2026-10-14 with the lines above: the package contained the three files, both dependency groups were
empty, and a throwaway console project restored from the local `.nupkg` reported SS0001, SS0002 and SS0003 at build
time.

## Also add

- `SheetStream.slnx`: `src/SheetStream.Analyzers/SheetStream.Analyzers.csproj` under `/src/` and
  `tests/SheetStream.Analyzers.Tests/SheetStream.Analyzers.Tests.csproj` under `/tests/`.
- `docs/toc.yml`: `- name: Analyzers` / `  href: analyzers.md`.
- When the package version that first carries the analyzer ships, move the three rules from
  `AnalyzerReleases.Unshipped.md` to `AnalyzerReleases.Shipped.md` under a `## Release <version>` heading.
