; Shipped analyzer releases
; https://github.com/dotnet/roslyn-analyzers/blob/main/src/Microsoft.CodeAnalysis.Analyzers/ReleaseTrackingAnalyzers.Help.md

## Release 0.4.0

### New Rules

Rule ID | Category | Severity | Notes
--------|----------|----------|-------
SSG001 | SheetStream | Error | A [SheetRow] type must be declared partial
SSG002 | SheetStream | Error | Types containing a [SheetRow] type must be declared partial
SSG003 | SheetStream | Warning | A [SheetRow] type has no mappable members
SSG004 | SheetStream | Error | Generic types cannot be [SheetRow] types
SSG005 | SheetStream | Info | The mapper cannot be registered automatically
