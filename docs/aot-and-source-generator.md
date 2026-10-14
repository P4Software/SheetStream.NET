# Trimming, Native AOT and the source generator

SheetStream works in trimmed and Native AOT applications. The only parts of the library that ever used
reflection are the object mappers behind `LoadFromCollection<T>`, `LoadFromAsyncEnumerable<T>` and
`ExcelReader.Read<T>`, and those now have a compile-time alternative: a Roslyn source generator that
ships inside the SheetStream package and writes the column mapping for you.

## Why

Reflection-based mapping has three costs that matter in modern deployments:

- **Trimming.** The IL trimmer cannot see which properties `GetProperties()` will return at run time, so it
  either keeps everything (defeating the point) or removes members the mapper needed.
- **Native AOT.** `Expression.Compile()` has no JIT to compile to; it falls back to an interpreter that is
  slower than the code it replaced.
- **Start-up.** Discovering members and compiling getters happens on the first call, per type.

A generated mapper has none of these: it is ordinary C# that reads `item.Amount` directly, calls the typed
`RowWriter.WriteCell(decimal, CellStyle?)` overload and never boxes a cell value.

## Using `[SheetRow]`

Mark the row type `partial` and add `[SheetRow]`. Everything else stays as it was: `[SheetColumn]`
(`Name`, `Order`, `NumberFormat`, `Ignore`) and `[DisplayName]` are honoured exactly like the reflection
path, for public properties and public fields alike.

```csharp
using SheetStream;

[SheetRow]
public partial class Order
{
    [SheetColumn("Order #", Order = 0)]
    public int Id { get; set; }

    [DisplayName("Customer name")]
    public string? Customer { get; set; }

    [SheetColumn(NumberFormat = "#,##0.00")]
    public decimal Amount { get; set; }

    public DateTime Placed { get; set; }

    [SheetColumn(Ignore = true)]
    public string InternalNote { get; set; } = "";
}
```

The generator adds two members to the type:

```csharp
public static ISheetRowMapper<Order> SheetMapper { get; }   // the mapping
public static void RegisterSheetMapper();                    // registers it in SheetMappers
```

On .NET 5 and later `RegisterSheetMapper()` carries `[ModuleInitializer]`, so it runs when the assembly
loads and the existing calls pick the mapper up automatically:

```csharp
sheet.LoadFromCollection(orders);                 // uses Order.SheetMapper, no reflection
var back = reader.Read<Order>("Orders").ToList(); // same
```

If you prefer not to depend on the registry, pass the mapper explicitly. These overloads contain no
reflection at all and are the ones to use when a trim or AOT analyzer is watching:

```csharp
sheet.LoadFromCollection(orders, Order.SheetMapper, new LoadOptions { TableName = "Orders" });
await sheet.LoadFromAsyncEnumerable(query.AsAsyncEnumerable(), Order.SheetMapper);
var back = reader.Read("Orders", Order.SheetMapper).ToList();
```

To see the generated file, expand *Dependencies > Analyzers > SheetStream.Generators* in Visual Studio, or
build with `-p:EmitCompilerGeneratedFiles=true` and look under `obj/.../generated`.

### Rules the generator follows

- Columns are ordered by `[SheetColumn(Order)]` ascending, then in declaration order. Within the same
  `Order`, public fields come before properties, because that is the order reflection's metadata tokens give
  and both paths must produce identical sheets.
- Inherited public members are included, base type first.
- `init`-only properties, get-only properties and `readonly` fields are written but never assigned when
  reading (`CanSet` is false for them). The reflection path can assign `init` properties; the generator
  cannot, because C# does not allow it outside an object initializer.
- A type without a parameterless constructor (a positional record, for example) gets a write-only mapper:
  `CanRead` is false and `Read` throws `NotSupportedException`.
- Nested types are supported when every containing type is also `partial`. Generic types are not.

### Diagnostics

| Id | Severity | Meaning |
|----|----------|---------|
| SSG001 | Error | The `[SheetRow]` type is not `partial`. |
| SSG002 | Error | A type containing the `[SheetRow]` type is not `partial`. |
| SSG003 | Warning | The type has no public instance properties or fields to map; nothing was generated. |
| SSG004 | Error | The type is generic or nested in a generic type. |
| SSG005 | Info | The type is private or protected, so it cannot be registered at module load; use `SheetMapper` explicitly or call `RegisterSheetMapper()`. |

## Manual registration on .NET Framework

Module initializers need .NET 5 or later and C# 9. When the consuming project targets .NET Framework (or
an older language version) the generator still produces `SheetMapper` and `RegisterSheetMapper()`, but
nothing calls the latter for you. Call it once at start-up, before the first export or import:

```csharp
Order.RegisterSheetMapper();
Invoice.RegisterSheetMapper();
```

Or skip the registry altogether and pass `Order.SheetMapper` to the explicit overloads shown above.

## Writing a mapper by hand

`ISheetRowMapper<T>` is a public interface, so a mapper does not have to be generated. Implement it for a
type you cannot mark `partial` (one from another assembly, say) and register it with
`SheetMappers.Register(myMapper)` or pass it explicitly. The contract is documented on the interface;
the generated code is a good template.

## What still needs reflection

- `LoadFromCollection<T>` / `LoadFromAsyncEnumerable<T>` / `ExcelReader.Read<T>` **without** a registered
  mapper fall back to reflection over the public properties and fields of `T`. Those generic parameters are
  annotated with `[DynamicallyAccessedMembers(PublicProperties | PublicFields)]`, so the trimmer keeps the
  members and the trim analyzer warns in your code when `T` is not statically known.
- `IDictionary<string, object?>` rows, `DataTable` and `IDataReader` sources never used member reflection
  and are unaffected.
- The reflection fallback compiles getters with `Expression.Compile()`. On Native AOT this runs through
  the expression interpreter: correct, but slower than a generated mapper.

Everything else in the library (writing, styles, charts, templates, the reader) is plain code with no
reflection.

## Publishing

SheetStream's `net8.0` build is marked `IsTrimmable` and `IsAotCompatible`, and it compiles with the trim,
single-file and AOT analyzers enabled and warnings treated as errors, so the library itself produces no
`IL2xxx` or `IL3xxx` warnings when you publish with `PublishTrimmed=true` or `PublishAot=true`. Any warning
you see will point at your own code, typically a `LoadFromCollection<T>` call with a type that has no
`[SheetRow]`; add the attribute (or pass a mapper) and it goes away.
