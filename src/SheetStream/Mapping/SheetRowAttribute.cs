using System;

namespace SheetStream
{
    /// <summary>
    /// Marks a <c>partial</c> class or struct whose column mapping is generated at compile time by the
    /// SheetStream source generator (shipped inside the SheetStream package). The generator adds a static
    /// <c>SheetMapper</c> property implementing <see cref="ISheetRowMapper{T}"/> and, on .NET 5 and later,
    /// registers it automatically so <c>LoadFromCollection</c> and <c>ExcelReader.Read&lt;T&gt;</c> need no
    /// reflection: the type then works when trimmed or published with Native AOT.
    /// </summary>
    /// <example>
    /// <code>
    /// [SheetRow]
    /// public partial class Order
    /// {
    ///     [SheetColumn("Order #", Order = 0)] public int Id { get; set; }
    ///     [SheetColumn(NumberFormat = "#,##0.00")] public decimal Amount { get; set; }
    /// }
    /// </code>
    /// </example>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct, Inherited = false, AllowMultiple = false)]
    public sealed class SheetRowAttribute : Attribute
    {
    }
}
