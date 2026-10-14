using System.ComponentModel;

namespace SheetStream.Generators.Tests;

public enum Status { Draft, Sent, Closed }

/// <summary>
/// Shape shared by <see cref="Order"/> (generated mapper) and <see cref="OrderPlain"/> (reflection). The two
/// types are kept member-for-member identical so a workbook written through each path can be compared.
/// </summary>
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

    public bool Paid { get; set; }

    public Status State { get; set; }

    [SheetColumn(Ignore = true)]
    public string Secret { get; set; } = "hidden";

    public double? Discount { get; set; }

    /// <summary>A public field: reflection sorts fields before properties of the same Order; the generator must too.</summary>
    [SheetColumn(Order = 5)]
    public string Reference = "";

    /// <summary>Init-only: written to the sheet but never assigned when reading through the generated mapper.</summary>
    public Guid Key { get; init; }

    public TimeSpan Duration { get; set; }

    public long Big { get; set; }

    public short Small { get; set; }

    public float Ratio { get; set; }

    public DateTimeOffset? Shipped { get; set; }

    [SheetColumn(Order = 1)]
    public byte Flags;
}

/// <summary>Reflection twin of <see cref="Order"/>: no [SheetRow], so the registry never has a mapper for it.</summary>
public sealed class OrderPlain
{
    [SheetColumn("Order #", Order = 0)]
    public int Id { get; set; }

    [DisplayName("Customer name")]
    public string? Customer { get; set; }

    [SheetColumn(NumberFormat = "#,##0.00")]
    public decimal Amount { get; set; }

    public DateTime Placed { get; set; }

    public bool Paid { get; set; }

    public Status State { get; set; }

    [SheetColumn(Ignore = true)]
    public string Secret { get; set; } = "hidden";

    public double? Discount { get; set; }

    [SheetColumn(Order = 5)]
    public string Reference = "";

    public Guid Key { get; init; }

    public TimeSpan Duration { get; set; }

    public long Big { get; set; }

    public short Small { get; set; }

    public float Ratio { get; set; }

    public DateTimeOffset? Shipped { get; set; }

    [SheetColumn(Order = 1)]
    public byte Flags;
}

/// <summary>A struct nested two levels deep; every container is partial.</summary>
public static partial class Geometry
{
    public partial class Shapes
    {
        [SheetRow]
        public partial struct Point
        {
            public int X;
            public int Y;

            [SheetColumn("Label")]
            public string? Name { get; set; }
        }
    }
}

/// <summary>A positional record: init-only properties and no parameterless constructor, so it is write-only.</summary>
[SheetRow]
public partial record Line(int Qty, string Sku);

/// <summary>Private nested type: the mapper is generated but cannot be registered by a module initializer (SSG005).</summary>
public partial class Holder
{
    [SheetRow]
    private partial class Hidden
    {
        public int Value { get; set; }
    }

    public static bool HiddenMapperWorks()
    {
        var mapper = Hidden.SheetMapper;
        return mapper.Headers.Count == 1 && mapper.Headers[0] == "Value" && !SheetMappers.TryGet<Hidden>(out _);
    }
}
