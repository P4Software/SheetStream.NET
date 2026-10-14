using MiniExcelLibs.Attributes;

namespace SheetStream.Benchmarks;

/// <summary>
/// One row of the "Long" scenario: ten columns of mixed types. A class with properties (not a struct or
/// tuple) because MiniExcel's idiomatic typed path reflects over public properties.
/// </summary>
/// <remarks>
/// The <see cref="ExcelColumnAttribute"/> formats are read only by MiniExcel; every other library applies
/// the same two number formats through its own style API.
/// </remarks>
public sealed class LongRow
{
    /// <summary>Sequential identifier (int).</summary>
    public int Id { get; set; }

    /// <summary>Text drawn from a small repeated set, so shared-string de-duplication matters.</summary>
    public string Name { get; set; } = "";

    /// <summary>Second repeated text column.</summary>
    public string Category { get; set; } = "";

    /// <summary>Integer quantity.</summary>
    public int Quantity { get; set; }

    /// <summary>Double shown with <see cref="Data.NumberFormat"/>.</summary>
    [ExcelColumn(Format = Data.NumberFormat)]
    public double Price { get; set; }

    /// <summary>Double shown with <see cref="Data.NumberFormat"/>.</summary>
    [ExcelColumn(Format = Data.NumberFormat)]
    public double Amount { get; set; }

    /// <summary>Date (no time part) shown with <see cref="Data.DateFormat"/>.</summary>
    [ExcelColumn(Format = Data.DateFormat)]
    public DateTime Date { get; set; }

    /// <summary>Third repeated text column.</summary>
    public string Region { get; set; } = "";

    /// <summary>Double shown with <see cref="Data.NumberFormat"/>.</summary>
    [ExcelColumn(Format = Data.NumberFormat)]
    public double Score { get; set; }

    /// <summary>Second integer column.</summary>
    public int Code { get; set; }
}

/// <summary>
/// Deterministic input data shared by every library. It is built once, outside the timed region, so each
/// library pays exactly the same (zero) cost for producing its input and the same memory floor for holding it.
/// </summary>
public static class Data
{
    /// <summary>Rows in the Long scenario.</summary>
    public const int LongRows = 100_000;

    /// <summary>Rows in the Wide scenario.</summary>
    public const int WideRows = 1_000;

    /// <summary>Columns in the Wide scenario.</summary>
    public const int WideColumns = 1_000;

    /// <summary>Excel number format applied to the three double columns of the Long scenario.</summary>
    public const string NumberFormat = "#,##0.00";

    /// <summary>Excel number format applied to the date column of the Long scenario.</summary>
    public const string DateFormat = "yyyy-mm-dd";

    /// <summary>Header captions of the Long scenario, in column order (they match the <see cref="LongRow"/> property names).</summary>
    public static readonly string[] LongHeaders =
        { "Id", "Name", "Category", "Quantity", "Price", "Amount", "Date", "Region", "Score", "Code" };

    private static readonly string[] Names = { "Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel" };
    private static readonly string[] Categories = { "Hardware", "Software", "Services", "Support", "Training" };
    private static readonly string[] Regions = { "North", "South", "East", "West" };

    /// <summary>Builds <paramref name="count"/> Long rows. The same index always yields the same values.</summary>
    public static LongRow[] CreateLong(int count)
    {
        var rows = new LongRow[count];
        var epoch = new DateTime(2020, 1, 1);
        for (int i = 0; i < count; i++)
        {
            int qty = (i * 7) % 1000;
            double price = (i % 997) * 1.25 + 0.99;
            rows[i] = new LongRow
            {
                Id = i + 1,
                Name = Names[i % Names.Length],
                Category = Categories[i % Categories.Length],
                Quantity = qty,
                Price = price,
                Amount = qty * price,
                Date = epoch.AddDays(i % 2000),
                Region = Regions[i % Regions.Length],
                Score = ((i * 31) % 10_000) / 100.0,
                Code = 100_000 + i,
            };
        }
        return rows;
    }

    /// <summary>Builds a <paramref name="rows"/> x <paramref name="columns"/> grid of doubles.</summary>
    public static double[,] CreateWide(int rows, int columns)
    {
        var grid = new double[rows, columns];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < columns; c++)
                grid[r, c] = (r * columns + c) * 0.5 + 0.25;
        return grid;
    }
}
