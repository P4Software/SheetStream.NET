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

    /// <summary>Rows in the Million scenario (same columns as Long).</summary>
    public const int MillionRows = 1_000_000;

    /// <summary>Rows in the VeryWide scenario.</summary>
    public const int VeryWideRows = 1_000;

    /// <summary>Columns in the VeryWide scenario; with <see cref="VeryWideRows"/> that is the same 10M cells as Million.</summary>
    public const int VeryWideColumns = 10_000;

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

    /// <summary>Columns of the two string scenarios (<see cref="Scenarios.SpreadCheetahOwn"/>, <see cref="Scenarios.StringHeavy"/>).</summary>
    public const int StringColumns = 10;

    /// <summary>Header captions of the styled scenario, in column order (they match the <see cref="StyledRow"/> property names).</summary>
    public static readonly string[] StyledHeaders =
        { "OrderDate", "Created", "ShipDate", "Net", "Tax", "Total", "Margin", "Discount" };

    /// <summary>Excel format of the date columns of the styled scenario.</summary>
    public const string StyledDateFormat = "yyyy-mm-dd";

    /// <summary>Excel format of the date-time column of the styled scenario.</summary>
    public const string DateTimeFormat = "yyyy-mm-dd hh:mm:ss";

    /// <summary>Excel format of the money columns of the styled scenario (same code as <see cref="NumberFormat"/>).</summary>
    public const string MoneyFormat = "#,##0.00";

    /// <summary>Excel format of the percentage columns of the styled scenario.</summary>
    public const string PercentFormat = "0.00%";

    private static readonly string[] Statuses = { "New", "Pending", "Shipped", "Delivered", "Cancelled" };
    private static readonly string[] Priorities = { "Low", "Normal", "High" };

    private static readonly string[] Countries =
    {
        "Argentina", "Australia", "Austria", "Belgium", "Brazil", "Canada", "Chile", "Colombia", "Costa Rica",
        "Denmark", "Ecuador", "Finland", "France", "Germany", "Greece", "Ireland", "Italy", "Japan", "Mexico",
        "Netherlands", "New Zealand", "Norway", "Panama", "Peru", "Poland", "Portugal", "Spain", "Sweden",
        "Switzerland", "United Kingdom", "United States", "Uruguay",
    };

    private static readonly string[] ProductCategories =
    {
        "Hardware", "Software", "Services", "Support", "Training", "Licences", "Accessories", "Networking",
        "Storage", "Security", "Consulting", "Cloud",
    };

    private static readonly string[] Cities =
    {
        "Madrid", "Lisbon", "Paris", "Berlin", "Rome", "Vienna", "Oslo", "Stockholm", "Helsinki", "Dublin",
        "London", "Toronto", "Montreal", "Chicago", "Houston", "Denver", "Miami", "Seattle", "Boston", "Austin",
        "Mexico City", "Bogota", "Lima", "Quito", "Santiago", "Buenos Aires", "Montevideo", "Panama City",
        "Sao Paulo", "Rio de Janeiro", "Tokyo", "Osaka", "Sydney", "Melbourne", "Auckland", "Zurich", "Geneva",
        "Brussels", "Amsterdam", "Warsaw", "Athens", "Copenhagen", "San Jose", "San Diego", "Phoenix",
        "Dallas", "Atlanta", "Orlando", "Calgary", "Vancouver",
    };

    // Names include characters outside ASCII and an apostrophe, because real customer data has them.
    private static readonly string[] FirstNames =
    {
        "Ana", "Luis", "Maria", "Jose", "Carmen", "Pedro", "Laura", "Miguel", "Sofia", "Diego", "Elena", "Pablo",
        "Lucia", "Andres", "Marta", "Javier", "Paula", "Carlos", "Isabel", "Sergio", "Julia", "Daniel", "Clara",
        "Hugo", "Eva", "Mateo", "Noa", "Leo", "Zoe", "Max", "Jean-Luc", "Siobhan", "Bjorn", "Renee", "Zoltan",
        "Amelie", "Henrik", "Ines", "Tomas", "Anais", "Ole", "Greta", "Ravi", "Mei", "Kenji", "Aisha", "Omar", "Yuki",
    };

    private static readonly string[] LastNames =
    {
        "García", "Müller", "O'Brien", "Smith & Sons", "Fernández", "Rossi", "Dubois", "Nielsen", "Kowalski",
        "Silva", "Santos", "López", "Martin", "Schmidt", "Johnson", "Brown", "Tanaka", "Nguyen", "Haddad", "Costa",
        "Peters", "Andersson", "Virtanen", "Papadopoulos", "Murphy", "Walsh", "Rivera", "Torres", "Ruiz", "Vargas",
        "Moreno", "Castro", "Ortega", "Delgado", "Navarro", "Romero", "Ramos", "Gil", "Serrano", "Blanco",
    };

    // The first words of each note. Half of them contain a character that must be escaped in XML (& < > or "),
    // and they come first so truncating the note to 60 characters never cuts the escaped character off.
    private static readonly string[] NoteStarts =
    {
        "Qty < 5 & rush delivery", "Price > list, ask for \"discount\"", "Fragile & heavy: handle with care",
        "Standard order", "Call back before 5pm", "Customer prefers e-mail", "Returns & exchanges allowed",
        "Gift wrap requested", "See ticket <open>", "Invoice with next shipment", "Needs signature on receipt",
        "No special instructions",
    };

    // The last words of each note: plain text of different lengths, so the notes vary between about 30 and 60 characters.
    private static readonly string[] NoteTails =
    {
        "confirmed by phone", "waiting for the warehouse to confirm stock", "customer asked for a call",
        "reviewed", "pending approval from the finance team", "second attempt",
    };

    /// <summary>
    /// Builds the <see cref="Scenarios.SpreadCheetahOwn"/> input: <paramref name="rows"/> x ten string cells,
    /// every value distinct, no value repeated within a row or a column, so nothing can be de-duplicated. That is
    /// the shape SpreadCheetah's published benchmark uses (100,000 x 10 string cells, each built from its row and
    /// column position); the text itself is ours.
    /// </summary>
    public static string[][] CreateOwnStrings(int rows)
    {
        var grid = new string[rows][];
        for (int r = 0; r < rows; r++)
        {
            var row = new string[StringColumns];
            for (int c = 0; c < row.Length; c++) row[c] = "Value " + r + "-" + c;
            grid[r] = row;
        }
        return grid;
    }

    /// <summary>
    /// Builds the <see cref="Scenarios.StringHeavy"/> input: <paramref name="rows"/> x ten string cells.
    /// Columns: order id (unique), status (5 values), country (32), category (12), customer name (at most 1,920
    /// distinct combinations), e-mail (unique), city (50), priority (3), SKU (unique) and a free-text note of at
    /// most 60 characters (unique; every note starts with one of 12 phrases, five of which need XML escaping).
    /// </summary>
    public static string[][] CreateStringHeavy(int rows)
    {
        var grid = new string[rows][];
        for (int i = 0; i < rows; i++)
        {
            // The multiplier is odd, so the 32-bit product is a bijection of i: the SKU is unique for any row count.
            string sku = "SKU-" + unchecked((uint)(i * 2654435761L)).ToString("X8");
            // The running number comes before the tail, so cutting the note to 60 characters can only shorten
            // the tail and every note stays unique.
            string note = NoteStarts[(i * 5) % NoteStarts.Length] + " #" + (i + 1) + " - " + NoteTails[(i * 3) % NoteTails.Length];
            if (note.Length > 60) note = note.Substring(0, 60);
            grid[i] = new[]
            {
                "SO-" + (i + 1).ToString("D8"),
                Statuses[(i * 3) % Statuses.Length],
                Countries[(i * 7) % Countries.Length],
                ProductCategories[(i * 5) % ProductCategories.Length],
                FirstNames[i % FirstNames.Length] + " " + LastNames[(i / FirstNames.Length + i) % LastNames.Length],
                "customer" + (i + 1) + "@example.com",
                Cities[(i * 11) % Cities.Length],
                Priorities[(i / 3) % Priorities.Length],
                sku,
                note,
            };
        }
        return grid;
    }

    /// <summary>Builds <paramref name="count"/> rows of the <see cref="Scenarios.StyledDates"/> input. The same index always yields the same values.</summary>
    public static StyledRow[] CreateStyled(int count)
    {
        var rows = new StyledRow[count];
        var epoch = new DateTime(2020, 1, 1);
        for (int i = 0; i < count; i++)
        {
            var order = epoch.AddDays(i % 2000);
            // Whole seconds only: a double serial number then represents the time exactly enough for the guard to
            // compare it with a tolerance far below one second.
            double net = Math.Round(((i * 37) % 500_000) / 100.0 + 1, 2);
            double tax = Math.Round(net * 0.07, 2);
            rows[i] = new StyledRow
            {
                OrderDate = order,
                Created = order.AddSeconds((i * 37) % 86_400),
                ShipDate = order.AddDays(1 + i % 14),
                Net = net,
                Tax = tax,
                Total = Math.Round(net + tax, 2),
                Margin = ((i * 13) % 4000) / 10_000.0,
                Discount = ((i * 7) % 2500) / 10_000.0,
            };
        }
        return rows;
    }
}

/// <summary>
/// One row of the "StyledDates" scenario: eight columns that all carry a number format. A class with properties
/// because MiniExcel's idiomatic typed path reflects over public properties; the formats in the attributes are
/// read only by MiniExcel (every other library applies the same formats through its own style API).
/// </summary>
public sealed class StyledRow
{
    /// <summary>Date without a time part, <see cref="Data.StyledDateFormat"/>.</summary>
    [ExcelColumn(Format = Data.StyledDateFormat)]
    public DateTime OrderDate { get; set; }

    /// <summary>Date and time (whole seconds), <see cref="Data.DateTimeFormat"/>.</summary>
    [ExcelColumn(Format = Data.DateTimeFormat)]
    public DateTime Created { get; set; }

    /// <summary>Date without a time part, <see cref="Data.StyledDateFormat"/>.</summary>
    [ExcelColumn(Format = Data.StyledDateFormat)]
    public DateTime ShipDate { get; set; }

    /// <summary>Money, <see cref="Data.MoneyFormat"/>.</summary>
    [ExcelColumn(Format = Data.MoneyFormat)]
    public double Net { get; set; }

    /// <summary>Money, <see cref="Data.MoneyFormat"/>.</summary>
    [ExcelColumn(Format = Data.MoneyFormat)]
    public double Tax { get; set; }

    /// <summary>Money, <see cref="Data.MoneyFormat"/>.</summary>
    [ExcelColumn(Format = Data.MoneyFormat)]
    public double Total { get; set; }

    /// <summary>A fraction shown as a percentage, <see cref="Data.PercentFormat"/>.</summary>
    [ExcelColumn(Format = Data.PercentFormat)]
    public double Margin { get; set; }

    /// <summary>A fraction shown as a percentage, <see cref="Data.PercentFormat"/>.</summary>
    [ExcelColumn(Format = Data.PercentFormat)]
    public double Discount { get; set; }
}
