namespace SheetStream.Benchmarks;

/// <summary>
/// How a scenario's input is shaped. Writers implement one method per shape (<see cref="ILibraryWriter"/>), and
/// the guard and the peak mode pick the input generator and the expected layout from it.
/// </summary>
public enum ScenarioShape
{
    /// <summary><see cref="LongRow"/> data, a bold header row, <c>WriteLong</c>.</summary>
    Long,

    /// <summary>A grid of doubles, no header, <c>WriteWide</c>.</summary>
    Wide,

    /// <summary>A grid of strings, no header, <c>WriteStrings</c>.</summary>
    Strings,

    /// <summary><see cref="StyledRow"/> data, a bold header row, <c>WriteStyled</c>.</summary>
    Styled,
}

/// <summary>
/// The scenarios the suite knows, with the sizes and output capacity each one needs. Everything that is keyed
/// by scenario name (benchmark classes, the guard, the peak children and <c>--verify</c>) goes through here,
/// so adding a scenario is one entry in this file plus one benchmark class.
/// </summary>
public static class Scenarios
{
    /// <summary>100,000 rows x 10 mixed columns plus a bold header row.</summary>
    public const string Long = "Long";

    /// <summary>1,000 rows x 1,000 columns of doubles, no header.</summary>
    public const string Wide = "Wide";

    /// <summary>1,000,000 rows x the same 10 columns as <see cref="Long"/>: ten times the rows, same shape.</summary>
    public const string Million = "Million";

    /// <summary>1,000 rows x 10,000 columns of doubles: the same 10M cells as <see cref="Million"/> in the other shape.</summary>
    public const string VeryWide = "VeryWide";

    /// <summary>
    /// SpreadCheetah's own published benchmark shape: 100,000 rows x 10 columns, every cell a distinct string, no
    /// header, no styles. It is the scenario on which SpreadCheetah's README shows its best numbers, so it is the
    /// one a SpreadCheetah user would expect to see here.
    /// </summary>
    public const string SpreadCheetahOwn = "SpreadCheetahOwn";

    /// <summary>
    /// 100,000 rows x 10 string columns with a realistic mix: low-cardinality columns (status, country, category),
    /// unique columns (ids, e-mail, SKU) and free text with characters that need XML escaping.
    /// </summary>
    public const string StringHeavy = "StringHeavy";

    /// <summary>
    /// 100,000 rows x 8 columns of dates, date-times, money and percentages (four distinct styles in every row)
    /// plus a bold header row.
    /// </summary>
    public const string StyledDates = "StyledDates";

    /// <summary>The read scenario: every reader reads the <see cref="Long"/> file that SheetStream wrote.</summary>
    public const string Read = "Read";

    /// <summary>The write scenarios, in report order.</summary>
    public static readonly string[] Write = { Long, Wide, SpreadCheetahOwn, StringHeavy, StyledDates, Million, VeryWide };

    /// <summary>Every scenario <c>--peak</c> runs, in report order. Reads come last because they need a written file.</summary>
    public static readonly string[] All = { Long, Wide, SpreadCheetahOwn, StringHeavy, StyledDates, Million, VeryWide, Read };

    /// <summary>The input shape of a scenario (<see cref="Read"/> uses the Long file, so it is Long-shaped).</summary>
    public static ScenarioShape Shape(string scenario) => scenario switch
    {
        Long or Million or Read => ScenarioShape.Long,
        Wide or VeryWide => ScenarioShape.Wide,
        SpreadCheetahOwn or StringHeavy => ScenarioShape.Strings,
        StyledDates => ScenarioShape.Styled,
        _ => throw Unknown(scenario),
    };

    /// <summary>True for the scenarios that use <see cref="LongRow"/> data and <c>WriteLong</c> (Long and Million).</summary>
    public static bool IsLongShape(string scenario) => Shape(scenario) == ScenarioShape.Long;

    /// <summary>True when the first row of the written sheet is a caption row rather than data.</summary>
    public static bool HasHeader(string scenario) => Shape(scenario) is ScenarioShape.Long or ScenarioShape.Styled;

    /// <summary>Data rows of a write scenario (header excluded).</summary>
    public static int Rows(string scenario) => scenario switch
    {
        Long or Read or SpreadCheetahOwn or StringHeavy or StyledDates => Data.LongRows,
        Wide => Data.WideRows,
        Million => Data.MillionRows,
        VeryWide => Data.VeryWideRows,
        _ => throw Unknown(scenario),
    };

    /// <summary>Columns per row of a write scenario.</summary>
    public static int Columns(string scenario) => scenario switch
    {
        Long or Million or Read => Data.LongHeaders.Length,
        SpreadCheetahOwn or StringHeavy => Data.StringColumns,
        StyledDates => Data.StyledHeaders.Length,
        Wide => Data.WideColumns,
        VeryWide => Data.VeryWideColumns,
        _ => throw Unknown(scenario),
    };

    /// <summary>
    /// Capacity of the reusable output buffer, chosen so it never grows while timed. 64 MB covers the largest
    /// Long or Wide output (MiniExcel, about 7.4 MB) many times over, and the string and styled scenarios are the
    /// same 100,000 rows (their outputs are within the same order of magnitude). The 10M-cell scenarios produce
    /// roughly ten times as much (measured: SheetStream 59.5 MB and MiniExcel 76.6 MB for Million, 47 MB and
    /// 50.6 MB for VeryWide), so 160 MB leaves headroom for any library.
    /// </summary>
    public static int OutputCapacity(string scenario) => scenario is Million or VeryWide ? 160 << 20 : 64 << 20;

    /// <summary>Throws for a scenario name that is not in <see cref="All"/>.</summary>
    public static ArgumentException Unknown(string scenario) =>
        new($"Unknown scenario '{scenario}'. Known: {string.Join(", ", All)}");
}
