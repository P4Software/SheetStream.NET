namespace SheetStream.Benchmarks;

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

    /// <summary>The read scenario: every reader reads the <see cref="Long"/> file that SheetStream wrote.</summary>
    public const string Read = "Read";

    /// <summary>The write scenarios, in report order.</summary>
    public static readonly string[] Write = { Long, Wide, Million, VeryWide };

    /// <summary>Every scenario <c>--peak</c> runs, in report order. Reads come last because they need a written file.</summary>
    public static readonly string[] All = { Long, Wide, Million, VeryWide, Read };

    /// <summary>True for the scenarios that use <see cref="LongRow"/> data and <c>WriteLong</c> (Long and Million).</summary>
    public static bool IsLongShape(string scenario) => scenario is Long or Million or Read;

    /// <summary>Data rows of a write scenario (header excluded).</summary>
    public static int Rows(string scenario) => scenario switch
    {
        Long or Read => Data.LongRows,
        Wide => Data.WideRows,
        Million => Data.MillionRows,
        VeryWide => Data.VeryWideRows,
        _ => throw Unknown(scenario),
    };

    /// <summary>Columns per row of a write scenario.</summary>
    public static int Columns(string scenario) => scenario switch
    {
        Long or Million or Read => Data.LongHeaders.Length,
        Wide => Data.WideColumns,
        VeryWide => Data.VeryWideColumns,
        _ => throw Unknown(scenario),
    };

    /// <summary>
    /// Capacity of the reusable output buffer, chosen so it never grows while timed. 64 MB covers the largest
    /// Long or Wide output (MiniExcel, about 7.4 MB) many times over. The 10M-cell scenarios produce roughly ten
    /// times as much (measured: SheetStream 59.5 MB and MiniExcel 76.6 MB for Million, 47 MB and 50.6 MB for
    /// VeryWide), so 160 MB leaves headroom for any library.
    /// </summary>
    public static int OutputCapacity(string scenario) => scenario is Million or VeryWide ? 160 << 20 : 64 << 20;

    /// <summary>Throws for a scenario name that is not in <see cref="All"/>.</summary>
    public static ArgumentException Unknown(string scenario) =>
        new($"Unknown scenario '{scenario}'. Known: {string.Join(", ", All)}");
}
