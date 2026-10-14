using BenchmarkDotNet.Attributes;

namespace SheetStream.Benchmarks;

/// <summary>
/// Shared plumbing: the reusable output buffer and the correctness guard. Each derived class runs every library
/// on one scenario; BenchmarkDotNet runs each [Benchmark] in its own process, so <see cref="Setup"/> runs per
/// library and the guard covers exactly the library about to be timed.
/// </summary>
[MemoryDiagnoser]
// One launch, 3 warm-up and 10 measured iterations: enough for stable means on operations that take from
// tens of milliseconds to several seconds, while keeping the whole suite well under 15 minutes.
[SimpleJob(launchCount: 1, warmupCount: 3, iterationCount: 10)]
public abstract class ScenarioBenchmarks
{
    /// <summary>64 MB covers the largest output any library produces, so the buffer never grows while timed.</summary>
    private const int OutputCapacity = 64 << 20;

    /// <summary>The output buffer; created once, reset before every write.</summary>
    protected ReusableOutput Output = null!;

    /// <summary>"Long" or "Wide".</summary>
    protected abstract string Scenario { get; }

    /// <summary>Builds the input and the output buffer, then runs the small-case correctness guard.</summary>
    [GlobalSetup]
    public void Setup()
    {
        Output = new ReusableOutput(OutputCapacity);
        BuildData();
        // The guard runs the library that the current benchmark method belongs to. BenchmarkDotNet does not
        // tell setup which method is next, so check all of them; it is cheap at the small size.
        foreach (var writer in Writers.All) Guard.CheckSmall(writer, Scenario);
    }

    /// <summary>Builds the full-size input for the scenario.</summary>
    protected abstract void BuildData();

    /// <summary>Empties the output buffer before each invocation (BenchmarkDotNet excludes this from timing).</summary>
    [IterationSetup]
    public void ResetOutput() => Output.Reset();
}

/// <summary>Long: 100,000 rows x 10 mixed-type columns plus a bold header row.</summary>
public class LongBenchmarks : ScenarioBenchmarks
{
    private LongRow[] _rows = null!;

    /// <inheritdoc />
    protected override string Scenario => "Long";

    /// <inheritdoc />
    protected override void BuildData() => _rows = Data.CreateLong(Data.LongRows);

    [Benchmark(Baseline = true, Description = "SheetStream")]
    public long SheetStream() => Run(Writers.Get("SheetStream"));

    [Benchmark(Description = "SheetStream (Fastest zip)")]
    public long SheetStreamFastest() => Run(Writers.Get("SheetStream-Fastest"));

    [Benchmark(Description = "EPPlus 4.5.3.3")]
    public long EPPlus() => Run(Writers.Get("EPPlus 4.5.3.3"));

    [Benchmark(Description = "ClosedXML")]
    public long ClosedXml() => Run(Writers.Get("ClosedXML"));

    [Benchmark(Description = "MiniExcel")]
    public long MiniExcel() => Run(Writers.Get("MiniExcel"));

    [Benchmark(Description = "NPOI XSSF")]
    public long NpoiXssf() => Run(Writers.Get("NPOI XSSF"));

    [Benchmark(Description = "NPOI SXSSF")]
    public long NpoiSxssf() => Run(Writers.Get("NPOI SXSSF"));

    /// <summary>Writes the full sheet and returns its size, so the work cannot be optimised away.</summary>
    private long Run(ILibraryWriter writer)
    {
        writer.WriteLong(Output, _rows);
        return Output.Length;
    }
}

/// <summary>Wide: 1,000 rows x 1,000 columns of doubles, no header.</summary>
public class WideBenchmarks : ScenarioBenchmarks
{
    private double[,] _grid = null!;

    /// <inheritdoc />
    protected override string Scenario => "Wide";

    /// <inheritdoc />
    protected override void BuildData() => _grid = Data.CreateWide(Data.WideRows, Data.WideColumns);

    [Benchmark(Baseline = true, Description = "SheetStream")]
    public long SheetStream() => Run(Writers.Get("SheetStream"));

    [Benchmark(Description = "SheetStream (Fastest zip)")]
    public long SheetStreamFastest() => Run(Writers.Get("SheetStream-Fastest"));

    [Benchmark(Description = "EPPlus 4.5.3.3")]
    public long EPPlus() => Run(Writers.Get("EPPlus 4.5.3.3"));

    [Benchmark(Description = "ClosedXML")]
    public long ClosedXml() => Run(Writers.Get("ClosedXML"));

    [Benchmark(Description = "MiniExcel")]
    public long MiniExcel() => Run(Writers.Get("MiniExcel"));

    [Benchmark(Description = "NPOI XSSF")]
    public long NpoiXssf() => Run(Writers.Get("NPOI XSSF"));

    [Benchmark(Description = "NPOI SXSSF")]
    public long NpoiSxssf() => Run(Writers.Get("NPOI SXSSF"));

    /// <summary>Writes the full sheet and returns its size, so the work cannot be optimised away.</summary>
    private long Run(ILibraryWriter writer)
    {
        writer.WriteWide(Output, _grid);
        return Output.Length;
    }
}
