using BenchmarkDotNet.Attributes;

namespace SheetStream.Benchmarks;

/// <summary>
/// Shared plumbing: the reusable output buffer and the correctness guard. Each derived class runs every library
/// on one scenario; BenchmarkDotNet runs each [Benchmark] in its own process, so <see cref="Setup"/> runs per
/// library and the guard covers exactly the library about to be timed.
/// </summary>
/// <remarks>
/// The job lives on each concrete class, not here: BenchmarkDotNet collects config attributes along the type
/// hierarchy, so a <c>[SimpleJob]</c> on the base plus a lighter one on a derived class would run both jobs.
/// </remarks>
[MemoryDiagnoser]
public abstract class ScenarioBenchmarks
{
    /// <summary>The output buffer; created once with the scenario's capacity, reset before every write.</summary>
    protected ReusableOutput Output = null!;

    /// <summary>One of <see cref="Scenarios.Write"/>.</summary>
    protected abstract string Scenario { get; }

    /// <summary>Builds the input and the output buffer, then runs the small-case correctness guard.</summary>
    [GlobalSetup]
    public void Setup()
    {
        Output = new ReusableOutput(Scenarios.OutputCapacity(Scenario));
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

/// <summary>
/// A Long-shaped scenario: <see cref="LongRow"/> data through <see cref="ILibraryWriter.WriteLong"/>. Long and
/// Million differ only in the row count, so the benchmark methods are declared once here.
/// </summary>
public abstract class LongShapeBenchmarks : ScenarioBenchmarks
{
    private LongRow[] _rows = null!;

    /// <inheritdoc />
    protected override void BuildData() => _rows = Data.CreateLong(Scenarios.Rows(Scenario));

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

/// <summary>
/// A Wide-shaped scenario: a grid of doubles through <see cref="ILibraryWriter.WriteWide"/>. Wide and VeryWide
/// differ only in the column count, so the benchmark methods are declared once here.
/// </summary>
public abstract class WideShapeBenchmarks : ScenarioBenchmarks
{
    private double[,] _grid = null!;

    /// <inheritdoc />
    protected override void BuildData() => _grid = Data.CreateWide(Scenarios.Rows(Scenario), Scenarios.Columns(Scenario));

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

/// <summary>Long: 100,000 rows x 10 mixed-type columns plus a bold header row.</summary>
// One launch, 3 warm-up and 10 measured iterations: enough for stable means on operations that take from
// tens of milliseconds to several seconds, while keeping the whole suite well under 15 minutes.
[SimpleJob(launchCount: 1, warmupCount: 3, iterationCount: 10)]
public class LongBenchmarks : LongShapeBenchmarks
{
    /// <inheritdoc />
    protected override string Scenario => Scenarios.Long;
}

/// <summary>Wide: 1,000 rows x 1,000 columns of doubles, no header.</summary>
[SimpleJob(launchCount: 1, warmupCount: 3, iterationCount: 10)]
public class WideBenchmarks : WideShapeBenchmarks
{
    /// <inheritdoc />
    protected override string Scenario => Scenarios.Wide;
}

/// <summary>Million: 1,000,000 rows x the same 10 columns as Long (10M cells).</summary>
// A lighter job (1 warm-up, 3 measured iterations): the in-memory libraries take 20 to 60 seconds per write
// at this size, so the standard 13 invocations per library would push the suite past an hour.
[SimpleJob(launchCount: 1, warmupCount: 1, iterationCount: 3)]
public class MillionBenchmarks : LongShapeBenchmarks
{
    /// <inheritdoc />
    protected override string Scenario => Scenarios.Million;
}

/// <summary>VeryWide: 1,000 rows x 10,000 columns of doubles (10M cells, the same count as Million).</summary>
[SimpleJob(launchCount: 1, warmupCount: 1, iterationCount: 3)]
public class VeryWideBenchmarks : WideShapeBenchmarks
{
    /// <inheritdoc />
    protected override string Scenario => Scenarios.VeryWide;
}

/// <summary>
/// Read: every reader reads the Long file (100,000 rows x 10 columns) that SheetStream wrote once in
/// <see cref="Setup"/>. Each benchmark returns the checksum of what it read, so the work cannot be optimised
/// away and the guard can compare it with the generator's figure.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 1, warmupCount: 3, iterationCount: 10)]
public class ReadBenchmarks
{
    private byte[] _file = null!;
    private MemoryStream _input = null!;

    /// <summary>Writes the input file (not timed) and runs the small-case reader guard for every reader.</summary>
    [GlobalSetup]
    public void Setup()
    {
        _file = ReadInput.Create(Data.LongRows);
        foreach (var reader in Readers.All) ReaderGuard.CheckSmall(reader);
    }

    /// <summary>Opens a fresh read-only stream over the shared bytes; readers may dispose it, so one per invocation.</summary>
    [IterationSetup]
    public void OpenInput() => _input = ReadInput.Open(_file);

    [Benchmark(Baseline = true, Description = "SheetStream ReadRows")]
    public long SheetStreamRows() => Run("SheetStream-Rows");

    [Benchmark(Description = "SheetStream Read<T>")]
    public long SheetStreamTyped() => Run("SheetStream-Typed");

    [Benchmark(Description = "MiniExcel Query<T>")]
    public long MiniExcelTyped() => Run("MiniExcel-Typed");

    [Benchmark(Description = "MiniExcel Query (dynamic)")]
    public long MiniExcelDynamic() => Run("MiniExcel-Dynamic");

    [Benchmark(Description = "ExcelDataReader")]
    public long ExcelDataReader() => Run("ExcelDataReader");

    [Benchmark(Description = "ClosedXML")]
    public long ClosedXml() => Run("ClosedXML");

    [Benchmark(Description = "EPPlus 4.5.3.3")]
    public long EPPlus() => Run("EPPlus 4.5.3.3");

    [Benchmark(Description = "NPOI XSSF")]
    public long NpoiXssf() => Run("NPOI XSSF");

    [Benchmark(Description = "NPOI SAX")]
    public long NpoiSax() => Run("NPOI SAX");

    private long Run(string reader) => Readers.Get(reader).ReadLong(_input);
}
