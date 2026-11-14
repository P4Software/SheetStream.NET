using System.Diagnostics;
using System.Reflection;
using System.Text;

namespace SheetStream.Tests;

/// <summary>Runs alone, because the tests in it measure what the whole process allocates.</summary>
[CollectionDefinition("Memory", DisableParallelization = true)]
public class MemoryCollection { }

/// <summary>
/// Every way this library opens a workbook, against every shape of file built to exhaust memory
/// (<see cref="HostileWorkbooks"/>). The rule under test: opening one either works in bounded memory or is
/// refused with one of the documented exception types, through every entry point, and never by building the
/// oversized thing first.
/// </summary>
/// <remarks>
/// <para>
/// Three checks, each catching what the others cannot. The outcome of every combination is compared with a
/// recorded table (<c>HostileWorkbookOutcomes.txt</c>), because "it did not crash" is too weak: a shape that
/// used to be refused and now reads, or the reverse, is a behaviour change someone has to look at. A refusal
/// must be cheap, so what it allocates is capped. And the whole table is run again in a separate process with
/// a hard limit on its heap, which is the only way to prove that nothing holds the hostile content even for a
/// moment: allocation counts cannot tell a 40 MB string from 40 MB of short-lived rows.
/// </para>
/// <para>
/// After an intended change in behaviour, run the tests with the environment variable
/// <c>SHEETSTREAM_RECORD_OUTCOMES=1</c> to rewrite the table, and review the difference like any other code.
/// </para>
/// </remarks>
[Collection("Memory")]
public class HostileWorkbookTests
{
    /// <summary>
    /// What refusing one of these files may allocate. Under <see cref="HostileWorkbooks.Size"/> as a string
    /// (40 MB), so a path that builds the hostile content before refusing it fails the test, and about
    /// twice the most any refusal has been measured at (15 MB), so another runtime or operating system does
    /// not trip it. The memory probe is the stricter check; this one catches a refusal that became costly.
    /// </summary>
    private const long RefusalBudget = 28L * 1024 * 1024;

    private static string OutcomesFile([System.Runtime.CompilerServices.CallerFilePath] string here = "")
        => Path.Combine(Path.GetDirectoryName(here)!, "HostileWorkbookOutcomes.txt");

    private static Dictionary<string, string> Recorded()
    {
        var table = new Dictionary<string, string>(StringComparer.Ordinal);
        var stream = typeof(HostileWorkbookTests).Assembly.GetManifestResourceStream("HostileWorkbookOutcomes.txt");
        if (stream == null) return table;
        using var reader = new StreamReader(stream);
        while (reader.ReadLine() is { } line)
        {
            if (line.Length == 0 || line[0] == '#') continue;
            var parts = line.Split('|');
            table[parts[0].Trim() + " | " + parts[1].Trim()] = parts[2].Trim();
        }
        return table;
    }

    public static IEnumerable<object[]> ShapeNames() => HostileWorkbooks.Shapes.Keys.Select(name => new object[] { name });

    [Fact]
    public async Task EveryCombination_DoesWhatTheRecordedTableSays()
    {
        var actual = new SortedDictionary<string, string>(StringComparer.Ordinal);
        var costly = new List<string>();
        foreach (var shape in HostileWorkbooks.Shapes.Keys)
        {
            var workbook = HostileWorkbooks.Build(shape);
            foreach (var entry in HostileWorkbooks.Entries.Keys)
            {
                GC.Collect();
                var before = GC.GetTotalAllocatedBytes(precise: true);
                var (outcome, _) = await HostileWorkbooks.Run(workbook, shape, entry);
                var allocated = GC.GetTotalAllocatedBytes(precise: true) - before;
                actual[shape + " | " + entry] = outcome;
                if (outcome == "Refused" && allocated > RefusalBudget)
                    costly.Add($"{shape} | {entry}: refusing allocated {allocated / 1048576.0:F1} MB");
            }
        }

        if (Environment.GetEnvironmentVariable("SHEETSTREAM_RECORD_OUTCOMES") == "1")
        {
            var text = new StringBuilder("# What each way of opening a workbook does with each hostile file. Recorded by HostileWorkbookTests;\n# see that class for how to update it. Reads = completed in bounded memory. Refused = a documented exception.\n");
            foreach (var pair in actual) text.Append(pair.Key).Append(" | ").Append(pair.Value).Append('\n');
            File.WriteAllText(OutcomesFile(), text.ToString());
        }

        var recorded = Recorded();
        var differences = new List<string>();
        foreach (var pair in actual)
        {
            if (!recorded.TryGetValue(pair.Key, out var expected)) differences.Add($"{pair.Key}: not in the recorded table (now {pair.Value})");
            else if (expected != pair.Value) differences.Add($"{pair.Key}: recorded {expected}, now {pair.Value}");
        }
        foreach (var key in recorded.Keys)
            if (!actual.ContainsKey(key)) differences.Add($"{key}: in the recorded table but no longer run");

        Assert.True(differences.Count == 0, "Outcomes differ from HostileWorkbookOutcomes.txt:\n" + string.Join("\n", differences));
        Assert.True(costly.Count == 0, "A refusal must not cost what it refuses:\n" + string.Join("\n", costly));
    }

    [Fact]
    public void EveryCombination_RunsInsideAHardMemoryLimit()
    {
        // The probe is a small console program built next to the tests. Its heap is capped by the runtime
        // itself, so any path that needs more memory than the cap, however briefly, ends in an
        // OutOfMemoryException that the probe reports by name.
        var folder = Path.GetDirectoryName(typeof(HostileWorkbookTests).Assembly.Location)!;
        var probe = Path.Combine(folder, "SheetStream.MemoryProbe.dll");
        Assert.True(File.Exists(probe), $"The memory probe was not built next to the tests: {probe}");

        var start = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            WorkingDirectory = folder,
        };
        start.ArgumentList.Add("exec");
        start.ArgumentList.Add(probe);
        // 96 MB for the whole process: the runtime, the test files themselves (a few megabytes packed) and
        // the work. Any one of the hostile files is 40 MB as a single string.
        start.Environment["DOTNET_GCHeapHardLimit"] = (96L * 1024 * 1024).ToString("X");
        start.Environment["DOTNET_gcServer"] = "0";
        start.Environment["DOTNET_gcConcurrent"] = "0";

        using var process = Process.Start(start)!;
        var errors = process.StandardError.ReadToEndAsync();
        var output = process.StandardOutput.ReadToEnd();
        Assert.True(process.WaitForExit(30 * 60 * 1000), "The memory probe did not finish in 30 minutes.");
        var last = output.Split('\n', StringSplitOptions.RemoveEmptyEntries).LastOrDefault() ?? "(no output)";
        Assert.True(process.ExitCode == 0, $"The memory probe failed (exit code {process.ExitCode}). Last line: {last.Trim()}\n{errors.Result}");
        Assert.Contains("combinations inside the limit", output);
    }
}