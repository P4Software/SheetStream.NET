using System.Diagnostics;
using System.Globalization;

namespace SheetStream.Benchmarks;

/// <summary>
/// The <c>--peak</c> mode. BenchmarkDotNet reports bytes allocated, not how much memory the process actually
/// held at its worst, so this mode re-launches the same executable once per library and scenario and reads
/// <see cref="Process.PeakWorkingSet64"/> in that fresh child. A "Baseline" child builds the input and writes
/// nothing, which is the floor every library sits on (runtime + loaded assemblies + input data + output buffer).
/// </summary>
public static class PeakRunner
{
    /// <summary>Runs per library and scenario; the median is reported to damp OS noise.</summary>
    private const int Runs = 3;

    /// <summary>Child marker printed before the machine-readable result line.</summary>
    private const string ResultPrefix = "PEAKRESULT ";

    /// <summary>Parent side: launches every child, collects results and prints Markdown tables.</summary>
    public static int RunParent()
    {
        string exe = Environment.ProcessPath ?? throw new InvalidOperationException("Cannot find this executable's path.");
        var names = new[] { "Baseline" }.Concat(Writers.All.Select(w => w.Name)).ToArray();

        foreach (var scenario in new[] { "Long", "Wide" })
        {
            Console.WriteLine();
            Console.WriteLine($"### Peak working set: {scenario} ({Runs} fresh processes each, median shown)");
            Console.WriteLine();
            Console.WriteLine("| Library | Peak working set (MB) | Above baseline (MB) | Cold write time (ms) | Output size (KB) | Full-size rows verified |");
            Console.WriteLine("|---|---:|---:|---:|---:|---|");
            double baselineMb = 0;
            foreach (var name in names)
            {
                var results = new List<ChildResult>();
                string? failure = null;
                for (int i = 0; i < Runs && failure == null; i++)
                {
                    var (result, error) = RunChild(exe, name, scenario);
                    if (result != null) results.Add(result);
                    else failure = error;
                }
                if (failure != null)
                {
                    Console.WriteLine($"| {name} | FAILED: {failure.Replace('|', '/').Replace('\n', ' ')} | | | | |");
                    continue;
                }
                double peakMb = Median(results.Select(r => r.PeakBytes / 1048576.0));
                double ms = Median(results.Select(r => r.ElapsedMs));
                if (name == "Baseline") baselineMb = peakMb;
                string above = name == "Baseline" ? "-" : (peakMb - baselineMb).ToString("F1", CultureInfo.InvariantCulture);
                string size = name == "Baseline" ? "-" : (results[0].OutputBytes / 1024.0).ToString("N0", CultureInfo.InvariantCulture);
                Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
                    $"| {name} | {peakMb:F1} | {above} | {(name == "Baseline" ? "-" : ms.ToString("N0", CultureInfo.InvariantCulture))} | {size} | {results[0].Verified} |"));
            }
        }
        return 0;
    }

    /// <summary>Child side: one write, then the peak working set, then (outside the measurement) a full-size verify.</summary>
    public static int RunChild(string library, string scenario)
    {
        // Build the input first, exactly as the benchmarks do; it is part of every library's floor.
        LongRow[]? rows = scenario == "Long" ? Data.CreateLong(Data.LongRows) : null;
        double[,]? grid = scenario == "Wide" ? Data.CreateWide(Data.WideRows, Data.WideColumns) : null;
        using var output = new ReusableOutput(64 << 20);

        var sw = Stopwatch.StartNew();
        if (library != "Baseline")
        {
            var writer = Writers.Get(library);
            if (rows != null) writer.WriteLong(output, rows);
            else writer.WriteWide(output, grid!);
        }
        sw.Stop();

        // Read the peak before verifying, so the Open XML SDK's own memory is not charged to the library.
        var me = Process.GetCurrentProcess();
        me.Refresh();
        long peak = me.PeakWorkingSet64;

        string verified = "-";
        if (library != "Baseline")
        {
            var report = rows != null
                ? Guard.Check(library, scenario, output, Data.LongRows, Data.LongHeaders.Length)
                : Guard.Check(library, scenario, output, Data.WideRows, Data.WideColumns);
            verified = $"yes ({report.Rows:N0} rows x {report.LastRowCells:N0} cells)";
        }

        Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
            $"{ResultPrefix}{peak}|{sw.Elapsed.TotalMilliseconds}|{output.Length}|{verified}"));
        return 0;
    }

    private sealed record ChildResult(long PeakBytes, double ElapsedMs, long OutputBytes, string Verified);

    private static (ChildResult? Result, string? Error) RunChild(string exe, string library, string scenario)
    {
        var psi = new ProcessStartInfo(exe)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        psi.ArgumentList.Add("--peak-child");
        psi.ArgumentList.Add(library);
        psi.ArgumentList.Add(scenario);

        using var p = Process.Start(psi)!;
        // Read stderr asynchronously so a chatty child cannot deadlock on a full pipe.
        var errTask = p.StandardError.ReadToEndAsync();
        string stdout = p.StandardOutput.ReadToEnd();
        p.WaitForExit();
        string stderr = errTask.Result;

        var line = stdout.Split('\n').FirstOrDefault(l => l.StartsWith(ResultPrefix, StringComparison.Ordinal));
        if (p.ExitCode != 0 || line == null)
        {
            string detail = (stderr + " " + stdout).Trim();
            return (null, $"exit code {p.ExitCode}: {(detail.Length > 300 ? detail[..300] : detail)}");
        }
        var parts = line[ResultPrefix.Length..].Trim().Split('|');
        return (new ChildResult(
            long.Parse(parts[0], CultureInfo.InvariantCulture),
            double.Parse(parts[1], CultureInfo.InvariantCulture),
            long.Parse(parts[2], CultureInfo.InvariantCulture),
            parts[3]), null);
    }

    private static double Median(IEnumerable<double> values)
    {
        var sorted = values.OrderBy(v => v).ToArray();
        return sorted.Length % 2 == 1 ? sorted[sorted.Length / 2] : (sorted[sorted.Length / 2 - 1] + sorted[sorted.Length / 2]) / 2;
    }
}
