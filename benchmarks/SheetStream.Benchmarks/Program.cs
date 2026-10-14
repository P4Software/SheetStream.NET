using BenchmarkDotNet.Running;
using SheetStream.Benchmarks;

// Entry point with three modes:
//   (no special switch)          BenchmarkDotNet, e.g. `dotnet run -c Release -- --filter *`
//   --verify                     small-case correctness guard for every library, plus the formatting each wrote
//   --peak                       peak working set per library and scenario, each in a fresh child process
//   --peak-child <lib> <scen>    internal: one child of --peak (the executable re-invokes itself)
if (args.Length > 0 && args[0] == "--peak") return PeakRunner.RunParent();
if (args.Length == 3 && args[0] == "--peak-child") return PeakRunner.RunChild(args[1], args[2]);
if (args.Length > 0 && args[0] == "--verify") return Verify();

BenchmarkSwitcher.FromAssembly(typeof(LongBenchmarks).Assembly).Run(args);
return 0;

// Runs the guard on every library and prints what each one actually wrote, so formatting differences are visible.
static int Verify()
{
    int failures = 0;
    Console.WriteLine("| Library | Scenario | Rows | Last-row cells | Header bold | Price cell type / format | Date cell type / format |");
    Console.WriteLine("|---|---|---:|---:|---|---|---|");
    foreach (var scenario in new[] { "Long", "Wide" })
    {
        foreach (var writer in Writers.All)
        {
            try
            {
                var r = Guard.CheckSmall(writer, scenario);
                Console.WriteLine(scenario == "Long"
                    ? $"| {writer.Name} | {scenario} | {r.Rows} | {r.LastRowCells} | {(r.HeaderBold ? "yes" : "no")} | {r.PriceType} / `{r.PriceFormat}` | {r.DateType} / `{r.DateFormat}` |"
                    : $"| {writer.Name} | {scenario} | {r.Rows} | {r.LastRowCells} | - | - | - |");
            }
            catch (Exception ex)
            {
                failures++;
                Console.WriteLine($"| {writer.Name} | {scenario} | FAILED: {ex.GetType().Name}: {ex.Message} | | | | |");
            }
        }
    }
    Console.WriteLine(failures == 0 ? "All libraries passed the correctness guard." : $"{failures} check(s) failed.");
    return failures == 0 ? 0 : 1;
}
