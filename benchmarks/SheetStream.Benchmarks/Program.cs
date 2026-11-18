using BenchmarkDotNet.Running;
using SheetStream.Benchmarks;

// Entry point with three modes:
//   (no special switch)          BenchmarkDotNet, e.g. `dotnet run -c Release -- --filter *`
//   --verify                     small-case correctness guard for every writer and reader, plus the formatting each wrote
//   --peak [scenario...]         peak working set per library and scenario, each in a fresh child process;
//                                with no scenario every scenario runs (see Scenarios.All)
//   --peak-child <lib> <scen>    internal: one child of --peak (the executable re-invokes itself)
if (args.Length > 0 && args[0] == "--peak") return PeakRunner.RunParent(args[1..]);
if (args.Length == 3 && args[0] == "--peak-child") return PeakRunner.RunChild(args[1], args[2]);
if (args.Length > 0 && args[0] == "--verify") return Verify();

BenchmarkSwitcher.FromAssembly(typeof(LongBenchmarks).Assembly).Run(args);
return 0;

// Runs the guard on every writer and reader and prints what each one actually wrote or read, so differences are visible.
static int Verify()
{
    int failures = 0;
    // What the string scenarios' generators produce, so the data's shape is visible next to the results.
    PrintDataProfile(Scenarios.SpreadCheetahOwn, Data.CreateOwnStrings(Scenarios.Rows(Scenarios.SpreadCheetahOwn)));
    PrintDataProfile(Scenarios.StringHeavy, Data.CreateStringHeavy(Scenarios.Rows(Scenarios.StringHeavy)));

    Console.WriteLine("| Library | Scenario | Rows | Last-row cells | Header bold | Price cell type / format | Date cell type / format | Cells value-checked | Formats in a data row |");
    Console.WriteLine("|---|---|---:|---:|---|---|---|---:|---|");
    foreach (var scenario in Scenarios.Write)
    {
        foreach (var writer in Writers.All)
        {
            try
            {
                var r = Guard.CheckSmall(writer, scenario);
                string values = r.ValuesChecked == 0 ? "-" : r.ValuesChecked.ToString();
                string header = Scenarios.HasHeader(scenario) ? (r.HeaderBold ? "yes" : "no") : "-";
                Console.WriteLine(Scenarios.IsLongShape(scenario)
                    ? $"| {writer.Name} | {scenario} | {r.Rows} | {r.LastRowCells} | {header} | {r.PriceType} / `{r.PriceFormat}` | {r.DateType} / `{r.DateFormat}` | {values} | - |"
                    : $"| {writer.Name} | {scenario} | {r.Rows} | {r.LastRowCells} | {header} | - | - | {values} | {(r.Formats == "-" ? "-" : $"`{r.Formats}`")} |");
            }
            catch (Exception ex)
            {
                failures++;
                Console.WriteLine($"| {writer.Name} | {scenario} | FAILED: {ex.GetType().Name}: {ex.Message} | | | | | | |");
            }
        }
    }

    // Readers: every one must reproduce the generator's checksum for the small Long file SheetStream wrote.
    Console.WriteLine();
    Console.WriteLine($"| Reader | Data rows | Checksum | Expected |");
    Console.WriteLine("|---|---:|---:|---:|");
    long expected = Checksum.Expected(Guard.SmallLongRows);
    foreach (var reader in Readers.All)
    {
        try
        {
            long checksum = ReaderGuard.CheckSmall(reader);
            Console.WriteLine($"| {reader.Name} | {Guard.SmallLongRows} | {checksum:N0} | {expected:N0} |");
        }
        catch (Exception ex)
        {
            failures++;
            Console.WriteLine($"| {reader.Name} | {Guard.SmallLongRows} | FAILED: {ex.GetType().Name}: {ex.Message} | {expected:N0} |");
        }
    }

    Console.WriteLine(failures == 0 ? "All libraries passed the correctness guard." : $"{failures} check(s) failed.");
    return failures == 0 ? 0 : 1;
}

// Prints how many distinct values each column holds, the longest value and how many cells contain a character
// that XML has to escape, so anyone can see that a "unique" column is unique and a "repeated" one is repeated.
static void PrintDataProfile(string scenario, string[][] rows)
{
    Console.WriteLine($"Data profile of {scenario} ({rows.Length:N0} rows x {rows[0].Length} columns):");
    Console.WriteLine("| Column | Distinct values | Longest value (chars) | Cells containing & < > or a double quote |");
    Console.WriteLine("|---:|---:|---:|---:|");
    for (int c = 0; c < rows[0].Length; c++)
    {
        var distinct = new HashSet<string>(StringComparer.Ordinal);
        int longest = 0, escaped = 0;
        foreach (var row in rows)
        {
            distinct.Add(row[c]);
            longest = Math.Max(longest, row[c].Length);
            if (row[c].IndexOfAny(new[] { '&', '<', '>', '"' }) >= 0) escaped++;
        }
        Console.WriteLine($"| {c + 1} | {distinct.Count:N0} | {longest} | {escaped:N0} |");
    }
    Console.WriteLine();
}
