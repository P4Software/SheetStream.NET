// Opens every hostile workbook through every entry point and prints what happened. Exit code 0 means every
// combination either read the file or refused it with a documented exception; anything else names the
// combination that failed. Run with DOTNET_GCHeapHardLimit set to make "used too much memory" a failure.
using System.Diagnostics;
using SheetStream.Tests;

var only = args.Length > 0 ? args[0] : null;
var combinations = 0;
var watch = Stopwatch.StartNew();
foreach (var shape in HostileWorkbooks.Shapes.Keys)
{
    if (only != null && !shape.Contains(only, StringComparison.OrdinalIgnoreCase)) continue;
    // Built for this shape only and dropped afterwards, so the files do not add up under the limit.
    var workbook = HostileWorkbooks.Build(shape);
    foreach (var entry in HostileWorkbooks.Entries.Keys)
    {

        // Printed before the work, so a process killed outright still shows which combination it was on.
        Console.Write($"{shape} | {entry} | ");
        Console.Out.Flush();
        try
        {
            var (outcome, _) = await HostileWorkbooks.Run(workbook, shape, entry);
            Console.WriteLine(outcome);
            combinations++;
        }
        catch (OutOfMemoryException)
        {
            Console.WriteLine("OUT OF MEMORY");
            Console.Error.WriteLine($"'{entry}' on '{shape}' needed more memory than the limit allows.");
            return 3;
        }
        catch (Exception ex)
        {
            Console.WriteLine("FAILED");
            Console.Error.WriteLine(ex.Message);
            return 2;
        }
    }
}
Console.WriteLine($"{combinations} combinations inside the limit, in {watch.Elapsed.TotalSeconds:F0} s, peak working set {Process.GetCurrentProcess().PeakWorkingSet64 / 1048576} MB.");
return 0;