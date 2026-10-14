using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// Feeds the reader damaged workbooks. An import screen receives files from users, so a broken file must
/// end in a clear "this file is damaged" exception, never a crash type such as NullReferenceException,
/// IndexOutOfRangeException or a hang.
/// </summary>
public class ReaderFuzzTests
{
    /// <summary>Exception types that mean "the input is bad", which callers can catch and report.</summary>
    private static bool IsExpected(Exception ex) =>
        ex is InvalidDataException || ex is XmlException || ex is FormatException || ex is OverflowException;

    private static byte[] Seed() => ExcelPackage.CreateBytes(wb =>
    {
        var bold = wb.Styles.Create(o => { o.Font.Bold = true; o.NumberFormat = "yyyy-mm-dd"; });
        var s = wb.Worksheets.Add("Data");
        s.Row().WriteCells("Name", "Qty", "When", "Ok");
        for (var i = 0; i < 20; i++)
            s.Row().WriteCell("Item " + i).WriteCell(i * 1.5).WriteCell(new DateTime(2026, 1, 1).AddDays(i), bold).WriteCell(i % 2 == 0)
                .WriteRichText(new RichText("a") { Bold = true }, "b").WriteFormula("B2*2", 3.0);
    });

    /// <summary>Rebuilds the package with one XML part mutated by <paramref name="mutate"/>.</summary>
    private static byte[] MutatePart(byte[] package, string part, Func<string, string> mutate)
    {
        var ms = new MemoryStream();
        ms.Write(package, 0, package.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(part)!;
            string text;
            using (var sr = new StreamReader(entry.Open())) text = sr.ReadToEnd();
            entry.Delete();
            using var w = new StreamWriter(zip.CreateEntry(part).Open(), new UTF8Encoding(false));
            w.Write(mutate(text));
        }
        return ms.ToArray();
    }

    private static void ReadEverything(byte[] bytes)
    {
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        foreach (var name in reader.SheetNames)
            foreach (var row in reader.ReadRows(name))
                for (var c = 0; c < row.Values.Count; c++)
                    _ = row.GetString(c);
    }

    [Theory]
    [InlineData("xl/worksheets/sheet1.xml")]
    [InlineData("xl/sharedStrings.xml")]
    [InlineData("xl/styles.xml")]
    [InlineData("xl/workbook.xml")]
    [InlineData("xl/_rels/workbook.xml.rels")]
    public async Task MutatedXml_FailsCleanlyOrReads(string part)
    {
        var seed = Seed();
        var rng = new Random(part.GetHashCode() ^ 12345); // Deterministic per part, so failures reproduce.
        // Fragments that stress the parser: huge or negative numbers, bad references, broken structure.
        string[] junk = { "<", ">", "\"", "999999999999", "-1", "ZZZZZ1", "XFE1", "&amp;", "<row>", "</c>", "t=\"s\"", "s=\"99999\"", "r=\"A0\"", "\u0000", "1E400" };

        var rejected = 0;
        for (var iteration = 0; iteration < 150; iteration++)
        {
            var mutated = MutatePart(seed, part, text =>
            {
                var sb = new StringBuilder(text);
                var edits = 1 + rng.Next(4);
                for (var e = 0; e < edits && sb.Length > 0; e++)
                {
                    var at = rng.Next(sb.Length);
                    switch (rng.Next(3))
                    {
                        case 0: sb.Remove(at, Math.Min(1 + rng.Next(8), sb.Length - at)); break;
                        case 1: sb.Insert(at, junk[rng.Next(junk.Length)]); break;
                        default: sb[at] = (char)rng.Next(32, 127); break;
                    }
                }
                return sb.ToString();
            });

            try
            {
                // Awaiting (instead of Task.Wait) keeps the test thread free; a reader that hangs on damaged
                // input still fails the test after the timeout instead of blocking the whole run.
                var task = Task.Run(() => ReadEverything(mutated));
                Assert.True(await Task.WhenAny(task, Task.Delay(TimeSpan.FromSeconds(10))) == task, $"Reader hung on iteration {iteration} of {part}.");
                await task;
            }
            catch (Exception ex) when (IsExpected(ex))
            {
                // A clear rejection is the correct outcome for damaged input.
                rejected++;
            }
            catch (Exception ex) when (ex is not Xunit.Sdk.XunitException)
            {
                Assert.Fail($"Iteration {iteration} of {part}: unexpected {ex.GetType().Name}: {ex.Message}");
            }
        }
        // Guard against a fuzzer that never reaches the error paths (for example a part the reader ignores).
        Assert.True(rejected > 0, $"No mutation of {part} was rejected; the fuzzer is not exercising the reader.");
    }

    [Fact]
    public void TruncatedPackages_FailCleanly()
    {
        var seed = Seed();
        foreach (var length in new[] { 0, 10, 100, seed.Length / 3, seed.Length / 2, seed.Length - 30 })
        {
            var cut = seed.Take(length).ToArray();
            var ex = Record.Exception(() => ReadEverything(cut));
            Assert.True(ex == null || IsExpected(ex), $"Length {length}: unexpected {ex?.GetType().Name}: {ex?.Message}");
        }
    }
}
