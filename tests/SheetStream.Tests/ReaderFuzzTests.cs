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
    /// <summary>
    /// How many damaged files each fuzz test tries per part. The default keeps the suite quick; set the
    /// environment variable <c>SHEETSTREAM_FUZZ_ITERATIONS</c> to a few thousand for a deep run before a
    /// release, and <c>SHEETSTREAM_FUZZ_SEED</c> to any number to try files the fixed seed never produces.
    /// </summary>
    private static int Iterations(int normal)
        => int.TryParse(Environment.GetEnvironmentVariable("SHEETSTREAM_FUZZ_ITERATIONS"), out var n) && n > 0 ? n : normal;

    private static int SeedOffset()
        => int.TryParse(Environment.GetEnvironmentVariable("SHEETSTREAM_FUZZ_SEED"), out var n) ? n : 0;

    /// <summary>Exception types that mean "the input is bad", which callers can catch and report.</summary>
    private static bool IsExpected(Exception ex) =>
        ex is InvalidDataException || ex is XmlException || ex is FormatException || ex is OverflowException;

    /// <summary>
    /// The seed workbook. By default its cells are written without their address (consecutive cells need none), so a
    /// mutation of the sheet part never meets an <c>r</c> attribute to damage. <paramref name="references"/> writes every
    /// cell with its address, as Excel does and as <see cref="PackageOptions.WriteCellReferences"/> asks, so the fuzzer
    /// still mutates cell references: bad columns, rows past the sheet, references out of order.
    /// </summary>
    private static byte[] Seed(bool references = false) => ExcelPackage.CreateBytes(wb =>
    {
        var bold = wb.Styles.Create(o => { o.Font.Bold = true; o.NumberFormat = "yyyy-mm-dd"; });
        var s = wb.Worksheets.Add("Data");
        s.Row().WriteCells("Name", "Qty", "When", "Ok");
        for (var i = 0; i < 20; i++)
            s.Row().WriteCell("Item " + i).WriteCell(i * 1.5).WriteCell(new DateTime(2026, 1, 1).AddDays(i), bold).WriteCell(i % 2 == 0)
                .WriteRichText(new RichText("a") { Bold = true }, "b").WriteFormula("B2*2", 3.0).WriteCell(TimeSpan.FromMinutes(i * 7));
    }, new PackageOptions { WriteCellReferences = references });

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

    [Fact]
    public void NegativeStyleIndex_ReadsAsPlainValue()
    {
        // Regression: s="-1" passed the "< Count" guard and indexed the style table at -1. Found by the fuzzer in CI.
        // Cells that already carry a style keep a single (now negative) s attribute; the rest get one added.
        var mutated = MutatePart(Seed(), "xl/worksheets/sheet1.xml", text =>
        {
            var result = System.Text.RegularExpressions.Regex.Replace(
                System.Text.RegularExpressions.Regex.Replace(text, "(<c\\b[^>]*?\\bs=\")\\d+(\")", "$1-1$2"),
                // A cell may or may not carry r (the writer leaves it off consecutive cells), so match the tag itself.
                "<c\\b(?![^>]*\\bs=)", "<c s=\"-1\"");
            // Guard the mutation itself: if the pattern stops matching, this test would silently test nothing.
            var cells = System.Text.RegularExpressions.Regex.Matches(result, "<c\\b").Count;
            Assert.True(cells > 0, "The seed sheet has no cells to mutate.");
            Assert.Equal(cells, System.Text.RegularExpressions.Regex.Matches(result, "<c\\b[^>]*\\bs=\"-1\"").Count);
            return result;
        });
        using var reader = ExcelReader.Open(new MemoryStream(mutated));
        var rows = reader.ReadRows(0).ToList();
        Assert.NotEmpty(rows);
        Assert.Contains(rows, r => r.Values.Any(v => v != null));
    }

    [Theory]
    [InlineData("xl/worksheets/sheet1.xml")]
    [InlineData("xl/sharedStrings.xml")]
    [InlineData("xl/styles.xml")]
    [InlineData("xl/workbook.xml")]
    [InlineData("xl/_rels/workbook.xml.rels")]
    public Task MutatedXml_FailsCleanlyOrReads(string part) => MutatePartUntilItFailsCleanly(Seed(), part, references: false);

    /// <summary>The same fuzzing of the sheet and the shared strings over a workbook whose cells all carry an <c>r</c> address.</summary>
    [Theory]
    [InlineData("xl/worksheets/sheet1.xml")]
    [InlineData("xl/sharedStrings.xml")]
    public Task MutatedXml_OfAWorkbookWithCellReferences_FailsCleanlyOrReads(string part) => MutatePartUntilItFailsCleanly(Seed(references: true), part, references: true);

    private static async Task MutatePartUntilItFailsCleanly(byte[] seed, string part, bool references)
    {
        // Deterministic per part, so failures reproduce. string.GetHashCode() is randomised per process on
        // .NET Core, so a stable FNV-1a hash is used instead; the first run with a random seed found the s="-1" bug.
        var seedHash = 2166136261u;
        foreach (var ch in part) seedHash = (seedHash ^ ch) * 16777619u;
        var rng = new Random(unchecked((int)seedHash) ^ 12345 ^ SeedOffset() ^ (references ? 0x5EED : 0));
        // Fragments that stress the parser: huge or negative numbers, bad references, broken structure.
        string[] junk = { "<", ">", "\"", "999999999999", "-1", "ZZZZZ1", "XFE1", "&amp;", "<row>", "</c>", "t=\"s\"", "s=\"99999\"", "r=\"A0\"", "\u0000", "1E400", "NaN", "1E300", "<a><a><a>" };

        var rejected = 0;
        for (int iteration = 0, iterations = Iterations(150); iteration < iterations; iteration++)
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

    /// <summary>A workbook with the things template and edit mode rewrite: shared strings, styles, a table, formulas, dates.</summary>
    private static byte[] TemplateSeed() => ExcelPackage.CreateBytes(wb =>
    {
        var bold = wb.Styles.Create(o => { o.Font.Bold = true; o.NumberFormat = "yyyy-mm-dd"; });
        var s = wb.Worksheets.Add("Data");
        s.AddTable("Items", "A1:D6");
        s.Row().WriteCells("Name", "Qty", "When", "Total");
        for (var i = 0; i < 5; i++)
            s.Row().WriteCell("Item " + i).WriteCell(i * 1.5).WriteCell(new DateTime(2026, 1, 1).AddDays(i), bold).WriteFormula("B2*2", 3.0);
        s.MergeCells("F1:G2");
        wb.Worksheets.Add("Other").Row().WriteCell("untouched").WriteRichText(new RichText("a") { Bold = true }, "b");
    });

    /// <summary>
    /// The same damage as <see cref="MutatedXml_FailsCleanlyOrReads"/>, through the entry points that load a
    /// workbook in order to write a changed copy. They parse far more of the file than the reader does
    /// (styles in full, tables, every cell's XML), and an application may hand them a workbook a user uploaded.
    /// </summary>
    [Theory]
    [InlineData("xl/worksheets/sheet1.xml")]
    [InlineData("xl/worksheets/sheet2.xml")]
    [InlineData("xl/sharedStrings.xml")]
    [InlineData("xl/styles.xml")]
    [InlineData("xl/workbook.xml")]
    [InlineData("xl/_rels/workbook.xml.rels")]
    [InlineData("xl/worksheets/_rels/sheet1.xml.rels")]
    [InlineData("xl/tables/table1.xml")]
    [InlineData("[Content_Types].xml")]
    [InlineData("_rels/.rels")]
    public async Task MutatedXml_ThroughEditAndTemplate_FailsCleanlyOrSaves(string part)
    {
        var seed = TemplateSeed();
        var seedHash = 2166136261u;
        foreach (var ch in part) seedHash = (seedHash ^ ch) * 16777619u;
        var rng = new Random(unchecked((int)seedHash) ^ 6789 ^ SeedOffset());
        string[] junk = { "<", ">", "\"", "999999999999", "-1", "ZZZZZ1", "XFE1", "&amp;", "<row>", "</c>", "t=\"s\"", "s=\"99999\"", "r=\"A0\"", "\u0000", "1E400", "NaN", "1E300", "<a><a><a>", "r=\"-3\"", "ref=\"A1:\"", "count=\"-1\"", "<![CDATA[", "]]>", "<f t=\"shared\" si=\"7\"/>", "numFmtId=\"-9\"", "xfId=\"70000\"", "<xf/>", "<sheet/>", "r:id=\"rId99\"", "Target=\"../../../x\"" };

        var actions = new (string Name, Action<byte[]> Run)[]
        {
            ("edit both sheets", bytes =>
            {
                using var package = ExcelPackage.Edit(new MemoryStream(bytes), Stream.Null, leaveOpen: true);
                foreach (var name in new[] { "Data", "Other" })
                {
                    var sheet = package.Workbook.Worksheets.Edit(name);
                    sheet.Cells["B2"].Value = 12.5;
                    _ = sheet.Cells["A1"].Value;
                    sheet.Cells[9, 1].Value = "appended";
                }
                package.Save();
            }),
            ("fill a template sheet", bytes =>
            {
                using var package = ExcelPackage.OpenTemplate(new MemoryStream(bytes), Stream.Null, leaveOpen: true);
                var sheet = package.Workbook.Worksheets.Open("Data");
                var style = sheet.TemplateStyle("C2");
                for (var i = 0; i < 4; i++) sheet.Row(7 + i).WriteCell("x").WriteCell(i).WriteCell(new DateTime(2026, 2, 1), style).WriteFormula("B7*2");
                package.Save();
            }),
            ("copy untouched", bytes =>
            {
                using var package = ExcelPackage.OpenTemplate(new MemoryStream(bytes), Stream.Null, leaveOpen: true);
                package.Save();
            }),
        };

        var rejected = 0;
        var saved = 0;
        for (int iteration = 0, iterations = Iterations(120); iteration < iterations; iteration++)
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

            foreach (var (name, run) in actions)
            {
                try
                {
                    var task = Task.Run(() => run(mutated));
                    Assert.True(await Task.WhenAny(task, Task.Delay(TimeSpan.FromSeconds(10))) == task, $"'{name}' hung on iteration {iteration} of {part}.");
                    await task;
                    saved++;
                }
                catch (Exception ex) when (IsExpected(ex))
                {
                    rejected++;
                }
                catch (ArgumentException ex) when (ex.ParamName == "name" && ex.Message.Contains("no sheet named"))
                {
                    // The damage renamed or removed the sheet the caller asked for. That is the caller's
                    // argument not matching the file, reported the same way as for a genuine workbook.
                    rejected++;
                }
                catch (Exception ex) when (ex is not Xunit.Sdk.XunitException)
                {
                    Assert.Fail($"'{name}', iteration {iteration} of {part}: unexpected {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
                }
            }
        }
        // Both must happen, or the fuzzer is not reaching the code: nothing rejected means the part is ignored,
        // nothing saved means the seed itself is broken.
        Assert.True(saved > 0, $"No mutation of {part} could be saved; the seed workbook is wrong.");
        Assert.True(rejected > 0 || part == "[Content_Types].xml" || part == "_rels/.rels" || part.EndsWith("sheet1.xml.rels", StringComparison.Ordinal), $"No mutation of {part} was rejected; the fuzzer is not exercising the template code.");
    }

    // The project copies tests/corpus/excel beside the test binaries.
    private static string CorpusFolder() => Path.Combine(AppContext.BaseDirectory, "corpus", "excel");

    public static IEnumerable<object[]> CorpusWorkbooks()
        => Directory.GetFiles(CorpusFolder(), "*.xlsx").Select(f => new object[] { Path.GetFileNameWithoutExtension(f) }).OrderBy(a => (string)a[0]);

    /// <summary>
    /// The same damage again, starting from workbooks Excel wrote. They hold parts and constructs the library
    /// never writes itself (themes, calculation chains, extension lists, pivot caches, drawings), and damage
    /// to those reaches code that damage to a library-written file cannot. The first run of this found a
    /// content types part whose root was the wrong element.
    /// </summary>
    [Theory]
    [MemberData(nameof(CorpusWorkbooks))]
    public async Task MutatedExcelWorkbooks_FailCleanlyOrWork(string name)
    {
        var seed = File.ReadAllBytes(Path.Combine(CorpusFolder(), name + ".xlsx"));
        string[] parts;
        string[] sheets;
        using (var zip = new ZipArchive(new MemoryStream(seed), ZipArchiveMode.Read))
            parts = zip.Entries.Select(e => e.FullName).Where(n => n.EndsWith(".xml", StringComparison.Ordinal) || n.EndsWith(".rels", StringComparison.Ordinal)).ToArray();
        using (var reader = ExcelReader.Open(new MemoryStream(seed))) sheets = reader.SheetNames.ToArray();

        var seedHash = 2166136261u;
        foreach (var ch in name) seedHash = (seedHash ^ ch) * 16777619u;
        var rng = new Random(unchecked((int)seedHash) ^ 2468 ^ SeedOffset());
        string[] junk = { "<", ">", "\"", "999999999999", "-1", "ZZZZZ1", "&amp;", "<row>", "</c>", "t=\"s\"", "s=\"99999\"", "r=\"A0\"", "1E400", "NaN", "<a><a><a>", "r=\"-3\"", "ref=\"A1:\"", "count=\"-1\"", "<![CDATA[", "]]>", "numFmtId=\"-9\"", "<sheet/>", "r:id=\"rId99\"", "Target=\"../../../x\"", "Target=\"/\"", "<Override/>", "<Relationship/>", "si=\"7\"", "cacheId=\"99\"" };

        var options = new PackageOptions { MaxEditableCells = 1_000_000 };
        var actions = new (string Name, Action<byte[]> Run)[]
        {
            ("read everything", bytes =>
            {
                using var reader = ExcelReader.Open(new MemoryStream(bytes));
                foreach (var sheet in reader.SheetNames)
                {
                    foreach (var row in reader.ReadRows(sheet))
                        for (var c = 0; c < row.Values.Count; c++) _ = row.GetString(c);
                    _ = reader.ReadMergedRanges(sheet).Count;
                    _ = reader.ReadHyperlinks(sheet).Count;
                }
            }),
            ("edit every sheet", bytes =>
            {
                using var package = ExcelPackage.Edit(new MemoryStream(bytes), Stream.Null, leaveOpen: true, options);
                foreach (var sheet in sheets) package.Workbook.Worksheets.Edit(sheet).Cells[3, 12].Value = "changed";
                package.Save();
            }),
            ("fill the first sheet", bytes =>
            {
                using var package = ExcelPackage.OpenTemplate(new MemoryStream(bytes), Stream.Null, leaveOpen: true, options);
                var sheet = package.Workbook.Worksheets.Open(sheets[0]);
                for (var i = 0; i < 3; i++) sheet.Row(900 + i).WriteCell("x").WriteCell(i);
                package.Save();
            }),
            ("copy untouched", bytes =>
            {
                using var package = ExcelPackage.OpenTemplate(new MemoryStream(bytes), Stream.Null, leaveOpen: true, options);
                package.Save();
            }),
        };

        for (int iteration = 0, iterations = Iterations(60); iteration < iterations; iteration++)
        {
            var part = parts[rng.Next(parts.Length)];
            var mutated = MutatePart(seed, part, text =>
            {
                var sb = new StringBuilder(text);
                var edits = 1 + rng.Next(4);
                for (var e = 0; e < edits && sb.Length > 0; e++)
                {
                    var at = rng.Next(sb.Length);
                    switch (rng.Next(4))
                    {
                        case 0: sb.Remove(at, Math.Min(1 + rng.Next(40), sb.Length - at)); break;
                        case 1: sb.Insert(at, junk[rng.Next(junk.Length)]); break;
                        case 2:
                            // Move a stretch elsewhere: this is what put a child element where a root belongs.
                            var length = Math.Min(1 + rng.Next(120), sb.Length - at);
                            var piece = sb.ToString(at, length);
                            sb.Remove(at, length);
                            sb.Insert(rng.Next(sb.Length + 1), piece);
                            break;
                        default: sb[at] = (char)rng.Next(32, 127); break;
                    }
                }
                return sb.ToString();
            });

            foreach (var (action, run) in actions)
            {
                try
                {
                    var task = Task.Run(() => run(mutated));
                    Assert.True(await Task.WhenAny(task, Task.Delay(TimeSpan.FromSeconds(20))) == task, $"'{action}' hung on iteration {iteration} of {name} ({part}).");
                    await task;
                }
                catch (Exception ex) when (IsExpected(ex))
                {
                }
                catch (ArgumentException ex) when (ex.ParamName == "name" && ex.Message.Contains("no sheet named"))
                {
                    // The damage renamed or removed a sheet the test asks for by its original name.
                }
                catch (Exception ex) when (ex is not Xunit.Sdk.XunitException)
                {
                    Assert.Fail($"'{action}', iteration {iteration} of {name}, damage in {part}: unexpected {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
                }
            }
        }
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
