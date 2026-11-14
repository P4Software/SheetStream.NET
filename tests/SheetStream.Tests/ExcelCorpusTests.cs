using System.Globalization;
using SheetStream.Internal;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// Reads workbooks that Excel wrote and compares every cell with what Excel itself says the cell holds.
/// Most other tests read files this library produced, which shows that its two halves agree with each other;
/// these show that the reader agrees with Excel, on files that carry everything Excel puts in a workbook.
/// </summary>
/// <remarks>
/// The corpus lives in <c>tests/corpus/excel</c> and is checked in, so no Excel is needed to run the tests.
/// <c>tests/corpus/build-excel-corpus.ps1</c> regenerates it on a machine that has Excel: it builds each
/// workbook through automation and asks Excel for the value, the displayed date, the merges and the links,
/// which it writes beside the workbook as <c>name.expected.tsv</c>.
/// </remarks>
public class ExcelCorpusTests
{
    // The project copies tests/corpus/excel beside the test binaries.
    private static string Folder() => Path.Combine(AppContext.BaseDirectory, "corpus", "excel");

    public static IEnumerable<object[]> Workbooks()
        => Directory.GetFiles(Folder(), "*.xlsx").Select(f => new object[] { Path.GetFileNameWithoutExtension(f) }).OrderBy(a => (string)a[0]);

    /// <summary>What Excel reported for one workbook.</summary>
    private sealed class Expected
    {
        public List<(string Name, string State)> Sheets { get; } = new();
        public Dictionary<(string Sheet, string Address), (string Kind, string Value)> Cells { get; } = new();
        public Dictionary<string, HashSet<string>> Merges { get; } = new();
        public Dictionary<string, List<(string Address, string Url, string Location, string Tip)>> Links { get; } = new();
    }

    private static string Unescape(string text)
    {
        var sb = new System.Text.StringBuilder();
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] != '\\' || i + 1 == text.Length) { sb.Append(text[i]); continue; }
            i++;
            sb.Append(text[i] switch { 't' => '\t', 'r' => '\r', 'n' => '\n', _ => text[i] });
        }
        return sb.ToString();
    }

    private static Expected Load(string name)
    {
        var expected = new Expected();
        foreach (var raw in File.ReadAllLines(Path.Combine(Folder(), name + ".expected.tsv")))
        {
            if (raw.Length == 0) continue;
            var parts = raw.Split('\t');
            var (sheet, kind, address, value) = (parts[0], parts[1], parts[2], parts.Length > 3 ? parts[3] : "");
            switch (kind)
            {
                case "meta":
                    expected.Sheets.Add((sheet, value));
                    break;
                case "merge":
                    if (!expected.Merges.TryGetValue(sheet, out var merges)) expected.Merges[sheet] = merges = new HashSet<string>();
                    merges.Add(address);
                    break;
                case "link":
                    if (!expected.Links.TryGetValue(sheet, out var links)) expected.Links[sheet] = links = new List<(string, string, string, string)>();
                    var fields = value.Split('|');
                    links.Add((address, Unescape(fields[0]), Unescape(fields[1]), Unescape(fields[2])));
                    break;
                default:
                    expected.Cells[(sheet, address)] = (kind, kind == "s" ? Unescape(value) : value);
                    break;
            }
        }
        return expected;
    }

    private static string A1(int row, int column) => CellReference.ToA1(row, column);

    /// <summary>Every non-empty cell of every sheet, as the reader returns it.</summary>
    private static Dictionary<(string Sheet, string Address), object> ReadCells(ExcelReader reader)
    {
        var cells = new Dictionary<(string, string), object>();
        foreach (var sheet in reader.SheetNames)
            foreach (var row in reader.ReadRows(sheet))
                for (var c = 0; c < row.Values.Count; c++)
                    if (row.Values[c] != null) cells[(sheet, A1(row.Index, c + 1))] = row.Values[c]!;
        return cells;
    }

    /// <summary>Compares what the reader returned with what Excel reported, cell by cell, and lists every difference.</summary>
    private static List<string> Compare(Expected expected, Dictionary<(string Sheet, string Address), object> actual, bool date1904)
    {
        var problems = new List<string>();
        foreach (var pair in expected.Cells)
        {
            var (kind, value) = pair.Value;
            var where = $"{pair.Key.Sheet}!{pair.Key.Address}";
            actual.TryGetValue(pair.Key, out var got);
            switch (kind)
            {
                case "s":
                    // A cell holding an empty text is as good as empty to a reader.
                    if (value.Length == 0 ? got != null && !Equals(got, "") : !Equals(got, value)) problems.Add($"{where}: Excel has text '{Short(value)}', reader has {Show(got)}");
                    break;
                case "b":
                    if (!Equals(got, value == "1")) problems.Add($"{where}: Excel has {(value == "1" ? "TRUE" : "FALSE")}, reader has {Show(got)}");
                    break;
                case "e":
                    if (!(got is ExcelErrorValue error) || error.Code != value) problems.Add($"{where}: Excel has error {value}, reader has {Show(got)}");
                    break;
                case "d":
                    var shown = DateTime.ParseExact(value, "yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
                    if (!(got is DateTime date) || Math.Abs((date - shown).TotalMilliseconds) > 1) problems.Add($"{where}: Excel shows the date {value}, reader has {Show(got)}");
                    break;
                case "t":
                    var days = double.Parse(value, CultureInfo.InvariantCulture);
                    if (!(got is TimeSpan span) || Math.Abs(span.TotalDays - days) > 1e-9) problems.Add($"{where}: Excel has an elapsed time of {value} days, reader has {Show(got)}");
                    break;
                case "n":
                    var number = double.Parse(value, CultureInfo.InvariantCulture);
                    // Excel's automation reports a time-of-day format (hh:mm) as a plain number; the reader gives
                    // such a cell as a DateTime. Both are the same serial, so that is accepted here, and
                    // FormatsDecideWhatIsADate pins which formats may do it.
                    var same = got is double d ? d.Equals(number) || Math.Abs(d - number) <= Math.Abs(number) * 1e-15
                        : got is DateTime time && Math.Abs(ExcelDate.ToSerial(time) - (date1904 ? number + 1462 : number)) < 1e-8;
                    if (!same) problems.Add($"{where}: Excel has the number {value}, reader has {Show(got)}");
                    break;
            }
        }
        foreach (var pair in actual)
            if (!expected.Cells.ContainsKey(pair.Key)) problems.Add($"{pair.Key.Sheet}!{pair.Key.Address}: reader has {Show(pair.Value)}, Excel has nothing there");
        return problems;

        static string Short(string text) => text.Length > 40 ? text.Substring(0, 40) + "..." : text;
        static string Show(object? value) => value switch
        {
            null => "nothing",
            string s => $"text '{Short(s)}'",
            DateTime dt => $"the date {dt:yyyy-MM-dd HH:mm:ss.fff}",
            TimeSpan ts => $"an elapsed time of {ts.TotalDays.ToString("R", CultureInfo.InvariantCulture)} days",
            double x => $"the number {x.ToString("R", CultureInfo.InvariantCulture)}",
            _ => $"{value.GetType().Name} {value}",
        };
    }

    private static void AssertNoProblems(string what, List<string> problems)
        => Assert.True(problems.Count == 0, $"{what}: {problems.Count} difference(s) from Excel.\n" + string.Join("\n", problems.Take(25)));

    [Theory]
    [MemberData(nameof(Workbooks))]
    public void Reader_AgreesWithExcel_OnEveryCell(string name)
    {
        var expected = Load(name);
        using var reader = ExcelReader.Open(Path.Combine(Folder(), name + ".xlsx"));

        Assert.Equal(expected.Sheets.Select(s => s.Name), reader.SheetNames);
        foreach (var (sheet, state) in expected.Sheets)
            Assert.Equal(state, reader.GetVisibility(sheet) switch { SheetVisibility.Hidden => "hidden", SheetVisibility.VeryHidden => "veryHidden", _ => "visible" });

        AssertNoProblems(name, Compare(expected, ReadCells(reader), name == "dates-1904"));
    }

    [Theory]
    [MemberData(nameof(Workbooks))]
    public async Task AsyncReader_AgreesWithExcel_OnEveryCell(string name)
    {
        var expected = Load(name);
        using var reader = await ExcelReader.OpenAsync(new MemoryStream(File.ReadAllBytes(Path.Combine(Folder(), name + ".xlsx"))));
        var cells = new Dictionary<(string, string), object>();
        foreach (var sheet in reader.SheetNames)
            await foreach (var row in reader.ReadRowsAsync(sheet))
                for (var c = 0; c < row.Values.Count; c++)
                    if (row.Values[c] != null) cells[(sheet, A1(row.Index, c + 1))] = row.Values[c]!;
        AssertNoProblems(name, Compare(expected, cells, name == "dates-1904"));
    }

    [Theory]
    [MemberData(nameof(Workbooks))]
    public void MergesAndLinks_AgreeWithExcel(string name)
    {
        var expected = Load(name);
        using var reader = ExcelReader.Open(Path.Combine(Folder(), name + ".xlsx"));
        foreach (var (sheet, _) in expected.Sheets)
        {
            var merges = reader.ReadMergedRanges(sheet).Select(m => m.Range).ToHashSet();
            Assert.Equal(expected.Merges.TryGetValue(sheet, out var wanted) ? wanted.OrderBy(x => x) : Enumerable.Empty<string>(), merges.OrderBy(x => x));

            var links = reader.ReadHyperlinks(sheet).Select(l => (l.Address, l.Url ?? "", l.Location ?? "", l.Tooltip ?? "")).OrderBy(l => l.Address).ToList();
            var wantedLinks = (expected.Links.TryGetValue(sheet, out var list) ? list : new List<(string, string, string, string)>()).OrderBy(l => l.Item1).ToList();
            Assert.Equal(wantedLinks, links);
        }
    }

    [Fact]
    public void FormatsDecideWhatIsADate_TheWayExcelDoes()
    {
        // Column A holds a number under the format named in column B. The formats below contain the letters
        // d, m, y, h or s only inside quotes, after a backslash, or not at all: Excel shows a plain number for
        // them, and so must the reader. Every other format in the sheet is a date, a time or an elapsed time.
        string[] plainNumbers = { "#,##0.00", "0.0%", "0.00E+00", "# ?/?", "$#,##0.00", "0.0 \"days\"", "0 \"months\"", "\"d\"0", "[Red]0;[Blue]-0", "General", "0 \\d" };
        using var reader = ExcelReader.Open(Path.Combine(Folder(), "formats.xlsx"));
        var seen = new HashSet<string>();
        foreach (var row in reader.ReadRows("Formats"))
        {
            if (!(row[1] is string format)) continue;
            seen.Add(format);
            if (plainNumbers.Contains(format)) Assert.True(row[0] is double, $"'{format}' must read as a number, got {row[0]?.GetType().Name}");
            else if (format.Contains("[h]") || format.Contains("[m]") || format.Contains("[s]")) Assert.True(row[0] is TimeSpan, $"'{format}' must read as an elapsed time, got {row[0]?.GetType().Name}");
            else Assert.True(row[0] is DateTime, $"'{format}' must read as a date or time, got {row[0]?.GetType().Name}");
        }
        Assert.All(plainNumbers, format => Assert.Contains(format, seen));
    }

    /// <summary>
    /// A workbook that is opened and saved again without a change, or with one cell added, must still hold
    /// what Excel put in it. This is where an edit or a template copy would lose data: a shared formula, a
    /// rich text, a date in the 1904 system, a cell at the edge of the grid.
    /// </summary>
    [Theory]
    [MemberData(nameof(Workbooks))]
    public void EditAndTemplateRoundTrips_KeepEveryValueExcelWrote(string name)
    {
        var expected = Load(name);
        var original = File.ReadAllBytes(Path.Combine(Folder(), name + ".xlsx"));
        var options = new PackageOptions { MaxEditableCells = 1_000_000 };

        // 1. Template copy, nothing touched.
        var copied = new MemoryStream();
        using (var package = ExcelPackage.OpenTemplate(new MemoryStream(original), copied, leaveOpen: true, options)) package.Save();
        using (var reader = ExcelReader.Open(new MemoryStream(copied.ToArray())))
            AssertNoProblems(name + " after an untouched template copy", Compare(expected, ReadCells(reader), name == "dates-1904"));

        // 2. Every sheet opened for editing and saved unchanged.
        var edited = new MemoryStream();
        using (var package = ExcelPackage.Edit(new MemoryStream(original), edited, leaveOpen: true, options))
        {
            foreach (var (sheet, _) in expected.Sheets) package.Workbook.Worksheets.Edit(sheet);
            package.Save();
        }
        using (var reader = ExcelReader.Open(new MemoryStream(edited.ToArray())))
        {
            Assert.Equal(expected.Sheets.Select(s => s.Name), reader.SheetNames);
            AssertNoProblems(name + " after editing every sheet without a change", Compare(expected, ReadCells(reader), name == "dates-1904"));
            foreach (var (sheet, _) in expected.Sheets)
            {
                var merges = reader.ReadMergedRanges(sheet).Select(m => m.Range).OrderBy(x => x);
                Assert.Equal(expected.Merges.TryGetValue(sheet, out var wanted) ? wanted.OrderBy(x => x) : Enumerable.Empty<string>(), merges);
            }
        }

        // 3. One cell added far from the data, on the first sheet: everything else must be as it was.
        var changed = new MemoryStream();
        var first = expected.Sheets[0].Name;
        using (var package = ExcelPackage.Edit(new MemoryStream(original), changed, leaveOpen: true, options))
        {
            package.Workbook.Worksheets.Edit(first).Cells["J777"].Value = "added";
            package.Save();
        }
        using (var reader = ExcelReader.Open(new MemoryStream(changed.ToArray())))
        {
            var cells = ReadCells(reader);
            Assert.Equal("added", cells[(first, "J777")]);
            if (!expected.Cells.ContainsKey((first, "J777"))) cells.Remove((first, "J777"));
            AssertNoProblems(name + " after adding one cell", Compare(expected, cells, name == "dates-1904"));
        }
    }
}
