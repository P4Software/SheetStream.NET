using System.IO.Compression;
using System.Text;
using System.Xml;
using SheetStream;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// Workbooks built to exhaust memory, and every way this library opens a workbook. Shared by the test project,
/// which checks what each combination does, and by the memory probe, which runs the same combinations in a
/// process that is not allowed more than a few tens of megabytes.
/// </summary>
/// <remarks>
/// A workbook is a zip archive, so each of these files is a few kilobytes on the wire and tens of megabytes
/// unpacked. When a new way of opening a workbook is added to the library, add it to <see cref="Entries"/>;
/// when a new kind of hostile file is thought of, add it to <see cref="Shapes"/>. Every shape then runs through
/// every entry point without anyone having to remember the pairing.
/// </remarks>
internal static class HostileWorkbooks
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string SheetPart = "xl/worksheets/sheet1.xml";
    private const string SheetHead = "<worksheet xmlns=\"" + Main + "\"><sheetData>";
    private const string SheetTail = "</sheetData></worksheet>";

    /// <summary>Characters of hostile content in each file: 20 million, about 40 MB once it is a string.</summary>
    public const int Size = 20_000_000;

    /// <summary>Limits tightened so the test files can be small; the defaults are 50 times larger.</summary>
    public static ReaderOptions Limits() => new ReaderOptions { MaxCellTextLength = 20_000, MaxLoadedPartBytes = 2_000_000 };
    // ------------------------------------------------------------------ the files

    /// <summary>A genuine workbook with shared strings, a style, a table and a formula, to be damaged part by part.</summary>
    private static byte[] Seed() => ExcelPackage.CreateBytes(wb =>
    {
        var bold = wb.Styles.Create(o => o.Font.Bold = true);
        var s = wb.Worksheets.Add("Data");
        s.AddTable("Items", "A1:B3");
        s.Row().WriteCell("Name", bold).WriteCell("Qty", bold);
        s.Row().WriteCell("a").WriteCell(1);
        s.Row().WriteCell("b").WriteCell(2);
    });

    public delegate void Body(TextWriter writer);

    // Declared before the shapes, which capture it while the class is being initialised.
    private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);

    /// <summary>Writes <paramref name="unit"/> over and over until <paramref name="chars"/> characters are out.</summary>
    private static void Repeat(TextWriter w, string unit, int chars = Size)
    {
        var block = string.Concat(Enumerable.Repeat(unit, Math.Max(1, 32_768 / unit.Length)));
        for (long done = 0; done < chars; done += block.Length) w.Write(block);
    }

    /// <summary>Every hostile shape by name: the part it replaces, its encoding and what it writes there.</summary>
    public static readonly Dictionary<string, (string Part, Encoding Encoding, Body Write)> Shapes = new()
    {
        // --- one oversized token in a sheet
        ["sheet: long text in a cell"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>"); Repeat(w, "A"); w.Write("</t></is></c></row>" + SheetTail); }),
        ["sheet: long cached value"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\" t=\"str\"><f>A2</f><v>"); Repeat(w, "A"); w.Write("</v></c></row>" + SheetTail); }),
        ["sheet: long formula"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\"><f>"); Repeat(w, "1+"); w.Write("1</f><v>1</v></c></row>" + SheetTail); }),
        ["sheet: long attribute"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\" x=\""); Repeat(w, "A"); w.Write("\"><v>1</v></c></row>" + SheetTail); }),
        ["sheet: long row attribute"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\" x=\""); Repeat(w, "A"); w.Write("\"><c r=\"A1\"><v>1</v></c></row>" + SheetTail); }),
        ["sheet: long element name"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><"); Repeat(w, "A"); w.Write("/></row>" + SheetTail); }),
        ["sheet: very many attributes"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\""); for (var i = 0; i < Size / 12; i++) w.Write(" a" + i + "=\"1\""); w.Write("><v>1</v></c></row>" + SheetTail); }),
        ["sheet: entity references"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>"); Repeat(w, "&amp;"); w.Write("</t></is></c></row>" + SheetTail); }),
        ["sheet: long text before the rows"] = (SheetPart, Utf8, w => { w.Write("<worksheet xmlns=\"" + Main + "\"><sheetPr><tabColor rgb=\""); Repeat(w, "F"); w.Write("\"/></sheetPr><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row>" + SheetTail); }),
        ["sheet: long text after the rows"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData><pageSetup orientation=\""); Repeat(w, "A"); w.Write("\"/></worksheet>"); }),

        // --- CDATA, where '<' is ordinary text
        ["cdata: one section as cell text"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t><![CDATA["); Repeat(w, "<"); w.Write("]]></t></is></c></row>" + SheetTail); }),
        ["cdata: one section outside any cell"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><![CDATA["); Repeat(w, "<"); w.Write("]]><c r=\"A1\"><v>1</v></c></row>" + SheetTail); }),
        ["cdata: one section of closing brackets"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><![CDATA["); Repeat(w, "]"); w.Write("x]]><c r=\"A1\"><v>1</v></c></row>" + SheetTail); }),
        ["cdata: one section before the rows"] = (SheetPart, Utf8, w => { w.Write("<worksheet xmlns=\"" + Main + "\"><![CDATA["); Repeat(w, "<a>"); w.Write("]]><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row>" + SheetTail); }),
        ["cdata: many sections in one cell"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>"); Repeat(w, "<![CDATA[" + new string('<', 5000) + "]]>"); w.Write("</t></is></c></row>" + SheetTail); }),
        ["cdata: in the shared strings"] = ("xl/sharedStrings.xml", Utf8, w => { w.Write("<sst xmlns=\"" + Main + "\"><si><t>Name</t></si><si><t><![CDATA["); Repeat(w, "<"); w.Write("]]></t></si></sst>"); }),

        // --- text the parser skips
        ["skipped: long comment"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><!--"); Repeat(w, "<"); w.Write("--><c r=\"A1\"><v>1</v></c></row>" + SheetTail); }),
        ["skipped: long processing instruction"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><?x "); Repeat(w, "<"); w.Write("?><c r=\"A1\"><v>1</v></c></row>" + SheetTail); }),

        // --- encodings in which bytes and characters disagree
        ["utf-16: attribute of U+3C3C"] = (SheetPart, Encoding.Unicode, w => { w.Write("<?xml version=\"1.0\" encoding=\"utf-16\"?>" + SheetHead + "<row r=\"1\"><c r=\"A1\" x=\""); Repeat(w, "㰼"); w.Write("\"><v>1</v></c></row>" + SheetTail); }),
        ["utf-16: cell text of U+3C3C"] = (SheetPart, Encoding.Unicode, w => { w.Write("<?xml version=\"1.0\" encoding=\"utf-16\"?>" + SheetHead + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>"); Repeat(w, "㰼"); w.Write("</t></is></c></row>" + SheetTail); }),
        ["utf-16 big-endian: attribute of U+3C3C"] = (SheetPart, Encoding.BigEndianUnicode, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\" x=\""); Repeat(w, "㰼"); w.Write("\"><v>1</v></c></row>" + SheetTail); }),

        // --- nesting
        ["nesting: in the rows"] = (SheetPart, Utf8, w => { w.Write(SheetHead); Repeat(w, "<a>"); }),
        ["nesting: inside a cell"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\"><f>"); Repeat(w, "<a>"); }),
        ["nesting: before the rows"] = (SheetPart, Utf8, w => { w.Write("<worksheet xmlns=\"" + Main + "\"><sheetPr>"); Repeat(w, "<a>"); }),
        ["nesting: after the rows"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>"); Repeat(w, "<a>"); }),
        ["nesting: in the shared strings"] = ("xl/sharedStrings.xml", Utf8, w => { w.Write("<sst xmlns=\"" + Main + "\"><si>"); Repeat(w, "<r>"); }),
        ["nesting: in the styles"] = ("xl/styles.xml", Utf8, w => { w.Write("<styleSheet xmlns=\"" + Main + "\"><fonts>"); Repeat(w, "<font>"); }),
        ["nesting: in the workbook"] = ("xl/workbook.xml", Utf8, w => { w.Write("<workbook xmlns=\"" + Main + "\">"); Repeat(w, "<a>"); }),

        // --- many small things, where no single token is large
        ["many: cells without references in one row"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\">"); Repeat(w, "<c><v>1</v></c>"); w.Write("</row>" + SheetTail); }),
        ["many: empty cells without references in one row"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\">"); Repeat(w, "<c/>"); w.Write("</row>" + SheetTail); }),
        ["many: shared strings"] = ("xl/sharedStrings.xml", Utf8, w => { w.Write("<sst xmlns=\"" + Main + "\">"); Repeat(w, "<si><t>x</t></si>"); w.Write("</sst>"); }),
        ["many: rich text runs in one shared string"] = ("xl/sharedStrings.xml", Utf8, w => { w.Write("<sst xmlns=\"" + Main + "\"><si>"); Repeat(w, "<r><t>x</t></r>"); w.Write("</si></sst>"); }),
        ["many: cell formats in the styles"] = ("xl/styles.xml", Utf8, w => { w.Write("<styleSheet xmlns=\"" + Main + "\"><fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs>"); Repeat(w, "<xf numFmtId=\"0\"/>"); w.Write("</cellXfs></styleSheet>"); }),
        ["many: number formats in the styles"] = ("xl/styles.xml", Utf8, w => { w.Write("<styleSheet xmlns=\"" + Main + "\"><numFmts>"); for (var i = 0; i < Size / 40; i++) w.Write("<numFmt numFmtId=\"" + (200 + i) + "\" formatCode=\"0.0\"/>"); w.Write("</numFmts><cellXfs><xf numFmtId=\"0\"/></cellXfs></styleSheet>"); }),
        ["many: defined names in the workbook"] = ("xl/workbook.xml", Utf8, w => { w.Write("<workbook xmlns=\"" + Main + "\" xmlns:r=\"" + Rel + "\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets><definedNames>"); Repeat(w, "<definedName name=\"n\">Data!$A$1</definedName>"); w.Write("</definedNames></workbook>"); }),
        ["many: sheets in the workbook"] = ("xl/workbook.xml", Utf8, w => { w.Write("<workbook xmlns=\"" + Main + "\" xmlns:r=\"" + Rel + "\"><sheets>"); Repeat(w, "<sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/>"); w.Write("</sheets></workbook>"); }),
        ["many: workbook relationships"] = ("xl/_rels/workbook.xml.rels", Utf8, w => { w.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"); Repeat(w, "<Relationship Id=\"rId1\" Type=\"" + Rel + "/worksheet\" Target=\"worksheets/sheet1.xml\"/>"); w.Write("</Relationships>"); }),
        ["many: sheet relationships"] = ("xl/worksheets/_rels/sheet1.xml.rels", Utf8, w => { w.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"); Repeat(w, "<Relationship Id=\"rId9\" Type=\"" + Rel + "/hyperlink\" Target=\"http://example.com/\" TargetMode=\"External\"/>"); w.Write("</Relationships>"); }),
        ["many: content types"] = ("[Content_Types].xml", Utf8, w => { w.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"); Repeat(w, "<Default Extension=\"png\" ContentType=\"image/png\"/>"); w.Write("</Types>"); }),
        ["many: rows"] = (SheetPart, Utf8, w => { w.Write(SheetHead); Repeat(w, "<row><c><v>1</v></c></row>"); w.Write(SheetTail); }),
        ["many: empty rows with attributes"] = (SheetPart, Utf8, w => { w.Write(SheetHead); for (var i = 1; i <= Size / 40; i++) w.Write("<row r=\"" + (i % 1_000_000 + 1) + "\" ht=\"15\" customHeight=\"1\"/>"); w.Write(SheetTail); }),
        ["many: merged ranges"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData><mergeCells>"); Repeat(w, "<mergeCell ref=\"A1:B2\"/>"); w.Write("</mergeCells></worksheet>"); }),
        ["many: hyperlinks with long tooltips"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData><hyperlinks>"); Repeat(w, "<hyperlink ref=\"A1\" location=\"Data!A1\" tooltip=\"" + new string('t', 30_000) + "\"/>"); w.Write("</hyperlinks></worksheet>"); }),
        ["many: column definitions"] = (SheetPart, Utf8, w => { w.Write("<worksheet xmlns=\"" + Main + "\"><cols>"); Repeat(w, "<col min=\"1\" max=\"1\" width=\"9\"/>"); w.Write("</cols><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row>" + SheetTail); }),
        ["many: kinds of collection in the styles"] = ("xl/styles.xml", Utf8, w => { w.Write("<styleSheet xmlns=\"" + Main + "\">"); for (var i = 0; i < Size / 12; i++) w.Write("<k" + i + "/>"); w.Write("<cellXfs><xf numFmtId=\"0\"/></cellXfs></styleSheet>"); }),

        // --- names, which the parser keeps for as long as it reads the part
        ["names: distinct elements in the rows"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\">"); for (var i = 0; i < Size / 10; i++) w.Write("<k" + i + "/>"); w.Write("</row>" + SheetTail); }),
        ["names: distinct elements after the rows"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>"); for (var i = 0; i < Size / 10; i++) w.Write("<k" + i + "/>"); w.Write("</worksheet>"); }),
        ["names: distinct attributes on many cells"] = (SheetPart, Utf8, w => { w.Write(SheetHead); for (var i = 0; i < Size / 40; i++) w.Write("<row><c k" + i + "=\"1\"><v>1</v></c></row>"); w.Write(SheetTail); }),
        ["names: distinct prefixes"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\">"); for (var i = 0; i < Size / 40; i++) w.Write("<p" + i + ":x xmlns:p" + i + "=\"urn:n" + i + "\"/>"); w.Write("</row>" + SheetTail); }),
        ["names: a few very long ones"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\">"); for (var i = 0; i < Size / 50_000; i++) w.Write("<k" + i + new string('n', 50_000) + "/>"); w.Write("</row>" + SheetTail); }),
        ["names: distinct elements in the shared strings"] = ("xl/sharedStrings.xml", Utf8, w => { w.Write("<sst xmlns=\"" + Main + "\"><si><t>Name</t></si><si>"); for (var i = 0; i < Size / 10; i++) w.Write("<k" + i + "/>"); w.Write("</si></sst>"); }),

        // --- shapes that cost time, or that cost what the caller collects (found by a separate adversarial review pass)
        ["time: relationship target of very many dot segments"] = ("xl/_rels/workbook.xml.rels", Utf8, w => { w.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"" + Rel + "/worksheet\" Target=\""); Repeat(w, "./", 15_000); w.Write("worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"" + Rel + "/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId3\" Type=\"" + Rel + "/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>"); }),
        ["time: one long number format used by many cell formats"] = ("xl/styles.xml", Utf8, w => { w.Write("<styleSheet xmlns=\"" + Main + "\"><numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\""); Repeat(w, "0", 15_000); w.Write("\"/></numFmts><fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs>"); Repeat(w, "<xf numFmtId=\"164\"/>", 1_500_000); w.Write("</cellXfs></styleSheet>"); }),
        ["collect: rows with one cell in the last column"] = (SheetPart, Utf8, w => { w.Write(SheetHead); for (var i = 1; i <= 20_000; i++) w.Write("<row r=\"" + i + "\"><c r=\"XFD" + i + "\"><v>1</v></c></row>"); w.Write(SheetTail); }),
        ["collect: header of identical names"] = (SheetPart, Utf8, w => { w.Write(SheetHead + "<row r=\"1\">"); Repeat(w, "<c t=\"inlineStr\"><is><t>a</t></is></c>", 16_000 * 38); w.Write("</row><row r=\"2\"><c r=\"A2\"><v>1</v></c></row>" + SheetTail); }),
        ["structure: two sheets with one name"] = ("xl/workbook.xml", Utf8, w => w.Write("<workbook xmlns=\"" + Main + "\" xmlns:r=\"" + Rel + "\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Data\" sheetId=\"2\" r:id=\"rId1\"/></sheets></workbook>")),
        ["structure: workbook that is an empty element"] = ("xl/workbook.xml", Utf8, w => w.Write("<workbook xmlns=\"" + Main + "\"/>")),
        ["structure: styles that are an empty element"] = ("xl/styles.xml", Utf8, w => w.Write("<styleSheet xmlns=\"" + Main + "\"/>")),

        // --- numbers that index something
        ["index: row number at the top of the integer range"] = (SheetPart, Utf8, w => w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2147483647\"><c><v>1</v></c></row>" + SheetTail)),
        ["index: negative row number"] = (SheetPart, Utf8, w => w.Write(SheetHead + "<row r=\"-5\"><c><v>1</v></c></row>" + SheetTail)),
        ["index: row number zero"] = (SheetPart, Utf8, w => w.Write(SheetHead + "<row r=\"0\"><c><v>1</v></c></row>" + SheetTail)),
        ["index: style far past the last one"] = (SheetPart, Utf8, w => w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\" s=\"2000000000\"><v>1</v></c><c r=\"B1\" s=\"-7\"><v>1</v></c></row>" + SheetTail)),
        ["index: shared string far past the last one"] = (SheetPart, Utf8, w => w.Write(SheetHead + "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>2000000000</v></c><c r=\"B1\" t=\"s\"><v>-3</v></c></row>" + SheetTail)),
        ["index: column far past the last one"] = (SheetPart, Utf8, w => w.Write(SheetHead + "<row r=\"1\"><c r=\"ZZZZZZZZZ1\"><v>1</v></c></row>" + SheetTail)),
        ["index: table range that is not a range"] = ("xl/tables/table1.xml", Utf8, w => w.Write("<table xmlns=\"" + Main + "\" id=\"1\" name=\"Items\" displayName=\"Items\" ref=\"ZZZZZZZZ9999999999:A1\" totalsRowCount=\"-4\"><autoFilter ref=\"\"/><tableColumns count=\"2000000000\"><tableColumn id=\"1\" name=\"Name\"/></tableColumns></table>")),

        ["many: table columns"] = ("xl/tables/table1.xml", Utf8, w => { w.Write("<table xmlns=\"" + Main + "\" id=\"1\" name=\"Items\" displayName=\"Items\" ref=\"A1:B3\"><autoFilter ref=\"A1:B3\"/><tableColumns count=\"2\">"); Repeat(w, "<tableColumn id=\"1\" name=\"Name\"/>"); w.Write("</tableColumns></table>"); }),
    };

    /// <summary>The seed workbook with one part replaced by the named shape: a few hundred kilobytes at most.</summary>
    public static byte[] Build(string shape)
    {
        var (part, encoding, write) = Shapes[shape];
        var input = new MemoryStream(Seed());
        var output = new MemoryStream();
        using (var from = new ZipArchive(input, ZipArchiveMode.Read))
        using (var to = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            var replaced = false;
            foreach (var entry in from.Entries)
            {
                using var target = to.CreateEntry(entry.FullName, CompressionLevel.Fastest).Open();
                if (!string.Equals(entry.FullName, part, StringComparison.OrdinalIgnoreCase))
                {
                    using var source = entry.Open();
                    source.CopyTo(target);
                    continue;
                }
                replaced = true;
                using var writer = new StreamWriter(target, encoding, 1 << 16);
                write(writer);
            }
            if (!replaced) throw new InvalidOperationException($"The seed workbook has no part '{part}'.");
        }
        return output.ToArray();
    }
    // ------------------------------------------------------------------ the ways in

    private static PackageOptions PackageLimits() => new PackageOptions { SourceLimits = Limits() };

    /// <summary>Every way of opening a workbook, by name. Each takes the file and does a typical piece of work with it.</summary>
    public static readonly Dictionary<string, Func<byte[], Task>> Entries = new()
    {
        ["reader: rows"] = bytes =>
        {
            using var reader = ExcelReader.Open(new MemoryStream(bytes), false, Limits());
            foreach (var name in reader.SheetNames.Take(3))
                foreach (var row in reader.ReadRows(name)) _ = row.Values.Count;
            return Task.CompletedTask;
        },
        ["reader: rows, async"] = async bytes =>
        {
            using var reader = await ExcelReader.OpenAsync(new MemoryStream(bytes), false, Limits());
            foreach (var name in reader.SheetNames.Take(3))
                await foreach (var row in reader.ReadRowsAsync(name)) _ = row.Values.Count;
        },
        ["reader: data table"] = bytes =>
        {
            using var reader = ExcelReader.Open(new MemoryStream(bytes), false, Limits());
            _ = reader.ReadDataTable("Data").Rows.Count;
            return Task.CompletedTask;
        },
        ["reader: merged ranges"] = bytes =>
        {
            using var reader = ExcelReader.Open(new MemoryStream(bytes), false, Limits());
            _ = reader.ReadMergedRanges(0).Count;
            return Task.CompletedTask;
        },
        ["reader: hyperlinks"] = bytes =>
        {
            using var reader = ExcelReader.Open(new MemoryStream(bytes), false, Limits());
            _ = reader.ReadHyperlinks(0).Count;
            return Task.CompletedTask;
        },
        ["edit: change a cell"] = bytes =>
        {
            using var package = ExcelPackage.Edit(new MemoryStream(bytes), Stream.Null, leaveOpen: true, PackageLimits());
            package.Workbook.Worksheets.Edit("Data").Cells["B2"].Value = 5;
            package.Save();
            return Task.CompletedTask;
        },
        ["edit: save untouched"] = bytes =>
        {
            using var package = ExcelPackage.Edit(new MemoryStream(bytes), Stream.Null, leaveOpen: true, PackageLimits());
            package.Save();
            return Task.CompletedTask;
        },
        ["template: fill rows"] = bytes =>
        {
            using var package = ExcelPackage.OpenTemplate(new MemoryStream(bytes), Stream.Null, leaveOpen: true, PackageLimits());
            var sheet = package.Workbook.Worksheets.Open("Data");
            for (var i = 0; i < 5; i++) sheet.Row(4 + i).WriteCell("x").WriteCell(i);
            package.Save();
            return Task.CompletedTask;
        },
        ["template: read a style"] = bytes =>
        {
            using var package = ExcelPackage.OpenTemplate(new MemoryStream(bytes), Stream.Null, leaveOpen: true, PackageLimits());
            _ = package.Workbook.Worksheets.Open("Data").TemplateStyle("B3");
            package.Save();
            return Task.CompletedTask;
        },
    };


    /// <summary>Exception types the documentation tells callers to catch around a file they did not write.</summary>
    public static bool IsDocumented(Exception ex) =>
        ex is InvalidDataException || ex is XmlException || ex is FormatException || ex is OverflowException;

    /// <summary>
    /// Runs one entry point on one file and says what happened: "Reads" when it completed, "Refused" when it
    /// threw a documented exception type, "No such sheet" when the file opened but lacks the sheet that was
    /// asked for. Anything else (a crash type, a hang) is thrown to the caller.
    /// </summary>
    public static async Task<(string Outcome, string Detail)> Run(byte[] workbook, string shape, string entry)
    {
        Exception? failure = null;
        // On its own task so a hang fails the run instead of blocking it.
        var task = Task.Run(async () =>
        {
            try { await Entries[entry](workbook); }
            catch (Exception ex) { failure = ex; }
        });
        if (await Task.WhenAny(task, Task.Delay(TimeSpan.FromSeconds(120))) != task)
            throw new TimeoutException($"'{entry}' did not finish within two minutes on '{shape}'.");
        if (failure == null) return ("Reads", "");
        if (failure is OutOfMemoryException) throw failure;
        // The file opened, but it does not have the sheet the caller asked for (the damage removed or renamed
        // it). That is the caller's argument not matching the file, reported as for any genuine workbook.
        if (failure is ArgumentException argument && (argument.ParamName == "name" || argument.ParamName == "sheetIndex" || argument.ParamName == "index"))
            return ("No such sheet", failure.Message);
        if (!IsDocumented(failure))
            throw new InvalidOperationException($"'{entry}' on '{shape}' threw {failure.GetType().Name}, which callers are not told to catch: {failure.Message}", failure);
        return ("Refused", failure.GetType().Name + ": " + failure.Message);
    }
}