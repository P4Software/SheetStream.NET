using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SheetStream;
using SheetStream.Reading;
using Xunit;

namespace SheetStream.NetFx.Tests
{
    /// <summary>
    /// End-to-end checks of the netstandard2.0 build on .NET Framework. They cover the paths where the two
    /// builds can differ: number formatting, the zip and XML stacks, temp files and file replacement.
    /// </summary>
    public class NetFrameworkTests
    {
        [Fact]
        public void RunsOnNetFramework()
        {
            // Guards the project itself: if it were retargeted to modern .NET these tests would prove nothing.
            Assert.StartsWith(".NET Framework", System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);
        }

        [Fact]
        public void WriteWithStylesTableAndAutoFit_ThenReadBack()
        {
            var bytes = ExcelPackage.CreateBytes(wb =>
            {
                var money = wb.Styles.Create(o => o.NumberFormat = "#,##0.00");
                var s = wb.Worksheets.Add("Data");
                s.AutoFitColumns();
                s.AddTable("Items", "A1:D1001");
                s.Row().WriteCells("Name", "Qty", "When", "Ok");
                for (var i = 0; i < 1000; i++)
                    s.Row().WriteCell("Item " + i).WriteCell(i * 1.5, money).WriteCell(new DateTime(2026, 1, 1).AddDays(i)).WriteCell(i % 2 == 0);
            });

            using (var reader = ExcelReader.Open(new MemoryStream(bytes)))
            {
                var rows = reader.ReadRows("Data").ToList();
                Assert.Equal(1001, rows.Count);
                Assert.Equal("Item 0", rows[1][0]);
                Assert.Equal(1.5, rows[2][1]);
                Assert.Equal(new DateTime(2026, 1, 2), rows[2][2]);
                Assert.Equal(true, rows[1][3]);
            }
        }

        [Fact]
        public void DoublesAndDecimalsRoundTripExactly()
        {
            // .NET Framework formats doubles differently from modern .NET ("R" is not always shortest there).
            double[] values = { 0.1, 1.0 / 3, 2.675, 1e-7, 123456789.123456789, -0.30000000000000004, double.MaxValue / 2 };
            var bytes = ExcelPackage.CreateBytes(wb =>
            {
                var s = wb.Worksheets.Add("Data");
                foreach (var v in values) s.Row().WriteCell(v).WriteCell(12345.6789m);
            });
            using (var reader = ExcelReader.Open(new MemoryStream(bytes)))
            {
                var rows = reader.ReadRows(0).ToList();
                for (var i = 0; i < values.Length; i++)
                {
                    Assert.Equal(values[i], (double)rows[i][0]!);
                    Assert.Equal(12345.6789, (double)rows[i][1]!);
                }
            }
        }

        [Fact]
        public void CachedRoundAndEarlyDates_MatchExcel()
        {
            var bytes = ExcelPackage.CreateBytes(wb =>
            {
                var s = wb.Worksheets.Add("Data");
                s.CacheFormulaResults();
                s.Row().WriteCell(1.005).WriteFormula("ROUND(A1,2)");
                s.Row().WriteCell(0.285).WriteFormula("ROUND(A2,2)");
                s.Row().WriteCell(new DateTime(1900, 1, 1));
            });
            using (var reader = ExcelReader.Open(new MemoryStream(bytes)))
            {
                var rows = reader.ReadRows(0).ToList();
                Assert.Equal(1.01, rows[0][1]);
                Assert.Equal(0.29, rows[1][1]);
                Assert.Equal(new DateTime(1900, 1, 1), rows[2][0]);
            }
        }

        [Fact]
        public void EditInPlace_CommitsOnSaveOnly()
        {
            var dir = Path.Combine(Path.GetTempPath(), "ss-netfx-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            var path = Path.Combine(dir, "book.xlsx");
            try
            {
                using (var p = ExcelPackage.Create(path))
                {
                    p.Workbook.Worksheets.Add("Data").Row().WriteCell("original").WriteCell(1);
                    p.Save();
                }
                var original = File.ReadAllBytes(path);

                using (var p = ExcelPackage.Edit(path))
                    p.Workbook.Worksheets.Edit("Data").Cells["A1"].Value = "not saved";
                Assert.Equal(original, File.ReadAllBytes(path));

                using (var p = ExcelPackage.Edit(path))
                {
                    p.Workbook.Worksheets.Edit("Data").Cells["A1"].Value = "saved";
                    p.Save();
                }
                Assert.Single(Directory.GetFiles(dir));
                using (var reader = ExcelReader.Open(path))
                    Assert.Equal("saved", reader.ReadRows(0).First()[0]);
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        [Fact]
        public void ReaderLimits_RejectAnOversizedCell()
        {
            var bytes = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Data").Row().WriteInlineString(new string('A', 50000)));
            var ex = Assert.Throws<InvalidDataException>(() =>
            {
                using (var reader = ExcelReader.Open(new MemoryStream(bytes), false, new ReaderOptions { MaxCellTextLength = 1000 }))
                    reader.ReadRows(0).ToList();
            });
            Assert.Contains("MaxCellTextLength", ex.Message);

            using (var reader = ExcelReader.Open(new MemoryStream(bytes)))
                Assert.Equal(50000, ((string)reader.ReadRows(0).First()[0]!).Length);
        }

        [Fact]
        public void TemplateFill_KeepsTemplateCellsAndAddsRows()
        {
            var template = ExcelPackage.CreateBytes(wb =>
            {
                var s = wb.Worksheets.Add("Lines");
                s.Row().WriteCells("Sku", "Qty");
            });
            var output = new MemoryStream();
            using (var p = ExcelPackage.OpenTemplate(new MemoryStream(template), output, leaveOpen: true))
            {
                var s = p.Workbook.Worksheets.Open("Lines");
                for (var i = 0; i < 10; i++) s.Row(2 + i).WriteCell("SKU-" + i).WriteCell(i);
                p.Save();
            }
            using (var reader = ExcelReader.Open(new MemoryStream(output.ToArray())))
            {
                var rows = reader.ReadRows("Lines").ToList();
                Assert.Equal(11, rows.Count);
                Assert.Equal("Sku", rows[0][0]);
                Assert.Equal("SKU-9", rows[10][0]);
            }
        }

        // ------------------------------------------------------------------ files built to exhaust memory
        //
        // The limits run on the netstandard2.0 build too, where the XML parser, the decoders and the array
        // searches are different implementations. The main test project covers every shape on modern .NET;
        // these repeat one of each kind here.

        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        /// <summary>A small genuine workbook with its sheet replaced by <paramref name="write"/>.</summary>
        private static byte[] WithSheet(System.Text.Encoding encoding, Action<TextWriter> write)
        {
            var seed = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Data").Row().WriteCell("Name").WriteCell(1));
            var output = new MemoryStream();
            using (var from = new System.IO.Compression.ZipArchive(new MemoryStream(seed), System.IO.Compression.ZipArchiveMode.Read))
            using (var to = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var entry in from.Entries)
                {
                    using (var target = to.CreateEntry(entry.FullName, System.IO.Compression.CompressionLevel.Fastest).Open())
                    {
                        if (entry.FullName != "xl/worksheets/sheet1.xml")
                        {
                            using (var source = entry.Open()) source.CopyTo(target);
                            continue;
                        }
                        using (var writer = new StreamWriter(target, encoding, 1 << 16)) write(writer);
                    }
                }
            }
            return output.ToArray();
        }

        private static void Repeat(TextWriter w, string unit, int chars)
        {
            var block = string.Concat(Enumerable.Repeat(unit, Math.Max(1, 32768 / unit.Length)));
            for (long done = 0; done < chars; done += block.Length) w.Write(block);
        }

        public static IEnumerable<object[]> HostileSheets()
        {
            var utf8 = new System.Text.UTF8Encoding(false);
            const int size = 20000000;
            var head = "<worksheet xmlns=\"" + Main + "\"><sheetData>";
            yield return new object[] { "long text", WithSheet(utf8, w => { w.Write(head + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>"); Repeat(w, "A", size); w.Write("</t></is></c></row></sheetData></worksheet>"); }) };
            yield return new object[] { "long attribute", WithSheet(utf8, w => { w.Write(head + "<row r=\"1\"><c r=\"A1\" x=\""); Repeat(w, "A", size); w.Write("\"><v>1</v></c></row></sheetData></worksheet>"); }) };
            yield return new object[] { "CDATA", WithSheet(utf8, w => { w.Write(head + "<row r=\"1\"><![CDATA["); Repeat(w, "<", size); w.Write("]]><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>"); }) };
            yield return new object[] { "UTF-16 attribute whose bytes look like markup", WithSheet(System.Text.Encoding.Unicode, w => { w.Write("<?xml version=\"1.0\" encoding=\"utf-16\"?>" + head + "<row r=\"1\"><c r=\"A1\" x=\""); Repeat(w, "㰼", size); w.Write("\"><v>1</v></c></row></sheetData></worksheet>"); }) };
            yield return new object[] { "nesting", WithSheet(utf8, w => { w.Write(head); Repeat(w, "<a>", size); }) };
            yield return new object[] { "distinct names", WithSheet(utf8, w => { w.Write(head + "<row r=\"1\">"); for (var i = 0; i < size / 10; i++) w.Write("<k" + i + "/>"); w.Write("</row></sheetData></worksheet>"); }) };
            yield return new object[] { "very many attributes", WithSheet(utf8, w => { w.Write(head + "<row r=\"1\"><c r=\"A1\""); for (var i = 0; i < size / 12; i++) w.Write(" a" + i + "=\"1\""); w.Write("><v>1</v></c></row></sheetData></worksheet>"); }) };
        }

        [Theory]
        [MemberData(nameof(HostileSheets))]
        public void HostileSheet_IsRefusedCheaply_ByTheReaderEditAndTemplateFill(string what, byte[] workbook)
        {
            var limits = new ReaderOptions { MaxCellTextLength = 20000, MaxLoadedPartBytes = 2000000 };
            var actions = new Action[]
            {
                () => { using (var reader = ExcelReader.Open(new MemoryStream(workbook), false, limits)) foreach (var row in reader.ReadRows(0)) GC.KeepAlive(row); },
                () =>
                {
                    using (var package = ExcelPackage.Edit(new MemoryStream(workbook), Stream.Null, leaveOpen: true, new PackageOptions { SourceLimits = limits }))
                    {
                        package.Workbook.Worksheets.Edit("Data").Cells["B2"].Value = 5;
                        package.Save();
                    }
                },
                () =>
                {
                    using (var package = ExcelPackage.OpenTemplate(new MemoryStream(workbook), Stream.Null, leaveOpen: true, new PackageOptions { SourceLimits = limits }))
                    {
                        package.Workbook.Worksheets.Open("Data").Row(4).WriteCell("x");
                        package.Save();
                    }
                },
            };
            foreach (var action in actions)
            {
                GC.Collect();
                var before = GC.GetTotalMemory(true);
                long peak = 0;
                Exception? thrown = null;
                // The heap is sampled while the work runs: .NET Framework has no allocation counter per
                // thread, and a path that held the 40 MB of hostile content would show here.
                var done = false;
                var sampler = new System.Threading.Thread(() => { while (!done) { peak = Math.Max(peak, GC.GetTotalMemory(false) - before); System.Threading.Thread.Sleep(1); } });
                sampler.Start();
                try { action(); }
                catch (Exception ex) { thrown = ex; }
                done = true;
                sampler.Join();

                Assert.True(thrown is InvalidDataException, what + ": expected InvalidDataException, got " + (thrown == null ? "no exception" : thrown.GetType().Name + ": " + thrown.Message));
                Assert.True(peak < 32L * 1024 * 1024, what + ": the heap grew by " + (peak / 1048576) + " MB while the file was being refused.");
            }
        }

        [Fact]
        public void DisposeWithoutSave_WritesNothing_AndCreateReplacesTheTargetOnlyOnSave()
        {
            var dir = Path.Combine(Path.GetTempPath(), "ss-netfx-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var path = Path.Combine(dir, "report.xlsx");
                using (var package = ExcelPackage.Create(path))
                    package.Workbook.Worksheets.Add("Data").Row().WriteCell("never saved");
                Assert.Empty(Directory.GetFiles(dir));

                using (var package = ExcelPackage.Create(path))
                {
                    package.Workbook.Worksheets.Add("Data").Row().WriteCell("first");
                    package.Save();
                }
                var first = File.ReadAllBytes(path);

                try
                {
                    using (var package = ExcelPackage.Create(path))
                    {
                        package.Workbook.Worksheets.Add("Data").Row().WriteCell("second");
                        throw new TimeoutException();
                    }
                }
                catch (TimeoutException) { }
                Assert.Equal(first, File.ReadAllBytes(path));

                using (var package = ExcelPackage.Create(path))
                {
                    package.Workbook.Worksheets.Add("Data").Row().WriteCell("third");
                    package.Save();
                }
                using (var reader = ExcelReader.Open(path)) Assert.Equal("third", reader.ReadRows(0).First()[0]);
                Assert.Single(Directory.GetFiles(dir));
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        // ------------------------------------------------------------------ where .NET Framework itself differs

        // The four documented types, plus the answer any workbook gives when asked for a sheet it does not
        // have: damage can empty the sheet list, and then "Data" is simply not there.
        private static bool IsDocumented(Exception ex)
            => ex is InvalidDataException || ex is System.Xml.XmlException || ex is FormatException || ex is OverflowException
               || (ex is ArgumentException argument && argument.ParamName == "name");

        private static byte[] SmallWorkbook() => ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            s.Row().WriteCell("Name").WriteCell(1.5).WriteCell(new DateTime(2026, 1, 2));
            s.Row().WriteCell("more").WriteCell(2).WriteFormula("B1*2", 3.0);
        });

        private static Exception? Open(byte[] bytes)
        {
            try
            {
                using (var reader = ExcelReader.Open(new MemoryStream(bytes)))
                    foreach (var name in reader.SheetNames)
                        foreach (var row in reader.ReadRows(name)) GC.KeepAlive(row);
                return null;
            }
            catch (Exception ex) { return ex; }
        }

        private static Exception? Edit(byte[] bytes)
        {
            try
            {
                using (var package = ExcelPackage.Edit(new MemoryStream(bytes), Stream.Null, leaveOpen: true))
                {
                    package.Workbook.Worksheets.Edit("Data").Cells["E5"].Value = 1;
                    package.Save();
                }
                return null;
            }
            catch (Exception ex) { return ex; }
        }

        [Theory]
        [InlineData('<')]
        [InlineData('>')]
        [InlineData('|')]
        [InlineData('"')]
        [InlineData('\u0001')]
        public void ZipEntryNameWindowsDoesNotAllow_IsReportedAsDamage(char illegal)
        {
            // .NET Framework reads the zip directory on first use, not when the archive is opened, and reports
            // such a name as ArgumentException ("Illegal characters in path") from deep inside a later call.
            var bytes = SmallWorkbook();
            var name = System.Text.Encoding.ASCII.GetBytes("docProps/app.xml");
            var replaced = 0;
            for (var i = 0; i + name.Length <= bytes.Length; i++)
            {
                var match = true;
                for (var k = 0; k < name.Length && match; k++) match = bytes[i + k] == name[k];
                if (!match) continue;
                bytes[i + 10] = (byte)illegal;   // docProps/a?p.xml, in the local header and in the directory.
                replaced++;
            }
            Assert.True(replaced >= 2, "The entry name should appear in the local header and in the directory.");

            foreach (var failure in new[] { Open(bytes), Edit(bytes) })
                Assert.True(failure is InvalidDataException, "Expected InvalidDataException, got " + (failure == null ? "no exception" : failure.GetType().Name + ": " + failure.Message));
        }

        [Fact]
        public void DamagedZipContainers_EndInTheDocumentedExceptionTypes()
        {
            var seed = SmallWorkbook();
            var failures = new List<string>();

            foreach (var length in new[] { 0, 10, 100, seed.Length / 3, seed.Length / 2, seed.Length - 30, seed.Length - 5 })
            {
                var cut = seed.Take(length).ToArray();
                foreach (var failure in new[] { Open(cut), Edit(cut) })
                    if (failure != null && !IsDocumented(failure)) failures.Add("truncated to " + length + ": " + failure.GetType().Name + ": " + failure.Message);
            }

            // Random bytes changed anywhere in the file: headers, the directory, compressed data.
            var random = new Random(4711);
            for (var iteration = 0; iteration < 3000; iteration++)
            {
                var bytes = (byte[])seed.Clone();
                var edits = 1 + random.Next(4);
                for (var e = 0; e < edits; e++) bytes[random.Next(bytes.Length)] = (byte)random.Next(256);
                foreach (var failure in new[] { Open(bytes), Edit(bytes) })
                    if (failure != null && !IsDocumented(failure) && failures.Count < 12) failures.Add("iteration " + iteration + ": " + failure.GetType().Name + ": " + failure.Message);
            }
            Assert.True(failures.Count == 0, string.Join("\n", failures));
        }

        [Fact]
        public void NumbersReadAsText_AreTheShortestFormThatReadsBack()
        {
            // .NET Framework's own round-trip format prints 17 digits where 16 would do. A reader has to hand
            // out the same text for a number on every runtime, and the same text Excel shows for it.
            var values = new[] { 1.0 / 3, 0.57000000000000006, 0.1, 2.675, 123456789.123456789, -0.30000000000000004, 42, 1e-7 };
            var expected = new[] { "0.3333333333333333", "0.5700000000000001", "0.1", "2.675", "123456789.12345679", "-0.30000000000000004", "42", "1E-07" };
            var bytes = ExcelPackage.CreateBytes(wb =>
            {
                var row = wb.Worksheets.Add("Data").Row();
                foreach (var v in values) row.WriteCell(v);
            });
            using (var reader = ExcelReader.Open(new MemoryStream(bytes)))
            {
                var row = reader.ReadRows(0).Single();
                for (var i = 0; i < values.Length; i++)
                {
                    Assert.Equal(expected[i], row.GetString(i));
                    Assert.Equal(values[i], double.Parse(row.GetString(i)!, System.Globalization.CultureInfo.InvariantCulture));
                }
            }
        }
    }
}
