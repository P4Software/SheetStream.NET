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
    /// The netstandard2.0 build on .NET Framework, for the fixes whose code differs from the modern build (a different
    /// shared string lookup, a different inline text encoder) or that must hold on both.
    /// </summary>
    public class ReviewFixesNetFxTests
    {
        [Fact]
        public void Skip_PastTheLastColumn_IsRefused_AndTheRowStaysValid()
        {
            var bytes = ExcelPackage.CreateBytes(wb =>
            {
                var s = wb.Worksheets.Add("S");
                var r = s.Row().WriteCell("a");
                Assert.Throws<ArgumentOutOfRangeException>(() => r.Skip(int.MaxValue));
                Assert.Equal(2, r.NextColumn);
                r.WriteCell("b");
                r.Skip(16382); // To the column after XFD: the row is complete.
                Assert.Throws<InvalidOperationException>(() => r.WriteCell("past the edge"));
                s.Row().WriteCell("next");
            });
            using (var reader = ExcelReader.Open(new MemoryStream(bytes)))
            {
                var rows = reader.ReadRows("S").ToList();
                Assert.Equal(2, rows.Count);
                Assert.Equal("a", rows[0][0]);
                Assert.Equal("b", rows[0][1]);
                Assert.Equal("next", rows[1][0]);
            }
        }

        [Fact]
        public void TextOfMoreThan32767Characters_IsRefused_AndTheWriterStaysUsable()
        {
            var longest = new string('w', 32767);
            var bytes = ExcelPackage.CreateBytes(wb =>
            {
                var s = wb.Worksheets.Add("S");
                var r = s.Row().WriteCell(longest);
                var e = Assert.Throws<ArgumentException>(() => r.WriteCell(new string('w', 32768)));
                Assert.Contains("32,767", e.Message);
                Assert.Throws<ArgumentException>(() => r.WriteInlineString(new string('w', 32768)));
                Assert.Throws<ArgumentException>(() => r.WriteRichText(new RichText(new string('w', 32768))));
                Assert.Equal(2, r.NextColumn);
                r.WriteCell("after");
            });
            using (var reader = ExcelReader.Open(new MemoryStream(bytes)))
            {
                var row = reader.ReadRows("S").Single();
                Assert.Equal(longest, row[0]);
                Assert.Equal("after", row[1]);
            }
        }

        [Fact]
        public void CarriageReturns_AreWrittenAsCharacterReferences_AndReadBackAsCarriageReturns()
        {
            var texts = new[] { "a\r\nb", "a\rb", "\r", " lead\r\ntrail\r\n ", "x & <y>\r\n\"q\"\r_x0041_\t", "é€😀\r\nz" };
            foreach (var storage in new[] { StringStorage.Auto, StringStorage.Shared, StringStorage.Inline })
            {
                var bytes = ExcelPackage.CreateBytes(wb =>
                {
                    var s = wb.Worksheets.Add("S");
                    foreach (var text in texts) s.Row().WriteCell(text).WriteInlineString(text);
                    s.Row().WriteRichText(new RichText("one\r\ntwo") { Bold = true }, new RichText("\rthree"));
                }, new PackageOptions { StringStorage = storage });

                using (var reader = ExcelReader.Open(new MemoryStream(bytes)))
                {
                    var rows = reader.ReadRows("S").ToList();
                    for (var i = 0; i < texts.Length; i++)
                    {
                        Assert.Equal(texts[i], rows[i][0]);
                        Assert.Equal(texts[i], rows[i][1]);
                    }
                    Assert.Equal("one\r\ntwo\rthree", rows[texts.Length][0]);
                }
                using (var zip = new System.IO.Compression.ZipArchive(new MemoryStream(bytes), System.IO.Compression.ZipArchiveMode.Read))
                foreach (var name in new[] { "xl/worksheets/sheet1.xml", "xl/sharedStrings.xml" })
                using (var part = new StreamReader(zip.GetEntry(name).Open()))
                    Assert.DoesNotContain("\r", part.ReadToEnd());
            }
        }

        [Fact]
        public void AUniqueColumnOnASecondSheet_DoesNotGrowTheSharedTable()
        {
            long Unique(int sheets)
            {
                long unique = 0;
                ExcelPackage.CreateBytes(wb =>
                {
                    for (var n = 0; n < sheets; n++)
                    {
                        var s = wb.Worksheets.Add("S" + n);
                        for (var i = 0; i < 20000; i++) s.Row().WriteCell("Order SO-" + i.ToString("D8"));
                    }
                    unique = wb.SharedStrings.UniqueCount;
                });
                return unique;
            }

            Assert.Equal(Unique(1), Unique(2));
        }
    }
}
