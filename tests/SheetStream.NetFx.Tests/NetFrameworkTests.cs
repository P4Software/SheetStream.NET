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
    }
}
