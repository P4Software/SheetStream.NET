using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SheetStream;
using SheetStream.Reading;
using SheetStream.Styles;

/// <summary>
/// Builds one workbook that uses every SheetStream feature, so a single file can be opened in Excel to
/// confirm nothing triggers a repair. Run with <c>dotnet run --project samples/SheetStream.Sample -- showcase</c>.
/// </summary>
internal static class Showcase
{
    /// <summary>An exported business object, shaped like what CifraHQ and NominaHQ exporters produce.</summary>
    private sealed class Invoice
    {
        [SheetColumn("Invoice #", Order = 0)]
        public string Number { get; set; } = "";

        public string Customer { get; set; } = "";

        public DateTime Issued { get; set; }

        [SheetColumn(NumberFormat = "$#,##0.00")]
        public decimal Amount { get; set; }

        public string Status { get; set; } = "";

        public bool Paid { get; set; }
    }

    public static void Run(string path)
    {
        var invoices = Enumerable.Range(1, 60).Select(i => new Invoice
        {
            Number = "INV-" + (1000 + i),
            Customer = new[] { "Acme Corp", "Globex", "Initech", "Umbrella", "Stark Industries" }[i % 5],
            Issued = new DateTime(2026, 1, 1).AddDays(i * 5),
            Amount = Math.Round(250m + i * 137.35m % 4000, 2),
            Status = new[] { "Open", "Closed", "On hold" }[i % 3],
            Paid = i % 3 == 1,
        }).ToList();

        using (var package = ExcelPackage.Create(path))
        {
            var wb = package.Workbook;
            wb.Properties.Title = "SheetStream feature showcase";
            wb.Properties.Author = "P4 Software";
            wb.Properties.Company = "P4 Software";

            var header = wb.Styles.Create(o =>
            {
                o.Font.Bold = true;
                o.Font.Color = Colors.White;
                o.Fill.BackgroundColor = Colors.ExcelBlue;
                o.Border.Style = BorderStyle.Thin;
                o.HorizontalAlignment = HorizontalAlignment.Center;
            });
            var title = wb.Styles.Create(o => { o.Font.Bold = true; o.Font.Size = 16; });
            var input = wb.Styles.Create(o => { o.Protection.Locked = false; o.Fill.BackgroundColor = "#FFF2CC"; o.Border.Style = BorderStyle.Dotted; });
            var bad = wb.Styles.CreateDifferential(o => { o.Font.Color = "#9C0006"; o.Fill.BackgroundColor = "#FFC7CE"; });
            var good = wb.Styles.CreateDifferential(o => { o.Font.Color = "#006100"; o.Fill.BackgroundColor = "#C6EFCE"; });

            // ---- Sheet 1: invoices as an Excel table, auto-fitted, with rules, validation and a chart.
            var data = wb.Worksheets.Add("Invoices");
            data.TabColor = "#4F81BD";
            data.AutoFitColumns();
            data.FreezePanes(rows: 2);
            data.PageSetup.Orientation = PageOrientation.Landscape;
            data.PageSetup.FitToWidth = 1;
            data.PageSetup.Footer = "Page &P of &N";
            data.PageSetup.PrintTitleRows = (2, 2);
            data.Row(1, title, 24).WriteRichText(title, new RichText("Invoices "), new RichText("2026") { Color = "#C00000" });
            var loaded = data.LoadFromCollection(invoices, new LoadOptions { HeaderStyle = header, TableName = "Invoices", TableStyle = "TableStyleMedium2" });
            var amount = "D" + (loaded.FirstRow + 1) + ":D" + loaded.LastRow;
            data.AddConditionalFormatting(amount).DataBar("#638EC6");
            data.AddConditionalFormatting("E3:E" + loaded.LastRow)
                .CellValue(ConditionalOperator.Equal, "\"On hold\"", bad)
                .CellValue(ConditionalOperator.Equal, "\"Closed\"", good);
            data.AddConditionalFormatting("F3:F" + loaded.LastRow).IconSet(IconSetType.ThreeSymbols);
            data.AddDataValidation("E3:E" + loaded.LastRow).List("Open", "Closed", "On hold").WithError("Unknown status", "Pick a status from the list.");
            data.Row().WriteCell("Total", header).Skip(2).WriteFormula($"SUBTOTAL(109,{amount})", (double)invoices.Sum(i => i.Amount), wb.Styles.Create(o => { o.Font.Bold = true; o.NumberFormat = "$#,##0.00"; }));
            data.AddComment(2, 4, "Amounts include tax.", "Finance");
            var chart = data.AddChart(ChartType.Column, "I3:Q20");
            chart.Title = "Amount by invoice";
            chart.YAxisTitle = "USD";
            chart.ValueAxisNumberFormat = "$#,##0";
            chart.Legend = LegendPosition.None;
            chart.AddSeries("D3:D22", "A3:A22", "=D2", "#4F81BD");

            // ---- Sheet 2: summary built with random access, a pie chart, links, a picture and protection.
            var summary = wb.Worksheets.Add("Summary");
            summary.ShowGridLines = false;
            summary.Column(1).SetWidth(22);
            summary.Column(2).SetWidth(16);
            summary.Cells["A1"].Value = "Summary";
            summary.Cells["A1"].Style = title;
            summary.Cells["A3"].Value = "Customer";
            summary.Cells["A3"].Style = header;
            summary.Cells["B3"].Value = "Total";
            summary.Cells["B3"].Style = header;
            var money = wb.Styles.Create(o => o.NumberFormat = "$#,##0.00");
            var row = 4;
            foreach (var g in invoices.GroupBy(i => i.Customer).OrderBy(g => g.Key))
            {
                summary.Cells[row, 1].Value = g.Key;
                summary.Cells[row, 2].Formula = $"SUMIFS(Invoices[Amount],Invoices[Customer],A{row})";
                summary.Cells[row, 2].Style = money;
                row++;
            }
            summary.Cells[row + 1, 1].Value = "Tax rate";
            summary.Cells[row + 1, 2].Value = 0.07;
            summary.Cells[row + 1, 2].Style = wb.Styles.Create(o => { o.NumberFormat = "0.0%"; o.Protection.Locked = false; o.Fill.BackgroundColor = "#FFF2CC"; });
            wb.DefinedNames.Add("TaxRate", $"'Summary'!B{row + 1}");
            summary.Cells[row + 2, 1].Value = "Back to invoices";
            summary.Cells[row + 2, 1].Style = wb.Styles.Hyperlink;
            summary.AddInternalLink(row + 2, 1, "'Invoices'!A1", "Jump to the invoice table");
            summary.Cells[row + 3, 1].Value = "P4 Software";
            summary.Cells[row + 3, 1].Style = wb.Styles.Hyperlink;
            summary.AddHyperlink(row + 3, 1, "https://github.com/P4Software/SheetStream");
            summary.MergeCells("A1:B1");
            var pie = summary.AddChart(ChartType.Pie, "D3:J18");
            pie.Title = "Share by customer";
            pie.DataLabels = ChartDataLabels.Percent;
            pie.AddSeries($"B4:B{row - 1}", $"A4:A{row - 1}", "Total");
            var banner = Path.Combine(AppContext.BaseDirectory, "banner.png");
            if (File.Exists(banner)) summary.AddPicture(banner, 20, 4);
            summary.Protect(configure: o => o.AllowSelectLockedCells = true);

            // ---- Sheet 3: outline grouping and a hidden lookup sheet.
            var grouped = wb.Worksheets.Add("Grouped");
            grouped.Row().WriteCell("Region", header).WriteCell("Sales", header);
            foreach (var region in new[] { "North", "South" })
            {
                grouped.Row().WriteCell(region, wb.Styles.Create(o => o.Font.Bold = true));
                for (var i = 1; i <= 3; i++)
                    grouped.Row(grouped.CurrentRow + 1, new RowOptions { OutlineLevel = 1 }).WriteCell($"  {region} store {i}").WriteCell(i * 1000);
            }
            var lists = wb.Worksheets.Add("Lists");
            lists.Visibility = SheetVisibility.Hidden;
            foreach (var s in new[] { "Open", "Closed", "On hold" }) lists.Row().WriteCell(s);
        }

        // Read it back with the streaming reader to prove the round trip.
        using var reader = ExcelReader.Open(path);
        Console.WriteLine($"Showcase written: {path}");
        Console.WriteLine($"Sheets: {string.Join(", ", reader.SheetNames)}");
        var back = reader.ReadRows("Invoices").Skip(2).Take(1).Single();
        Console.WriteLine($"First invoice read back: {back[0]}, {back[1]}, {back[2]:yyyy-MM-dd}, {back[3]}");
    }
}
