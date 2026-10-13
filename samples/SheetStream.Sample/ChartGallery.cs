using System;
using SheetStream;

/// <summary>
/// Writes one chart of every <see cref="ChartType"/> on a single sheet, so all chart kinds can be opened in
/// Excel together. Run with <c>dotnet run --project samples/SheetStream.Sample -- charts</c>.
/// </summary>
internal static class ChartGallery
{
    public static void Run(string path)
    {
        using var package = ExcelPackage.Create(path);
        var sheet = package.Workbook.Worksheets.Add("Charts");
        sheet.Row().WriteCells("Month", "Sales", "Cost", "X");
        for (var m = 1; m <= 6; m++)
            sheet.Row().WriteCell("M" + m).WriteCell(m * 100.0 + (m % 2) * 40).WriteCell(m * 60.0).WriteCell(m * 1.5);

        var types = (ChartType[])Enum.GetValues(typeof(ChartType));
        for (var i = 0; i < types.Length; i++)
        {
            // Three charts per band of rows, each 8 columns wide and 15 rows tall.
            var top = 10 + (i / 3) * 16;
            var left = 1 + (i % 3) * 8;
            var placement = Cell(top, left) + ":" + Cell(top + 14, left + 7);
            var chart = sheet.AddChart(types[i], placement);
            chart.Title = types[i].ToString();
            chart.Legend = LegendPosition.Bottom;
            if (types[i] == ChartType.Scatter || types[i] == ChartType.ScatterLines)
            {
                chart.XAxisTitle = "X";
                chart.AddSeries("B2:B7", "D2:D7", "=B1");
                chart.AddSeries("C2:C7", "D2:D7", "=C1", "#C0504D");
            }
            else
            {
                chart.DataLabels = types[i] == ChartType.Pie || types[i] == ChartType.Doughnut ? ChartDataLabels.Percent : ChartDataLabels.None;
                chart.AddSeries("B2:B7", "A2:A7", "=B1");
                chart.AddSeries("C2:C7", "A2:A7", "=C1", "#C0504D");
            }
        }
        Console.WriteLine($"Chart gallery written: {path} ({types.Length} charts)");
    }

    private static string Cell(int row, int column)
    {
        var letters = "";
        for (var n = column; n > 0; n = (n - 1) / 26) letters = (char)('A' + (n - 1) % 26) + letters;
        return letters + row;
    }
}
