// The Open XML SDK exposes nearly every member as nullable; in test assertions a null is a failure anyway.
#pragma warning disable CS8602, CS8603, CS8604
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SheetStream.Internal;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SheetStream.Tests;

/// <summary>
/// AutoFit sizing checked against Excel's own AutoFit. The expected widths were obtained by auto-fitting the same
/// texts in Excel 2016 (Calibri 11, 96 dpi) and reading the stored <c>col width</c> back: the UI "column width"
/// plus 5/7 for the cell padding, e.g. ten digits give 10.29 in the UI and 11.0 in the file.
/// </summary>
public class AutoFitTests
{
    /// <summary>Writes one AutoFit sheet and returns its stored column widths by 1-based column index.</summary>
    private static Dictionary<int, double> Widths(Action<Workbook, Worksheet> fill, double minWidth = 1, double maxWidth = 255,
        Func<string, bool, double>? measure = null, PackageOptions? options = null)
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var sheet = wb.Worksheets.Add("Fit");
            if (measure != null) sheet.AutoFitColumns(measure, minWidth, maxWidth);
            else sheet.AutoFitColumns(minWidth, maxWidth);
            fill(wb, sheet);
        }, options);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var sheetId = doc.WorkbookPart.Workbook.Sheets.Elements<X.Sheet>().Single().Id;
        var ws = ((WorksheetPart)doc.WorkbookPart.GetPartById(sheetId)).Worksheet;
        // Adjacent columns with the same width are coalesced into one <col min max> element; expand them.
        var widths = new Dictionary<int, double>();
        var cols = ws.GetFirstChild<X.Columns>();
        if (cols != null)
            foreach (var c in cols.Elements<X.Column>())
                for (var i = (int)c.Min.Value; i <= (int)c.Max.Value; i++) widths[i] = c.Width.Value;
        return widths;
    }

    [Fact]
    public void TenDigits_MeasureTenUnits_AndGetExcelsWidth()
    {
        Assert.Equal(10.0, TextMeasure.Measure("0000000000", false));
        Assert.Equal(10.0, TextMeasure.Measure("0000000000", true));

        var cols = Widths((wb, s) => s.Row().WriteCell("0000000000").WriteCell(1234567890));
        Assert.InRange(cols[1], 10.9, 11.2); // Excel: 11.0 (UI 10.29).
        Assert.Equal(cols[1], cols[2]);      // A number shows the same digits.
    }

    [Fact]
    public void NarrowGlyphs_GiveNarrowerColumnsThanWideOnes()
    {
        var cols = Widths((wb, s) => s.Row().WriteCell("iiiiiiiiii").WriteCell("WWWWWWWWWW").WriteCell("0000000000"));
        Assert.InRange(cols[1], 6.3, 7.0);    // Excel: 6.71.
        Assert.InRange(cols[2], 19.5, 20.2);  // Excel: 19.86.
        Assert.True(cols[1] < cols[3] && cols[3] < cols[2]);
        Assert.True(cols[2] > 2.5 * cols[1]);
    }

    [Fact]
    public void Headers_MatchExcelAutoFit()
    {
        var cols = Widths((wb, s) =>
        {
            var bold = wb.Styles.Create(o => o.Font.Bold = true);
            s.Row().WriteCell("Description", bold).WriteCell("Qty", bold).WriteCell("Customer Name", bold).WriteCell("Qty");
        });
        Assert.InRange(cols[1], 10.9, 11.5);  // Excel: 11.14 for bold "Description" (regular measures the same).
        Assert.InRange(cols[2], 3.7, 4.4);    // Excel: 4.14.
        Assert.InRange(cols[3], 15.0, 15.7);  // Excel: 15.43.
        Assert.Equal(cols[2], cols[4]);       // No glyph of "Qty" widens in Calibri Bold.
    }

    [Fact]
    public void BoldText_IsWiderWhenItsGlyphsWiden()
    {
        // W, U, G, M, Y gain a pixel in Calibri Bold; most glyphs do not.
        Assert.True(TextMeasure.Measure("WAREHOUSE", true) > TextMeasure.Measure("WAREHOUSE", false));
        Assert.Equal(TextMeasure.Measure("Description", true), TextMeasure.Measure("Description", false));

        var cols = Widths((wb, s) =>
        {
            var bold = wb.Styles.Create(o => o.Font.Bold = true);
            s.Row().WriteCell("WAREHOUSE", bold).WriteCell("WAREHOUSE");
        });
        Assert.True(cols[1] > cols[2], $"bold {cols[1]} should exceed regular {cols[2]}");
        Assert.InRange(cols[1] - cols[2], 0.2, 0.5);
    }

    [Fact]
    public void LargerFont_ScalesWidth()
    {
        var cols = Widths((wb, s) =>
        {
            var big = wb.Styles.Create(o => o.Font.Size = 22);
            s.Row().WriteCell("Total", big).WriteCell("Total");
        });
        Assert.InRange(cols[1] / cols[2], 1.7, 2.0); // Glyphs double, the fixed padding does not.
    }

    [Fact]
    public void Dates_MeasureTheirDisplayedText()
    {
        var cols = Widths((wb, s) =>
        {
            var iso = wb.Styles.Create(o => o.NumberFormat = "yyyy-mm-dd");
            var dmy = wb.Styles.Create(o => o.NumberFormat = "d-mmm-yy");
            var longDate = wb.Styles.Create(o => o.NumberFormat = "dddd, mmmm d, yyyy");
            s.Row().WriteCell(new DateTime(2026, 10, 13))
                .WriteCell(new DateTime(2026, 10, 13, 15, 30, 0))
                .WriteCell(new DateTime(2026, 10, 13), iso)
                .WriteCell(new DateTime(2026, 10, 13), dmy)
                .WriteCell(new DateTime(2026, 10, 13), longDate)
                .WriteCell(TimeSpan.FromHours(36));
        });
        Assert.InRange(cols[1], 10.5, 11.0);  // Built-in short date shows "10/13/2026" in en-US; Excel: 10.71.
        Assert.InRange(cols[2], 15.7, 16.2);  // "10/13/2026 15:30" (Excel's "10/13/26 15:30" measured 13.86; two more digits here).
        Assert.InRange(cols[3], 10.2, 10.7);  // "2026-10-13"; Excel: 10.43.
        Assert.InRange(cols[4], 9.2, 9.7);    // "13-Oct-26"; Excel: 9.43 (the month name comes from the value).
        Assert.InRange(cols[5], 24.0, 25.0);  // "Tuesday, October 13, 2026"; Excel: 24.29.
        Assert.InRange(cols[6], 7.9, 8.4);    // "36:00:00"; Excel: 8.14.
    }

    [Fact]
    public void FormattedNumbers_MeasureTheirDisplayedText()
    {
        var cols = Widths((wb, s) =>
        {
            var money = wb.Styles.Create(o => o.NumberFormat = "#,##0.00");
            var percent = wb.Styles.Create(o => o.NumberFormat = "0%");
            var dollars = wb.Styles.Create(o => o.NumberFormat = "$#,##0.00");
            var euros = wb.Styles.Create(o => o.NumberFormat = "#,##0.00 \"EUR\"");
            s.Row().WriteCell(1234567.89, money).WriteCell(0.5, percent).WriteCell(1234.5, dollars).WriteCell(1234.5, euros)
                .WriteCell(123456789).WriteCell(1234567890123.0).WriteCell(true);
        });
        Assert.InRange(cols[1], 11.5, 12.1);  // "1,234,567.89"; Excel: 11.71.
        Assert.InRange(cols[2], 4.2, 4.8);    // "50%"; Excel: 4.57.
        Assert.InRange(cols[3], 8.9, 9.4);    // "$1,234.50"; Excel: 9.14.
        Assert.InRange(cols[4], 11.8, 12.3);  // "1,234.50 EUR"; Excel: 12.0.
        Assert.InRange(cols[5], 9.8, 10.3);   // General integer "123456789"; Excel: 10.0.
        Assert.InRange(cols[6], 11.6, 12.3);  // General beyond 11 digits: "1.23457E+12"; Excel: 12.0.
        Assert.InRange(cols[7], 5.1, 5.6);    // "TRUE"; Excel: 5.43.
    }

    [Theory]
    [InlineData("#,##0.00", 1234567.89, "1,234,567.89")]
    [InlineData("0.00%", 0.1234, "12.34%")]
    [InlineData("$#,##0.00", 1234.5, "$1,234.50")]
    [InlineData("#,##0.00 \"EUR\"", 1234.5, "1,234.50 EUR")]
    [InlineData("_(#,##0.00_);(#,##0.00)", 1234.5, "(1,234.50)")]
    [InlineData("[$€-407] #,##0.00;[Red]-[$€-407] #,##0.00", 99.5, "€ 99.50")]
    [InlineData("0.00E+00", 123456.0, "1.23E+05")]
    [InlineData("yyyy-mm-dd", 0.0, "0000-00-00")]
    [InlineData("m/d/yy h:mm", 0.0, "00/00/00 00:00")]
    [InlineData("dddd, mmmm d, yyyy", 46308.0, "Tuesday, October 00, 0000")] // 2026-10-13
    [InlineData("d-mmm-yy", 46308.0, "00-Oct-00")]
    [InlineData("[h]:mm:ss", 1.5, "00:00:00")]      // 36 hours
    [InlineData("[h]:mm:ss", 45.0, "0000:00:00")]   // 1080 hours
    [InlineData("[mm]:ss", 0.01, "00:00")]          // 14 minutes, still two digits
    [InlineData("h:mm AM/PM", 0.0, "00:00 AM")]
    [InlineData("General", 1234.5678, "1234.5678")]
    [InlineData("General", 1234567890123.0, "1.23457E+12")]
    [InlineData("General", 0.000012345, "1.2345E-05")]
    [InlineData("@", 42.0, "42")]
    public void NumberDisplayFormat_RendersWhatExcelShows(string code, double value, string expected)
        => Assert.Equal(expected, NumberDisplayFormat.Parse(code).Render(value));

    [Fact]
    public void AccentedLetters_MeasureLikeTheirBaseLetter()
    {
        Assert.Equal(TextMeasure.Measure("Descripción", false), TextMeasure.Measure("Descripcion", false));
        Assert.Equal(TextMeasure.Measure("Ñ", false), TextMeasure.Measure("N", false));
        Assert.Equal(2.0, TextMeasure.Measure("漢", false));          // CJK: two digits wide.
        Assert.Equal(1.0, TextMeasure.Measure("Ж", false));          // Unknown script: one digit.
        Assert.Equal(TextMeasure.Measure("e", false), TextMeasure.Measure("é", false)); // Combining accent adds nothing.
        // Every code unit measures without throwing.
        for (var c = 0; c <= 0xFFFF; c++) TextMeasure.PixelWidth((char)c, c % 2 == 0);
    }

    [Fact]
    public void MultiLineText_IsAsWideAsItsWidestLine()
    {
        var cols = Widths((wb, s) => s.Row().WriteCell("ab\ncdefgh\r\nx").WriteCell("cdefgh"));
        Assert.Equal(cols[2], cols[1]);
    }

    [Fact]
    public void CustomMeasurer_IsHonouredAndToldAboutBold()
    {
        var calls = new List<(string Text, bool Bold)>();
        double Measure(string text, bool bold) { calls.Add((text, bold)); return text.Length * 3; }

        var cols = Widths((wb, s) =>
        {
            var bold = wb.Styles.Create(o => o.Font.Bold = true);
            s.Row().WriteCell("Head", bold).WriteCell("ab\ncd");
            s.Row().WriteCell("abcd").WriteCell(12);
        }, measure: Measure);

        // 12 units of text + Excel's padding, through the same conversion as the built-in table (stored to 2 decimals).
        Assert.Equal(Worksheet.ExcelColumnWidth(12), cols[1], 2);
        Assert.InRange(cols[1], 13.0, 13.3);
        Assert.Contains(("Head", true), calls);
        Assert.Contains(("abcd", false), calls);
        Assert.Contains(("12", false), calls);
        Assert.Contains(("cd", false), calls); // Lines of a multi-line cell are measured one by one.
        Assert.DoesNotContain(calls, c => c.Text.Contains('\n'));
    }

    [Fact]
    public void PackageOptions_TextMeasurer_AppliesToEverySheet()
    {
        var options = new PackageOptions { TextMeasurer = (text, bold) => 20 };
        var cols = Widths((wb, s) => s.Row().WriteCell("x"), options: options);
        Assert.Equal(Worksheet.ExcelColumnWidth(20), cols[1], 2);

        // A per-sheet measurer wins over the package one.
        cols = Widths((wb, s) => s.Row().WriteCell("x"), measure: (text, bold) => 5, options: options);
        Assert.Equal(Worksheet.ExcelColumnWidth(5), cols[1], 2);
    }

    [Fact]
    public void Clamps_AndEmptyColumns_KeepTheirBehaviour()
    {
        var cols = Widths((wb, s) =>
        {
            s.Row().WriteCell("a").Skip(1).WriteCell(new string('x', 400));
            s.Column(2).SetWidth(12.5);
        }, minWidth: 6, maxWidth: 80);
        Assert.Equal(6, cols[1]);       // Below the minimum.
        Assert.Equal(12.5, cols[2]);    // Explicit width untouched by AutoFit.
        Assert.Equal(80, cols[3]);      // Clamped to the maximum.
        Assert.False(cols.ContainsKey(4)); // Never written: no column record.
    }

    [Fact]
    public void ExcelColumnWidth_FollowsTheEcmaFormula()
    {
        // Truncate((pixels + 5) / 7 * 256) / 256 with ten digits laid out at 72.5 px.
        Assert.Equal(Math.Truncate((72.5 + 5) / 7 * 256) / 256, Worksheet.ExcelColumnWidth(10), 3);
        Assert.Equal(5.0 / 7, Worksheet.ExcelColumnWidth(0), 2); // Padding only.
    }
}
