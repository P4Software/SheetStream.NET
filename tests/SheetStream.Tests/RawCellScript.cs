using System;
using System.Collections.Generic;
using System.Text;
using SheetStream.Styles;

namespace SheetStream.Tests;

/// <summary>The kinds of cell a <see cref="RawCellScript"/> can hold, one per public way of writing a cell.</summary>
internal enum ScriptKind
{
    Int, Long, Double, Float, Decimal, Date, TimeSpan, Bool,
    /// <summary><c>WriteCell(string)</c>: shared string table, or inline when the package turns the table off.</summary>
    Text,
    /// <summary><c>WriteInlineString</c>.</summary>
    Inline,
    /// <summary><c>WriteCell((string?)null, style)</c>: an empty cell that exists only when it has a style.</summary>
    NullText,
    /// <summary><c>WriteEmpty(style)</c>.</summary>
    Empty,
    /// <summary><c>Skip(n)</c>, with <see cref="ScriptCell.Integer"/> columns.</summary>
    Skip,
}

/// <summary>One call on a <see cref="RowWriter"/>. A plain struct so scripts are cheap to build in thousands.</summary>
internal struct ScriptCell
{
    public ScriptKind Kind;
    public long Integer;
    public double Number;
    public decimal Money;
    public DateTime Date;
    public TimeSpan Span;
    public string? Text;
    /// <summary>Index into the style array handed to <see cref="RawCellScript.Apply"/>, or -1 for no style.</summary>
    public int Style;
}

/// <summary>
/// A list of rows of cell writes, described independently of the library so the same script can drive the real
/// <see cref="RowWriter"/> and the reference writer in <c>RawCellWritingTests</c>. It uses only public API, so the
/// file also compiles against an older build of the library, which is how output was compared with the version
/// that wrote every cell through <c>XmlWriter</c> calls.
/// </summary>
internal sealed class RawCellScript
{
    /// <summary>Rows in order; each is written with <c>Row()</c>, so row numbers are 1, 2, 3...</summary>
    public readonly List<List<ScriptCell>> Rows = new List<List<ScriptCell>>();

    /// <summary>How many styles <see cref="Apply"/> expects to be given.</summary>
    public const int StyleCount = 40;

    /// <summary>Creates the styles scripts refer to by index, with ids that run well past one digit.</summary>
    public static CellStyle[] CreateStyles(Workbook workbook)
    {
        var styles = new CellStyle[StyleCount];
        for (var i = 0; i < styles.Length; i++)
        {
            var color = "#" + (0x102030 + i * 0x010305).ToString("X6");
            styles[i] = workbook.Styles.Create(o => o.Fill.BackgroundColor = color);
        }
        return styles;
    }

    /// <summary>Starts a row and returns the list to add cells to.</summary>
    public List<ScriptCell> NewRow()
    {
        var row = new List<ScriptCell>();
        Rows.Add(row);
        return row;
    }

    /// <summary>Writes every row of the script to <paramref name="sheet"/> through the public row writer.</summary>
    public void Apply(Worksheet sheet, CellStyle[] styles)
    {
        foreach (var row in Rows)
        {
            var w = sheet.Row();
            foreach (var cell in row)
            {
                var style = cell.Style < 0 ? null : styles[cell.Style];
                switch (cell.Kind)
                {
                    case ScriptKind.Int: w.WriteCell((int)cell.Integer, style); break;
                    case ScriptKind.Long: w.WriteCell(cell.Integer, style); break;
                    case ScriptKind.Double: w.WriteCell(cell.Number, style); break;
                    case ScriptKind.Float: w.WriteCell((float)cell.Number, style); break;
                    case ScriptKind.Decimal: w.WriteCell(cell.Money, style); break;
                    case ScriptKind.Date: w.WriteCell(cell.Date, style); break;
                    case ScriptKind.TimeSpan: w.WriteCell(cell.Span, style); break;
                    case ScriptKind.Bool: w.WriteCell(cell.Integer != 0, style); break;
                    case ScriptKind.Text: w.WriteCell(cell.Text, style); break;
                    case ScriptKind.Inline: w.WriteInlineString(cell.Text, style); break;
                    case ScriptKind.NullText: w.WriteCell((string?)null, style); break;
                    case ScriptKind.Empty: w.WriteEmpty(style); break;
                    case ScriptKind.Skip: w.Skip((int)cell.Integer); break;
                    default: throw new InvalidOperationException(cell.Kind.ToString());
                }
            }
        }
    }

    // ---- Builders --------------------------------------------------------------------------------------------------

    private static readonly double[] AwkwardDoubles =
    {
        0, -0.0, 1, -1, 0.1, -0.1, 1.0 / 3, 2.0 / 3, 100, 1e15, 1e16, 123456789012345678, 1e21, 1e-5, 1.5e-7, 12345.6789,
        0.30000000000000004, double.Epsilon, -double.Epsilon, double.MaxValue, double.MinValue, 1.7976931348623157e308,
        9007199254740993, 45000.5, 45000.999988425923, 36526, 0.000001, 1234567.891,
        double.NaN, double.PositiveInfinity, double.NegativeInfinity,
    };

    private static readonly long[] AwkwardLongs =
    {
        0, 1, -1, 9, 10, -10, 99, 100, 299, 300, 301, int.MaxValue, int.MinValue, (long)int.MaxValue + 1, (long)int.MinValue - 1,
        long.MaxValue, long.MinValue, long.MinValue + 1, 999999999999999999, 1000000000000000000,
    };

    private static readonly decimal[] AwkwardDecimals =
    {
        0m, 1m, -1m, 1.10m, 0.0000000000000000000000000001m, decimal.MaxValue, decimal.MinValue, 123456789.123456789m, -0.5m, 100m,
        79228162514264337593543950335m, -0.0m, 0.1m, 1234567890.0987654321m,
    };

    private static readonly DateTime[] AwkwardDates =
    {
        new DateTime(1900, 1, 1), new DateTime(1900, 2, 28), new DateTime(1900, 3, 1), new DateTime(1899, 12, 31), new DateTime(1899, 12, 30, 6, 0, 0),
        new DateTime(2026, 11, 15), new DateTime(2026, 11, 15, 13, 45, 30, 123), new DateTime(2100, 1, 1, 23, 59, 59), new DateTime(9999, 12, 31),
        new DateTime(1970, 1, 1), new DateTime(2000, 2, 29, 0, 0, 1),
    };

    /// <summary>
    /// Text that has tripped XML writers: markup, entities, quotes, every kind of white space, Excel's own
    /// <c>_xHHHH_</c> escape, control characters, lone and paired surrogates, the two non-characters, and a
    /// handful of non-ASCII scripts.
    /// </summary>
    public static readonly string[] NastyStrings =
    {
        "", " ", "  ", "a", "plain text", "Hello, World", "Tom & Jerry", "a<b", "a>b", "a<b>c</b>", "&amp;", "&lt;", "&#65;", "<![CDATA[x]]>", "]]>", "]", "]]",
        "\"quoted\"", "it's", "'single'", "a\tb", "\ta", "a\t", "a\nb", "a\r\nb", "a\rb", "\n", "\r", "\r\n", "line1\nline2\n", "\n\n\n",
        " leading", "trailing ", " both ", "\u00A0nbsp", "nbsp\u00A0", "\u2003em", "em\u2003", "\u3000ideographic", "\u0085", "\u2028", "\u2029", "\uFEFF", "\u200B",
        "_x0041_", "_x005F_", "_X0041_", "_x00_", "_x0041", "x0041_", "a_x000D_b", "__x0041__", "_x004G_", "_x0041_x0042_", "_", "__",
        "\0", "a\0b", "\u0001", "\u0008", "\u000B", "\u000C", "\u001F", "\u007F", "\u0080", "\u009F", "\uFFFE", "\uFFFF", "\uFFFD",
        "\uD800", "\uDBFF", "\uDC00", "\uDFFF", "a\uD800b", "a\uDC00b", "\uD800\uD800", "\uDC00\uD800", "\uD83D\uDE00", "x\uD83D\uDE00y", "\uD83D", "\uDE00\uD83D",
        "\uD83D\uDE00\uD83D\uDE00", "\uD83D\uDE00&\uD83D\uDE00", "\uD800\uDC00", "\uDBFF\uDFFF",
        "\u00E9", "\u00F1and\u00FA", "\u65E5\u672C\u8A9E", "\u041F\u0440\u0438\u0432\u0435\u0442", "\u0645\u0631\u062D\u0628\u0627", "\uD83D\uDE00", "\u0301combining", "a\u0301",
        "&&&&", "<<<<", ">>>>", "&<>&<>", "<>", "><", "&amp;&amp;", "5 > 3 && 2 < 4", "a & b < c > d \" e ' f",
    };

    /// <summary>A script that writes every value in the curated lists above in every kind of cell, with and without styles.</summary>
    public static RawCellScript Curated()
    {
        var script = new RawCellScript();
        var style = 0;

        foreach (var d in AwkwardDoubles)
        {
            var row = script.NewRow();
            row.Add(new ScriptCell { Kind = ScriptKind.Double, Number = d, Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Double, Number = d, Style = style++ % StyleCount });
            row.Add(new ScriptCell { Kind = ScriptKind.Float, Number = d, Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Float, Number = d, Style = style++ % StyleCount });
        }
        foreach (var l in AwkwardLongs)
        {
            var row = script.NewRow();
            row.Add(new ScriptCell { Kind = ScriptKind.Long, Integer = l, Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Long, Integer = l, Style = style++ % StyleCount });
            var narrowed = (int)Math.Max(int.MinValue, Math.Min(int.MaxValue, l));
            row.Add(new ScriptCell { Kind = ScriptKind.Int, Integer = narrowed, Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Int, Integer = narrowed, Style = style++ % StyleCount });
        }
        foreach (var m in AwkwardDecimals)
        {
            var row = script.NewRow();
            row.Add(new ScriptCell { Kind = ScriptKind.Decimal, Money = m, Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Decimal, Money = m, Style = style++ % StyleCount });
        }
        foreach (var d in AwkwardDates)
        {
            var row = script.NewRow();
            row.Add(new ScriptCell { Kind = ScriptKind.Date, Date = d, Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Date, Date = d, Style = style++ % StyleCount });
            row.Add(new ScriptCell { Kind = ScriptKind.TimeSpan, Span = TimeSpan.FromDays(d.ToOADate() % 1000), Style = -1 });
        }
        {
            var row = script.NewRow();
            for (var i = 0; i < 4; i++)
            {
                row.Add(new ScriptCell { Kind = ScriptKind.Bool, Integer = i & 1, Style = i < 2 ? -1 : i });
                row.Add(new ScriptCell { Kind = ScriptKind.Empty, Style = i < 2 ? -1 : i });
                row.Add(new ScriptCell { Kind = ScriptKind.NullText, Style = i < 2 ? -1 : i });
            }
        }
        foreach (var s in NastyStrings)
        {
            var row = script.NewRow();
            row.Add(new ScriptCell { Kind = ScriptKind.Text, Text = s, Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = s, Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Text, Text = s, Style = style++ % StyleCount });
            row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = s, Style = style++ % StyleCount });
        }
        // Gaps: the r attribute appears exactly where a cell is not the next one in line.
        {
            var row = script.NewRow();
            row.Add(new ScriptCell { Kind = ScriptKind.Skip, Integer = 2 });
            row.Add(new ScriptCell { Kind = ScriptKind.Int, Integer = 1, Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Empty, Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Text, Text = "after an unstyled empty cell", Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Skip, Integer = 0 });
            row.Add(new ScriptCell { Kind = ScriptKind.Bool, Integer = 1, Style = 3 });
            row.Add(new ScriptCell { Kind = ScriptKind.Skip, Integer = 700 });
            row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = "far right", Style = 5 });
            row.Add(new ScriptCell { Kind = ScriptKind.Skip, Integer = 3 });
            row.Add(new ScriptCell { Kind = ScriptKind.Double, Number = 2.5, Style = -1 });
        }
        script.NewRow(); // A row with no cells at all is written as a self-closing element.
        script.NewRow().Add(new ScriptCell { Kind = ScriptKind.Skip, Integer = 5 }); // A row of nothing but a gap.
        script.NewRow().Add(new ScriptCell { Kind = ScriptKind.Empty, Style = -1 });  // A row whose only cell writes nothing.
        return script;
    }

    /// <summary>One row per 64 characters of the BMP, each in the middle of a text, as a shared and as an inline string.</summary>
    public static RawCellScript CharacterSweep()
    {
        var script = new RawCellScript();
        for (var start = 0; start < 0x10000; start += 64)
        {
            var row = script.NewRow();
            for (var c = start; c < start + 64; c++)
            {
                var text = "a" + (char)c + "b";
                row.Add(new ScriptCell { Kind = ScriptKind.Text, Text = text, Style = -1 });
                row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = text, Style = -1 });
                // At the very start and end, where white space triggers xml:space and a lone surrogate has no partner.
                row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = ((char)c).ToString(), Style = -1 });
                row.Add(new ScriptCell { Kind = ScriptKind.Text, Text = "x" + (char)c, Style = -1 });
            }
        }
        // Every seventh high surrogate beside a few low ones, and the other way round: pairs, reversed pairs and orphans.
        for (var hi = 0xD800; hi <= 0xDBFF; hi += 7)
        {
            var row = script.NewRow();
            foreach (var lo in new[] { 0xDC00, 0xDC01, 0xDD00, 0xDFFF })
            {
                row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = "" + (char)hi + (char)lo, Style = -1 });
                row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = "" + (char)lo + (char)hi, Style = -1 });
                row.Add(new ScriptCell { Kind = ScriptKind.Text, Text = "<" + (char)hi + "&" + (char)lo + ">", Style = -1 });
            }
        }
        return script;
    }

    /// <summary>
    /// Strings around the size of the row writer's staging buffer, with the characters that need an entity placed at
    /// scattered offsets, so a flush lands in the middle of text, of an entity and between two cells.
    /// </summary>
    public static RawCellScript BufferBoundaries(int bufferLength)
    {
        var script = new RawCellScript();
        var random = new Random(4711);
        for (var length = bufferLength - 260; length <= bufferLength + 260; length += 3)
        {
            var row = script.NewRow();
            foreach (var special in new[] { '&', '<', '>', 'x', ' ' })
            {
                var chars = new char[length];
                for (var i = 0; i < chars.Length; i++) chars[i] = 'k';
                for (var i = 0; i < 6; i++) chars[random.Next(length)] = special;
                chars[length - 1] = special;
                if (special == ' ') chars[0] = ' ';
                var kind = special == '&' || special == '>' ? ScriptKind.Text : ScriptKind.Inline;
                row.Add(new ScriptCell { Kind = ScriptKind.Int, Integer = length, Style = -1 });
                row.Add(new ScriptCell { Kind = kind, Text = new string(chars), Style = length % 7 == 0 ? 2 : -1 });
            }
        }
        // Strings that are nothing but entities, well past the buffer, and ones several buffers long.
        foreach (var c in new[] { '&', '<', '>' })
        {
            var row = script.NewRow();
            row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = new string(c, bufferLength + 17), Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = new string(c, bufferLength * 3 + 1), Style = -1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Text, Text = new string(c, bufferLength * 2), Style = -1 });
        }
        {
            var row = script.NewRow();
            row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = new string('z', bufferLength * 5 + 3), Style = 1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = new string('z', bufferLength), Style = 1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = new string('z', bufferLength - 1), Style = 1 });
            row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = new string('z', bufferLength + 1), Style = 1 });
        }
        // Many short cells in one row, so the buffer fills between cells and not inside one.
        {
            var row = script.NewRow();
            for (var i = 0; i < 3000; i++)
                row.Add(new ScriptCell { Kind = i % 3 == 0 ? ScriptKind.Double : i % 3 == 1 ? ScriptKind.Inline : ScriptKind.Int, Number = i * 1.5, Integer = i, Text = "v" + i, Style = i % 5 == 0 ? i % StyleCount : -1 });
        }
        return script;
    }

    private static readonly string[] FuzzTokens =
    {
        "a", "b", "Z", "0", "9", " ", " ", "  ", "\t", "\n", "\r", "\r\n", "&", "&", "<", "<", ">", ">", "\"", "'", "]", "]]>", "_", "x", "F", "_x0041_", "_x005F_",
        "_x000D_", "\0", "\u0001", "\u001F", "\u007F", "\u0085", "\u00A0", "\u2028", "\uFEFF", "\uFFFE", "\uFFFF", "\uD800", "\uDC00", "\uDBFF", "\uDFFF",
        "\uD83D\uDE00", "\u00E9", "\u65E5\u672C", "&amp;", "<b>", "word", "Tom", "\u2014",
    };

    /// <summary>
    /// A random script from a fixed seed: every kind of cell, random styles and gaps, and strings built from the
    /// nasty tokens above, mostly short and now and then several buffers long.
    /// </summary>
    public static RawCellScript Fuzz(int seed, int rows)
    {
        var script = new RawCellScript();
        var r = new Random(seed);
        for (var rowIndex = 0; rowIndex < rows; rowIndex++)
        {
            var row = script.NewRow();
            var cells = r.Next(0, 14);
            for (var c = 0; c < cells; c++)
            {
                var style = r.Next(3) == 0 ? r.Next(StyleCount) : -1;
                var pick = r.Next(100);
                if (pick < 14) row.Add(new ScriptCell { Kind = ScriptKind.Int, Integer = RandomLong(r), Style = style });
                else if (pick < 22) row.Add(new ScriptCell { Kind = ScriptKind.Long, Integer = RandomLong(r), Style = style });
                else if (pick < 32) row.Add(new ScriptCell { Kind = ScriptKind.Double, Number = RandomDouble(r), Style = style });
                else if (pick < 36) row.Add(new ScriptCell { Kind = ScriptKind.Float, Number = RandomDouble(r), Style = style });
                else if (pick < 40) row.Add(new ScriptCell { Kind = ScriptKind.Decimal, Money = RandomDecimal(r), Style = style });
                else if (pick < 46) row.Add(new ScriptCell { Kind = ScriptKind.Date, Date = RandomDate(r), Style = style });
                else if (pick < 48) row.Add(new ScriptCell { Kind = ScriptKind.TimeSpan, Span = TimeSpan.FromSeconds(r.Next(-100000, 100000) / 7.0), Style = style });
                else if (pick < 52) row.Add(new ScriptCell { Kind = ScriptKind.Bool, Integer = r.Next(2), Style = style });
                else if (pick < 72) row.Add(new ScriptCell { Kind = ScriptKind.Text, Text = RandomText(r), Style = style });
                else if (pick < 90) row.Add(new ScriptCell { Kind = ScriptKind.Inline, Text = RandomText(r), Style = style });
                else if (pick < 93) row.Add(new ScriptCell { Kind = ScriptKind.NullText, Style = style });
                else if (pick < 96) row.Add(new ScriptCell { Kind = ScriptKind.Empty, Style = style });
                else row.Add(new ScriptCell { Kind = ScriptKind.Skip, Integer = r.Next(0, 4) });
            }
        }
        return script;
    }

    private static long RandomLong(Random r)
    {
        var value = (((long)r.Next() << 32) | (uint)r.Next()) >> r.Next(0, 64);
        return r.Next(2) == 0 ? -value : value;
    }

    private static double RandomDouble(Random r)
    {
        if (r.Next(3) == 0) return BitConverter.Int64BitsToDouble(((long)r.Next() << 32) | (uint)r.Next());
        return Math.Round((r.NextDouble() - 0.3) * Math.Pow(10, r.Next(-4, 9)), r.Next(0, 6));
    }

    private static decimal RandomDecimal(Random r)
        => new decimal(r.Next(), r.Next(), r.Next(0, 1 << 20), r.Next(2) == 0, (byte)r.Next(0, 29));

    private static DateTime RandomDate(Random r)
    {
        var date = new DateTime(1900, 1, 1).AddSeconds(r.NextDouble() * 6e9);
        return r.Next(2) == 0 ? date.Date : date;
    }

    private static string RandomText(Random r)
    {
        // Mostly short, sometimes a few hundred characters, rarely longer than the row writer's buffer.
        var pick = r.Next(100);
        var tokens = pick < 70 ? r.Next(0, 8) : pick < 96 ? r.Next(8, 80) : pick < 99 ? r.Next(200, 3000) : r.Next(8000, 24000);
        var sb = new StringBuilder();
        for (var i = 0; i < tokens; i++) sb.Append(FuzzTokens[r.Next(FuzzTokens.Length)]);
        return sb.ToString();
    }
}
