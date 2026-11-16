using System;
using System.Globalization;
using SheetStream.Reading;
using SheetStream.Styles;

namespace SheetStream.Internal
{
    /// <summary>
    /// The one limit on the text of a cell that Excel enforces when it opens a file: 32,767 characters. A workbook with a
    /// longer cell is refused as a whole ("Unable to get the Open property"), not just that cell, so the library
    /// refuses the text when it is given, in plain words, before anything of the cell is written.
    /// </summary>
    internal static class CellText
    {
        /// <summary>Most characters (UTF-16 code units, which is what Excel counts) a cell can hold.</summary>
        public const int MaxLength = 32767;

        /// <summary>
        /// Throws when <paramref name="length"/> is more than <see cref="MaxLength"/>. The length is that of the text the
        /// cell holds, not of its spelling in the file: a control character is written as <c>_x0001_</c> and a literal
        /// <c>_x0041_</c> as <c>_x005F_x0041_</c>, and Excel counts each as the one or seven characters the reader sees.
        /// </summary>
        /// <param name="length">The length of the text, or of all the runs of a rich-text cell together.</param>
        /// <param name="parameterName">The name of the argument that carried the text.</param>
        public static void EnsureFits(long length, string parameterName)
        {
            if (length > MaxLength) ThrowTooLong(length, parameterName);
        }

        /// <summary>
        /// Checks the text a cell value will become. A string or rich text is measured as it is. A number, a boolean, a date, a
        /// <see cref="TimeSpan"/>, a character, a <see cref="Guid"/>, an enumeration, an error value or a style turns into
        /// something short (or is not text), so those cost one type test each and nothing is allocated. Any other object
        /// (a <see cref="System.Text.StringBuilder"/>, a type with its own <c>ToString</c>) is written as text by
        /// <see cref="RowWriter.WriteCell(object?, CellStyle?)"/>, so it is converted here by the same rule (invariant
        /// culture for <see cref="IFormattable"/>, <c>ToString()</c> otherwise) and its length checked: refused when it is
        /// assigned, not when the sheet is sealed, where the failure would make the whole package unsavable.
        /// </summary>
        public static void EnsureFits(object? value, string parameterName)
        {
            switch (value)
            {
                case null:
                case DBNull _:
                case double _: case float _: case decimal _: case int _: case long _: case short _: case byte _: case sbyte _:
                case ushort _: case uint _: case ulong _: case bool _: case char _:
                case DateTime _: case DateTimeOffset _: case TimeSpan _: case Guid _: case Enum _:
                case ExcelErrorValue _: case CellStyle _:
                    return;
                case string text: EnsureFits(text.Length, parameterName); return;
                case RichText run: EnsureFits(run.Text.Length, parameterName); return;
                case RichText[] runs: EnsureFits(Length(runs), parameterName); return;
                case IFormattable formattable: EnsureFits(formattable.ToString(null, CultureInfo.InvariantCulture)?.Length ?? 0, parameterName); return;
                default: EnsureFits(value.ToString()?.Length ?? 0, parameterName); return;
            }
        }

        /// <summary>The characters of all the runs of a rich-text cell together (a null run counts nothing: the caller rejects it).</summary>
        public static long Length(RichText[] runs)
        {
            long total = 0;
            foreach (var run in runs)
                if (run != null) total += run.Text.Length;
            return total;
        }

        private static void ThrowTooLong(long length, string parameterName)
            => throw new ArgumentException(
                "Excel cannot open a cell with more than " + MaxLength.ToString("N0", CultureInfo.InvariantCulture) + " characters; this text has "
                + length.ToString("N0", CultureInfo.InvariantCulture) + ". Shorten it or split it over several cells.", parameterName);
    }
}