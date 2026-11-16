using System;
using System.Globalization;

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

        /// <summary>Checks the text of a cell value that is a string or rich text; any other kind of value has no such limit.</summary>
        public static void EnsureFits(object? value, string parameterName)
        {
            switch (value)
            {
                case string text: EnsureFits(text.Length, parameterName); break;
                case RichText run: EnsureFits(run.Text.Length, parameterName); break;
                case RichText[] runs: EnsureFits(Length(runs), parameterName); break;
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