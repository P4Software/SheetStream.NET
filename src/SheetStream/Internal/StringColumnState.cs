using System;

namespace SheetStream.Internal
{
    /// <summary>
    /// The numbers behind <see cref="StringStorage.Auto"/>. They come from measurements (see the changelog and
    /// <c>benchmarks/README.md</c>), not from taste, so a change here needs new measurements.
    /// </summary>
    internal static class StringColumnPolicy
    {
        /// <summary>How many cells of a column are looked up in the shared string table before the column is judged, unless it is judged earlier (<see cref="EarlyExitCells"/>).</summary>
        public const int ProbeCells = 512;

        /// <summary>
        /// A column in which not one of the first cells repeated is unique (ids, e-mail addresses, free text) and is judged
        /// at once, which spares such a column most of the sampling: each sampled text is hashed, stored in the table and
        /// kept until the save. A sheet of fewer rows than this is stored exactly as <see cref="StringStorage.Shared"/> stores it.
        /// </summary>
        public const int EarlyExitCells = 128;

        /// <summary>
        /// Share a column when at least this many percent of the cells in the sample were already in the table.
        /// A column of unique text (ids, e-mail addresses, free text) has none; one with a few accidental
        /// repeats stays well under it.
        /// </summary>
        public const int EnterPercent = 5;

        /// <summary>
        /// Repeated text this long (on average) is worth sharing however few distinct values there are: the zip
        /// compressor only sees a repeat if it is within 32 KB of the last one, and a long text that comes back
        /// every few hundred rows is stored again every time. Shorter text from a small set is not worth the time
        /// (measured: the file is the same size, within 1%, and the write is faster inline).
        /// </summary>
        public const int MinRepeatedLength = 16;

        /// <summary>
        /// Or many distinct values: at this many distinct texts in one sample, the repeats are far apart, so
        /// sharing makes the file smaller and Excel opens it faster, whatever the length.
        /// </summary>
        public const int MinDistinct = 200;

        /// <summary>Cells of a shared column counted before deciding whether it still repeats.</summary>
        public const int MonitorCells = 4096;

        /// <summary>A shared column that repeats less than this percent over <see cref="MonitorCells"/> cells goes inline for good.</summary>
        public const int LeavePercent = 2;
    }

    /// <summary>How <see cref="StringStorage.Auto"/> currently stores the text of one column.</summary>
    internal enum ColumnMode : byte
    {
        /// <summary>The first cells: looked up in the table and counted, until there are enough to judge.</summary>
        Sampling = 0,

        /// <summary>The column repeats its values: every cell goes through the table, and the column is watched.</summary>
        Shared = 1,

        /// <summary>The column does not repeat enough to be worth sharing: written inline from here on, with no lookup.</summary>
        Inline = 2,
    }

    /// <summary>
    /// What <see cref="RowWriter"/> remembers about one column while it decides how to store its text. One
    /// instance per column of the sheet being written, in an array that only grows to the highest column
    /// that holds text (at most 16,385 entries, index 0 unused) and never with the number of rows.
    /// </summary>
    internal struct StringColumnState
    {
        /// <summary>The column's current way of storing text.</summary>
        public ColumnMode Mode;

        /// <summary>Cells looked up in the table in the current window.</summary>
        public int Seen;

        /// <summary>Of <see cref="Seen"/>, how many were already in the table.</summary>
        public int Hits;

        /// <summary>Of <see cref="Hits"/>, the total length in characters of the repeated texts.</summary>
        public int HitChars;

        /// <summary>
        /// Records one cell that went through the table and, when a window of cells is complete, decides: a sampled
        /// column becomes shared or inline (for good), a shared column stays shared while it keeps repeating.
        /// </summary>
        /// <param name="hit">True when the text was already in the table.</param>
        /// <param name="length">Length of the text in characters.</param>
        public void Observe(bool hit, int length)
        {
            Seen++;
            if (hit)
            {
                Hits++;
                HitChars += length;
            }
            if (Mode == ColumnMode.Sampling)
            {
                if (Seen == StringColumnPolicy.EarlyExitCells && Hits == 0)
                {
                    Mode = ColumnMode.Inline;
                    Seen = 0;
                    return;
                }
                if (Seen < StringColumnPolicy.ProbeCells) return;
                var repeats = (long)Hits * 100 >= (long)Seen * StringColumnPolicy.EnterPercent;
                var worthIt = HitChars >= (long)Hits * StringColumnPolicy.MinRepeatedLength || Seen - Hits >= StringColumnPolicy.MinDistinct;
                Mode = repeats && worthIt ? ColumnMode.Shared : ColumnMode.Inline;
            }
            else
            {
                if (Seen < StringColumnPolicy.MonitorCells) return;
                if ((long)Hits * 100 < (long)Seen * StringColumnPolicy.LeavePercent) Mode = ColumnMode.Inline;
            }
            Seen = 0;
            Hits = 0;
            HitChars = 0;
        }
    }
}