using System;
using System.Runtime.CompilerServices;

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
        /// A column in which not one of the first cells repeated looks unique (ids, e-mail addresses, free text) and is
        /// judged at once, which spares such a column most of the sampling: each sampled text is hashed, stored in the
        /// table and kept until the save. It is not judged for good: <see cref="FirstLookAt"/> gives it a second look. A
        /// sheet of fewer rows than this is stored exactly as <see cref="StringStorage.Shared"/> stores it.
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

        // ---- the second look -----------------------------------------------------------------------------------------
        //
        // A column judged unique is not necessarily unique: it may repeat among thousands of distinct values (customer
        // names, product names), too far apart to show in its first cells. Its repeats are also too far apart for the zip
        // compressor, whose window is 32 KB, so written inline it costs several times the bytes (measured: 3 to 6 times
        // the file) and more time than a shared column. So a column judged unique is looked at again, twice, and
        // is shared if the repeats it shows mean that its different values do not fit the compressor's window.

        /// <summary>How many cells of a column count, from the start, before its first second look begins. It begins at once for a column that was sampled to the end of <see cref="ProbeCells"/>.</summary>
        public const int FirstLookAt = 512;

        /// <summary>How many cells of a column count before the last second look begins (a column that repeats among more values than the first could see is caught here).</summary>
        public const int LastLookAt = 8192;

        /// <summary>How many cells one second look looks up in the table. The cells are written as shared strings like the first sample, so no text in the table is left without a cell.</summary>
        public const int LookCells = 512;

        /// <summary>A look shows a repeat only when this many cells repeated an earlier text: a column of unique text has none, so a few are already a signal.</summary>
        public const int MinLookHits = 3;

        /// <summary>
        /// A column is shared only if the texts it repeats cannot all be remembered by the zip compressor: the estimated
        /// number of different values times their average length (the vocabulary) is more than this many characters. The
        /// compressor's window is 32 KB; this is a margin under it. A vocabulary that fits in the window is folded away
        /// by the compressor, and measured on 1,920 customer names (13 characters, 25 KB) sharing them made the file
        /// larger at the default zip level and the write slower, because the names are made of a few dozen words that the
        /// compressor finds again and again, while the table indexes are digits that it cannot shorten.
        /// </summary>
        public const int WindowBytes = 24 * 1024;

        /// <summary>
        /// A column is shared only if, by the estimate, at least one cell in this many repeats an earlier one so far;
        /// columns of nearly unique text with a few accidental repeats are not worth the table. (The check of a shared
        /// column, <see cref="LeavePercent"/>, still takes a wrong choice back.)
        /// </summary>
        public const int MaxCellsPerRepeat = 100;
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

        /// <summary>Judged unique so far: written inline, with no lookup, while the cells are counted until the next second look.</summary>
        Waiting = 3,

        /// <summary>A second look: the cells go through the table, written as shared strings, and are counted to estimate how many different values the column has.</summary>
        Looking = 4,
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

        /// <summary>Second looks begun so far.</summary>
        public byte Looks;

        /// <summary>Cells looked up in the table in the current window.</summary>
        public int Seen;

        /// <summary>Of <see cref="Seen"/>, how many were already in the table.</summary>
        public int Hits;

        /// <summary>Of <see cref="Hits"/>, the total length in characters of the repeated texts.</summary>
        public int HitChars;

        /// <summary>Of <see cref="Seen"/>, the total length in characters of the texts: with the estimate of the number of different values it gives the size of the vocabulary.</summary>
        public int Chars;

        /// <summary>Cells of the column counted so far while it was sampled, waiting or looking.</summary>
        public int Total;

        /// <summary>Different texts this column has put in the table (the ones that were new when they came).</summary>
        public int Known;

        /// <summary>
        /// In a second look: for every cell looked up, the texts of this column already in the table, added up. It is
        /// the number of pairs of cells that could have been equal; with <see cref="Hits"/> of them that were, one pair
        /// in <c>Pairs / Hits</c> is, which for a column that draws at random from <c>n</c> values is <c>n</c> (the
        /// birthday estimate of the number of different values).
        /// </summary>
        public long Pairs;

        /// <summary>
        /// True when this cell is written inline without a lookup: the column is judged inline, or it is waiting for
        /// its next look, which this cell then counts (and when the count is reached, the column starts looking and
        /// the cell is not inline). The caller must call it once for every text cell of the column.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool WritesInline() => Mode == ColumnMode.Inline || (Mode == ColumnMode.Waiting && StaysWaiting());

        private bool StaysWaiting()
        {
            Total++;
            if (Total < (Looks == 0 ? StringColumnPolicy.FirstLookAt : StringColumnPolicy.LastLookAt)) return true;
            Mode = ColumnMode.Looking;
            Looks++;
            Seen = 0;
            Hits = 0;
            HitChars = 0;
            Chars = 0;
            Pairs = 0;
            return false;
        }

        /// <summary>
        /// Records one cell that went through the table and, when a window of cells is complete, decides: a sampled
        /// column becomes shared, inline for good or waiting for a second look, a column that is looking becomes shared or
        /// goes back to waiting (or is inline for good after its last look), and a shared column stays shared while it
        /// keeps repeating.
        /// </summary>
        /// <param name="hit">True when the text was already in the table.</param>
        /// <param name="length">Length of the text in characters.</param>
        public void Observe(bool hit, int length)
        {
            Seen++;
            Total++;
            Chars += length;
            if (Mode == ColumnMode.Looking) Pairs += Known;
            if (hit)
            {
                Hits++;
                HitChars += length;
            }
            else Known++;

            if (Mode == ColumnMode.Sampling)
            {
                if (Seen == StringColumnPolicy.EarlyExitCells && Hits == 0)
                {
                    Mode = ColumnMode.Waiting;
                    Seen = 0;
                    Chars = 0;
                    return;
                }
                if (Seen < StringColumnPolicy.ProbeCells) return;
                var repeats = (long)Hits * 100 >= (long)Seen * StringColumnPolicy.EnterPercent;
                Mode = !repeats ? ColumnMode.Waiting : WorthSharing() ? ColumnMode.Shared : ColumnMode.Inline;
            }
            else if (Mode == ColumnMode.Looking)
            {
                if (Seen < StringColumnPolicy.LookCells) return;
                if (RepeatsBeyondTheWindow()) Mode = ColumnMode.Shared;
                else Mode = Looks >= 2 ? ColumnMode.Inline : ColumnMode.Waiting;
            }
            else
            {
                if (Seen < StringColumnPolicy.MonitorCells) return;
                if ((long)Hits * 100 < (long)Seen * StringColumnPolicy.LeavePercent) Mode = ColumnMode.Inline;
            }
            Seen = 0;
            Hits = 0;
            HitChars = 0;
            Chars = 0;
            Pairs = 0;
        }

        /// <summary>Whether the repeats of the window are worth the table: long text, or many different values.</summary>
        private bool WorthSharing() =>
            HitChars >= (long)Hits * StringColumnPolicy.MinRepeatedLength || Seen - Hits >= StringColumnPolicy.MinDistinct;

        /// <summary>
        /// The verdict of a second look. With <c>n = Pairs / Hits</c> the estimated number of different values, the column
        /// is shared if it does repeat (a few hits, and by the estimate one cell in a hundred of those so far) and if its
        /// vocabulary, <c>n</c> values of the average length in the window, is more than the compressor's window can hold, so
        /// that the compressor cannot fold the repeats away by itself. All in whole numbers: <c>n * Chars / Seen &gt;= WindowBytes</c>
        /// is <c>Pairs * Chars &gt;= WindowBytes * Hits * Seen</c>. (A vocabulary of at least 24 KB means long text or at least
        /// a thousand different values, so no separate test of whether the repeats are worth the table is needed.)
        /// </summary>
        private bool RepeatsBeyondTheWindow() =>
            Hits >= StringColumnPolicy.MinLookHits
            && Pairs <= (long)StringColumnPolicy.MaxCellsPerRepeat * Total * Hits / 2
            && Pairs * Chars >= (long)StringColumnPolicy.WindowBytes * Hits * Seen;
    }
}