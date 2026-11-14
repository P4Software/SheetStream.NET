using System.IO;

namespace SheetStream.Internal
{
    /// <summary>
    /// A running total of what one workbook makes this library keep in memory, checked against
    /// <see cref="Reading.ReaderOptions.MaxLoadedPartBytes"/>. Everything that outlives the read of a single
    /// row is charged here: shared strings, style tables, sheet and relationship lists, cells loaded for editing.
    /// </summary>
    /// <remarks>
    /// Capping the unpacked size of a part is not enough on its own, because what a part costs in memory can
    /// be several times its size in the file: a cell format is eighteen bytes of XML and well over a hundred
    /// as an object, and a character is one byte in the file and two in a string. Charging what is actually
    /// kept, across all parts together, turns the limit into a statement about memory instead of about files.
    /// The figures charged are estimates that err on the high side.
    /// </remarks>
    internal sealed class MemoryBudget
    {
        /// <summary>Rough cost of a string object beyond its characters, plus the reference that holds it.</summary>
        private const int StringOverhead = 32;

        private readonly long _limit;
        private readonly string? _what;
        private long _used;

        /// <param name="limit">The most that may be charged, in bytes.</param>
        /// <param name="what">
        /// What is being held, for the error message, when it is something other than the workbook's own
        /// parts: "a DataTable of sheet 'Orders'", say.
        /// </param>
        public MemoryBudget(long limit, string? what = null)
        {
            _limit = limit;
            _what = what;
        }

        /// <summary>Bytes charged so far.</summary>
        public long Used => _used;

        /// <summary>Charges <paramref name="bytes"/> and refuses the workbook once the total passes the limit.</summary>
        public void Charge(long bytes)
        {
            _used += bytes;
            if (_used <= _limit) return;
            if (_what != null)
                throw new InvalidDataException($"{_what} needs more than {_limit:N0} bytes of memory. Process the rows as they are read instead of collecting them, or, if the file is genuine and the memory is available, raise ReaderOptions.MaxLoadedPartBytes.");
            throw new InvalidDataException($"The workbook needs more than {_limit:N0} bytes of memory for the parts that are kept in memory (shared strings, styles, sheet lists, cells loaded for editing). The file is damaged or was built to exhaust memory. If it is genuine, raise ReaderOptions.MaxLoadedPartBytes.");
        }

        /// <summary>Charges a string that is being kept: two bytes a character plus the object itself.</summary>
        public void ChargeText(string? text) => Charge(text == null ? 8 : 2L * text.Length + StringOverhead);

        /// <summary>Charges a string of <paramref name="chars"/> characters that is being kept.</summary>
        public void ChargeText(long chars) => Charge(2L * chars + StringOverhead);
    }
}
