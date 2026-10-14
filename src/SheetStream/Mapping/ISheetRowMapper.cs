using System;
using System.Collections.Generic;
using SheetStream.Styles;

namespace SheetStream
{
    /// <summary>
    /// Describes how a type maps to worksheet columns and moves values in both directions without
    /// reflection. <c>LoadFromCollection</c> uses it to write rows and <c>ExcelReader.Read&lt;T&gt;</c> to
    /// bind them back. Implementations are generated at compile time for types marked with
    /// <see cref="SheetRowAttribute"/>, but a hand-written mapper works just as well; register it with
    /// <see cref="SheetMappers.Register{T}"/> or pass it to the explicit overloads.
    /// </summary>
    /// <typeparam name="T">The row type.</typeparam>
    /// <remarks>
    /// Columns are zero-based and ordered exactly as they appear in the sheet: by <c>[SheetColumn(Order)]</c>
    /// ascending, then declaration order. A mapper must be stateless and safe to share between threads.
    /// </remarks>
    public interface ISheetRowMapper<T>
    {
        /// <summary>Header text per column, in column order.</summary>
        IReadOnlyList<string> Headers { get; }

        /// <summary>Excel number format per column (from <c>[SheetColumn(NumberFormat)]</c>), or null for the default.</summary>
        IReadOnlyList<string?> NumberFormats { get; }

        /// <summary>True when <see cref="Create"/> can build a new instance, so the mapper can be used for reading.</summary>
        bool CanRead { get; }

        /// <summary>The declared type of the member behind a column, used to convert cell values when reading.</summary>
        Type MemberType(int column);

        /// <summary>True when <see cref="SetValue"/> can assign the column (false for read-only, init-only or get-only members).</summary>
        bool CanSet(int column);

        /// <summary>
        /// Resolves a header read from a sheet to a column index, or -1 when no settable member matches.
        /// Matching is case-insensitive and accepts the <c>[SheetColumn]</c> name, the <c>[DisplayName]</c>
        /// and the member name, like the reflection-based reader does.
        /// </summary>
        int FindColumn(string header);

        /// <summary>
        /// Writes every column of <paramref name="item"/> into <paramref name="row"/> using the typed
        /// <see cref="RowWriter.WriteCell(double, CellStyle?)"/> overloads, so no value is boxed per cell.
        /// <paramref name="styles"/> has one entry per column (null for the default style).
        /// </summary>
        void WriteRow(T item, RowWriter row, CellStyle?[] styles);

        /// <summary>Reads one member as an object; a convenience for callers that need values outside a row writer.</summary>
        object? GetValue(T item, int column);

        /// <summary>Creates an empty instance to fill while reading. Throws <see cref="NotSupportedException"/> when <see cref="CanRead"/> is false.</summary>
        T Create();

        /// <summary>
        /// Assigns an already converted value (of <see cref="MemberType"/>, or null) to a column. Columns where
        /// <see cref="CanSet"/> is false are ignored. <paramref name="item"/> is passed by reference so the
        /// same mapper works for structs.
        /// </summary>
        void SetValue(ref T item, int column, object? value);
    }
}
