using System;
using System.Collections.Generic;
using SheetStream.Internal;

namespace SheetStream.Reading
{
    /// <summary>Reading rows into objects through a compile-time <see cref="ISheetRowMapper{T}"/>.</summary>
    public sealed partial class ExcelReader
    {
        /// <summary>
        /// Maps rows to objects through an explicit <see cref="ISheetRowMapper{T}"/> (for example
        /// <c>Order.SheetMapper</c> generated for a <c>[SheetRow]</c> type). Behaves like
        /// <see cref="Read{T}(string)"/>: the first row is the header, headers are matched case-insensitively,
        /// values are converted to the member type and a failed conversion throws a <see cref="FormatException"/>
        /// naming the row and column. No reflection is used, so this overload is safe for trimmed and
        /// Native AOT applications.
        /// </summary>
        public IEnumerable<T> Read<T>(string sheetName, ISheetRowMapper<T> mapper)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            if (!mapper.CanRead) throw new NotSupportedException($"The mapper for {typeof(T).Name} cannot create instances (the type has no parameterless constructor), so it can only be used for writing.");
            FindSheet(sheetName); // Fail fast on a wrong sheet name instead of on the first MoveNext.
            return ReadMapped(sheetName, mapper);
        }

        private IEnumerable<T> ReadMapped<T>(string sheetName, ISheetRowMapper<T> mapper)
        {
            // Sheet column -> mapper column, resolved once from the header row; -1 for columns nobody wants.
            int[]? columns = null;
            foreach (var row in ReadRows(sheetName))
            {
                if (columns == null)
                {
                    columns = BindColumns(mapper, row);
                    continue;
                }

                var item = mapper.Create();
                var any = false;
                for (var sheetColumn = 0; sheetColumn < columns.Length; sheetColumn++)
                {
                    var column = columns[sheetColumn];
                    if (column < 0) continue;
                    var raw = row[sheetColumn];
                    if (raw == null) continue;
                    any = true;
                    object? converted;
                    try { converted = Convert(raw, mapper.MemberType(column)); }
                    catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException || ex is ArgumentException)
                    {
                        throw new FormatException($"Row {row.Index}, column {CellReference.ColumnLetters(sheetColumn + 1)} ('{mapper.Headers[column]}'): cannot convert '{raw}' to {mapper.MemberType(column).Name}.", ex);
                    }
                    mapper.SetValue(ref item, column, converted);
                }
                // Blank rows are skipped, the same rule the reflection reader applies.
                if (any) yield return item;
            }
        }

        private static int[] BindColumns<T>(ISheetRowMapper<T> mapper, ExcelRow header)
        {
            var columns = new int[header.Values.Count];
            for (var i = 0; i < columns.Length; i++)
            {
                var text = header.GetString(i)?.Trim();
                columns[i] = string.IsNullOrEmpty(text) ? -1 : mapper.FindColumn(text!);
                if (columns[i] >= 0 && !mapper.CanSet(columns[i])) columns[i] = -1;
            }
            return columns;
        }
    }
}
