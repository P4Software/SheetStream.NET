using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using SheetStream.Styles;

namespace SheetStream
{
    public sealed partial class Worksheet
    {
        /// <summary>
        /// Writes a collection as rows, one column per public property, starting on the next row. Header text
        /// comes from <see cref="SheetColumnAttribute"/>, <see cref="DisplayNameAttribute"/> or the property
        /// name; <c>IDictionary&lt;string, object&gt;</c> items (ExpandoObject, Dapper rows) use the first
        /// item's keys. Values stream one row at a time, so any size of enumerable is fine.
        /// </summary>
        /// <example>
        /// <code>
        /// sheet.LoadFromCollection(orders, new LoadOptions { HeaderStyle = header, TableName = "Orders" });
        /// </code>
        /// </example>
        /// <remarks>
        /// A mapper registered in <see cref="SheetMappers"/> (generated for <c>[SheetRow]</c> types) is used
        /// when present; otherwise the columns are discovered through reflection, which needs the public
        /// properties and fields of <typeparamref name="T"/> to survive trimming.
        /// </remarks>
        public LoadResult LoadFromCollection<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(IEnumerable<T> items, LoadOptions? options = null)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (SheetMappers.TryGet<T>(out var mapper)) return LoadFromCollection(items, mapper, options);
            options ??= new LoadOptions();

            if (typeof(IDictionary<string, object?>).IsAssignableFrom(typeof(T)) || typeof(T) == typeof(object) && items is IEnumerable<IDictionary<string, object?>>)
                return LoadDictionaries(items.Cast<IDictionary<string, object?>>(), options);

            var members = MemberMap.For(typeof(T));
            var styles = NumberFormatStyles(members.Length, i => members[i].NumberFormat);
            var headers = members.Select(m => m.Header).ToArray();
            using (var e = items.GetEnumerator())
            {
                return LoadCore(headers, styles, options, e.MoveNext, i => e.Current == null ? null : members[i].Getter(e.Current!));
            }
        }

        /// <summary>
        /// Writes a collection as rows through an explicit <see cref="ISheetRowMapper{T}"/> (for example
        /// <c>Order.SheetMapper</c> generated for a <c>[SheetRow]</c> type). No reflection is involved, so this
        /// overload is safe for trimmed and Native AOT applications and never boxes a cell value.
        /// </summary>
        public LoadResult LoadFromCollection<T>(IEnumerable<T> items, ISheetRowMapper<T> mapper, LoadOptions? options = null)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            options ??= new LoadOptions();
            var headers = CopyHeaders(mapper);
            var styles = NumberFormatStyles(headers.Length, i => mapper.NumberFormats[i]);
            using (var e = items.GetEnumerator())
            {
                // Null items keep their row (all columns empty) exactly like the reflection path does.
                return LoadCore(headers, styles, options, e.MoveNext, (row, s) => { if (e.Current != null) mapper.WriteRow(e.Current, row, s); else WriteEmptyRow(row, s); });
            }
        }

        /// <summary>Snapshot of the mapper's headers; the writer indexes them in a tight loop and must not observe later changes.</summary>
        private static string[] CopyHeaders<T>(ISheetRowMapper<T> mapper)
        {
            var headers = new string[mapper.Headers.Count];
            for (var i = 0; i < headers.Length; i++) headers[i] = mapper.Headers[i] ?? throw new InvalidOperationException($"Mapper for {typeof(T).Name} returned a null header for column {i}.");
            if (mapper.NumberFormats.Count != headers.Length)
                throw new InvalidOperationException($"Mapper for {typeof(T).Name} has {headers.Length} headers but {mapper.NumberFormats.Count} number formats.");
            return headers;
        }

        /// <summary>One style per column for the columns that carry a number format, null elsewhere.</summary>
        private CellStyle?[] NumberFormatStyles(int count, Func<int, string?> numberFormat)
        {
            var styles = new CellStyle?[count];
            for (var i = 0; i < count; i++)
            {
                var format = numberFormat(i);
                if (format != null) styles[i] = Workbook.Styles.Create(o => o.NumberFormat = format);
            }
            return styles;
        }

        private static void WriteEmptyRow(RowWriter row, CellStyle?[] styles)
        {
            foreach (var style in styles) row.WriteEmpty(style);
        }

        /// <summary>Writes a <see cref="DataTable"/> as rows, with column captions as headers.</summary>
        public LoadResult LoadFromDataTable(DataTable table, LoadOptions? options = null)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            options ??= new LoadOptions();
            var headers = new string[table.Columns.Count];
            for (var i = 0; i < headers.Length; i++)
                headers[i] = string.IsNullOrEmpty(table.Columns[i].Caption) ? table.Columns[i].ColumnName : table.Columns[i].Caption;
            var index = -1;
            return LoadCore(headers, new CellStyle?[headers.Length], options,
                () => ++index < table.Rows.Count,
                i => table.Rows[index][i]);
        }

        /// <summary>
        /// Writes rows given as arrays, the direct replacement for EPPlus' <c>Cells.LoadFromArrays</c>. With
        /// <see cref="LoadOptions.PrintHeaders"/> (the default) the first array is the header row and gets
        /// <see cref="LoadOptions.HeaderStyle"/>; every later array is a data row. Rows may be shorter than the
        /// header (missing cells stay empty); longer rows have their extra values dropped, because a table or
        /// filter needs a fixed column count. Values are written like <see cref="RowWriter.WriteCell(object?, CellStyle?)"/>.
        /// </summary>
        /// <example>
        /// <code>
        /// sheet.LoadFromArrays(new[] { new object?[] { "Sku", "Qty" }, new object?[] { "A-1", 3 } }, new LoadOptions { TableName = "Stock" });
        /// </code>
        /// </example>
        public LoadResult LoadFromArrays(IEnumerable<object?[]> rows, LoadOptions? options = null)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            options ??= new LoadOptions();
            using (var e = rows.GetEnumerator())
            {
                if (!e.MoveNext())
                    return new LoadResult(CurrentRow + 1, CurrentRow, options.StartColumn, options.StartColumn - 1, 0);

                // The first array sizes the sheet. Without headers it is data, so it must be written too.
                var first = e.Current ?? new object?[0];
                var headers = new string[first.Length];
                for (var i = 0; i < headers.Length; i++)
                    headers[i] = options.PrintHeaders
                        ? (first[i] == null ? Internal.CellReference.ColumnLetters(options.StartColumn + i) : Convert.ToString(first[i], System.Globalization.CultureInfo.InvariantCulture) ?? "")
                        : Internal.CellReference.ColumnLetters(options.StartColumn + i);
                var pendingFirst = !options.PrintHeaders;
                return LoadCore(headers, new CellStyle?[headers.Length], options,
                    () => { if (pendingFirst) { pendingFirst = false; return true; } return e.MoveNext(); },
                    i => e.Current != null && i < e.Current.Length ? e.Current[i] : null);
            }
        }

        /// <summary>
        /// Writes everything an <see cref="IDataReader"/> returns, streaming row by row; the natural fit for
        /// exporting a large SQL query without materialising it.
        /// </summary>
        public LoadResult LoadFromDataReader(IDataReader reader, LoadOptions? options = null)
        {
            if (reader == null) throw new ArgumentNullException(nameof(reader));
            options ??= new LoadOptions();
            var headers = new string[reader.FieldCount];
            for (var i = 0; i < headers.Length; i++) headers[i] = reader.GetName(i);
            return LoadCore(headers, new CellStyle?[headers.Length], options, reader.Read, i => reader.IsDBNull(i) ? null : reader.GetValue(i));
        }

        private LoadResult LoadDictionaries(IEnumerable<IDictionary<string, object?>> items, LoadOptions options)
        {
            using (var e = items.GetEnumerator())
            {
                if (!e.MoveNext())
                    return new LoadResult(CurrentRow + 1, CurrentRow, options.StartColumn, options.StartColumn - 1, 0);
                var headers = e.Current.Keys.ToArray();
                var first = true;
                return LoadCore(headers, new CellStyle?[headers.Length], options,
                    () => { if (first) { first = false; return true; } return e.MoveNext(); },
                    i => e.Current.TryGetValue(headers[i], out var v) ? v : null);
            }
        }

        /// <summary>Row writer for sources that surface values as objects (DataTable, IDataReader, dictionaries, reflection).</summary>
        private LoadResult LoadCore(string[] headers, CellStyle?[] memberStyles, LoadOptions options, Func<bool> moveNext, Func<int, object?> value)
            => LoadCore(headers, memberStyles, options, moveNext, (row, styles) =>
            {
                for (var i = 0; i < styles.Length; i++) row.WriteCell(value(i), styles[i]);
            });

        /// <summary>
        /// Shared row writer behind every LoadFrom* method. <paramref name="writeRow"/> receives the open row
        /// (already positioned at the start column) and the effective per-column styles, and writes exactly
        /// one cell per header; a mapper does this with typed WriteCell calls so nothing is boxed.
        /// </summary>
        private LoadResult LoadCore(string[] headers, CellStyle?[] memberStyles, LoadOptions options, Func<bool> moveNext, Action<RowWriter, CellStyle?[]> writeRow)
        {
            if (options.StartColumn < 1 || options.StartColumn + Math.Max(headers.Length, 1) - 1 > Internal.CellReference.MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(options), "StartColumn plus the column count must fit within 16384 columns.");
            if (options.TableName != null && !options.PrintHeaders)
                throw new ArgumentException("A table needs a header row; set PrintHeaders = true or leave TableName empty.", nameof(options));

            var styles = new CellStyle?[headers.Length];
            for (var i = 0; i < headers.Length; i++)
                styles[i] = options.ColumnStyles.TryGetValue(headers[i], out var s) ? s : memberStyles[i];

            var firstRow = CurrentRow + 1;
            var lastColumn = options.StartColumn + headers.Length - 1;
            if (options.PrintHeaders)
            {
                var header = Row();
                header.Skip(options.StartColumn - 1);
                foreach (var h in headers) header.WriteCell(h, options.HeaderStyle);
            }

            var rows = 0;
            while (moveNext())
            {
                var row = Row();
                row.Skip(options.StartColumn - 1);
                writeRow(row, styles);
                rows++;
            }

            var lastRow = CurrentRow < firstRow ? firstRow : CurrentRow;
            ExcelTable? table = null;
            if (options.TableName != null)
            {
                // A table needs at least one data row; Excel rejects a header-only table.
                if (rows == 0) Row().Skip(options.StartColumn - 1).WriteEmpty();
                lastRow = CurrentRow;
                table = AddTable(options.TableName, firstRow, options.StartColumn, lastRow, lastColumn, headers);
                table.StyleName = options.TableStyle;
            }
            else if (options.AutoFilter && options.PrintHeaders && headers.Length > 0)
            {
                AutoFilter(firstRow, options.StartColumn, lastColumn);
            }
            return new LoadResult(firstRow, lastRow, options.StartColumn, lastColumn, rows, (string[])headers.Clone(), table);
        }

        /// <summary>
        /// Cached, compiled member accessors per type: the reflection fallback used when no
        /// <see cref="ISheetRowMapper{T}"/> is registered. Every entry point is annotated so the trimmer keeps
        /// the public properties and fields of the types that flow in here.
        /// </summary>
        internal sealed class MemberMap
        {
            /// <summary>The members the reflection path looks at; shared with ExcelReader so both sides agree.</summary>
            internal const DynamicallyAccessedMemberTypes RequiredMembers = DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields;

            private static readonly ConcurrentDictionary<Type, MemberMap[]> Cache = new ConcurrentDictionary<Type, MemberMap[]>();

            private MemberMap(MemberInfo member, Type type, string header, string? numberFormat, Func<object, object?> getter)
            {
                Member = member;
                Type = type;
                Header = header;
                NumberFormat = numberFormat;
                Getter = getter;
            }

            public MemberInfo Member { get; }
            public Type Type { get; }
            public string Header { get; }
            public string? NumberFormat { get; }
            public Func<object, object?> Getter { get; }

            public static MemberMap[] For([DynamicallyAccessedMembers(RequiredMembers)] Type type)
            {
                // Not GetOrAdd(type, Build): handing an annotated method to a Func<Type, ...> loses the
                // annotation and the trim analyzer flags it (IL2111). Building twice under a race is harmless.
                return Cache.TryGetValue(type, out var cached) ? cached : Cache.GetOrAdd(type, Build(type));
            }

            private static MemberMap[] Build([DynamicallyAccessedMembers(RequiredMembers)] Type type)
            {
                var candidates = new List<(MemberMap Map, int Order, int Token)>();
                const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance;
                foreach (var m in type.GetProperties(flags).Cast<MemberInfo>().Concat(type.GetFields(flags)))
                {
                    if (m is PropertyInfo p && (!p.CanRead || p.GetIndexParameters().Length > 0)) continue;
                    var attr = m.GetCustomAttribute<SheetColumnAttribute>();
                    if (attr?.Ignore == true) continue;
                    var memberType = m is PropertyInfo pi ? pi.PropertyType : ((FieldInfo)m).FieldType;
                    var header = attr?.Name ?? m.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? m.Name;

                    // Compile obj => (object)((T)obj).Member once; reflection GetValue per cell would dominate large exports.
                    var obj = Expression.Parameter(typeof(object), "obj");
                    var access = Expression.MakeMemberAccess(Expression.Convert(obj, type), m);
                    var getter = Expression.Lambda<Func<object, object?>>(Expression.Convert(access, typeof(object)), obj).Compile();

                    candidates.Add((new MemberMap(m, memberType, header, attr?.NumberFormat, getter), attr?.Order ?? int.MaxValue, m.MetadataToken));
                }
                return candidates.OrderBy(c => c.Order).ThenBy(c => c.Token).Select(c => c.Map).ToArray();
            }
        }
    }
}
