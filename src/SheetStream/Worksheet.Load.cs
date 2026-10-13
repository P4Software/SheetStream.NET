using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
        public LoadResult LoadFromCollection<T>(IEnumerable<T> items, LoadOptions? options = null)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            options ??= new LoadOptions();

            if (typeof(IDictionary<string, object?>).IsAssignableFrom(typeof(T)) || typeof(T) == typeof(object) && items is IEnumerable<IDictionary<string, object?>>)
                return LoadDictionaries(items.Cast<IDictionary<string, object?>>(), options);

            var members = MemberMap.For(typeof(T));
            var styles = new CellStyle?[members.Length];
            for (var i = 0; i < members.Length; i++)
            {
                if (members[i].NumberFormat != null)
                {
                    var format = members[i].NumberFormat;
                    styles[i] = Workbook.Styles.Create(o => o.NumberFormat = format);
                }
            }
            var headers = members.Select(m => m.Header).ToArray();
            using (var e = items.GetEnumerator())
            {
                return LoadCore(headers, styles, options, e.MoveNext, i => e.Current == null ? null : members[i].Getter(e.Current!));
            }
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

        /// <summary>Shared row writer behind every LoadFrom* method.</summary>
        private LoadResult LoadCore(string[] headers, CellStyle?[] memberStyles, LoadOptions options, Func<bool> moveNext, Func<int, object?> value)
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
                for (var i = 0; i < headers.Length; i++) row.WriteCell(value(i), styles[i]);
                rows++;
            }

            var lastRow = CurrentRow < firstRow ? firstRow : CurrentRow;
            if (options.TableName != null)
            {
                // A table needs at least one data row; Excel rejects a header-only table.
                if (rows == 0) Row().Skip(options.StartColumn - 1).WriteEmpty();
                lastRow = CurrentRow;
                var table = AddTable(options.TableName, firstRow, options.StartColumn, lastRow, lastColumn, headers);
                table.StyleName = options.TableStyle;
            }
            else if (options.AutoFilter && options.PrintHeaders && headers.Length > 0)
            {
                AutoFilter(firstRow, options.StartColumn, lastColumn);
            }
            return new LoadResult(firstRow, lastRow, options.StartColumn, lastColumn, rows);
        }

        /// <summary>Cached, compiled member accessors per type, shared by the writer and <c>ExcelReader.Read&lt;T&gt;</c>.</summary>
        internal sealed class MemberMap
        {
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

            public static MemberMap[] For(Type type) => Cache.GetOrAdd(type, Build);

            private static MemberMap[] Build(Type type)
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
