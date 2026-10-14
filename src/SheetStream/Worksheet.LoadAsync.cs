#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using SheetStream.Styles;

namespace SheetStream
{
    public sealed partial class Worksheet
    {
        /// <summary>
        /// Writes an <see cref="IAsyncEnumerable{T}"/> as rows, awaiting each item, so an EF Core query
        /// (<c>AsAsyncEnumerable()</c>) or a paged API streams straight into the sheet without ever being
        /// materialised. Columns, headers and options work exactly like
        /// <see cref="LoadFromCollection{T}(IEnumerable{T}, LoadOptions?)"/>, including the use of a mapper
        /// registered in <see cref="SheetMappers"/> before falling back to reflection.
        /// </summary>
        /// <remarks>Only available on .NET 8 and later; netstandard2.0 has no built-in IAsyncEnumerable.</remarks>
        /// <example>
        /// <code>
        /// await sheet.LoadFromAsyncEnumerable(db.Orders.AsNoTracking().AsAsyncEnumerable(), new LoadOptions { TableName = "Orders" }, ct);
        /// </code>
        /// </example>
        public Task<LoadResult> LoadFromAsyncEnumerable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(IAsyncEnumerable<T> items, LoadOptions? options = null, CancellationToken cancellationToken = default)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (SheetMappers.TryGet<T>(out var mapper)) return LoadFromAsyncEnumerable(items, mapper, options, cancellationToken);
            var members = MemberMap.For(typeof(T));
            var headers = members.Select(m => m.Header).ToArray();
            var memberStyles = NumberFormatStyles(members.Length, i => members[i].NumberFormat);
            return LoadFromAsyncEnumerableCore(items, headers, memberStyles, options ?? new LoadOptions(), cancellationToken, (item, row, styles) =>
            {
                for (var i = 0; i < styles.Length; i++) row.WriteCell(members[i].Getter(item!), styles[i]);
            });
        }

        /// <summary>
        /// Writes an <see cref="IAsyncEnumerable{T}"/> through an explicit <see cref="ISheetRowMapper{T}"/>, so
        /// no reflection is involved: the overload to use in trimmed or Native AOT applications.
        /// </summary>
        public Task<LoadResult> LoadFromAsyncEnumerable<T>(IAsyncEnumerable<T> items, ISheetRowMapper<T> mapper, LoadOptions? options = null, CancellationToken cancellationToken = default)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            var headers = CopyHeaders(mapper);
            var memberStyles = NumberFormatStyles(headers.Length, i => mapper.NumberFormats[i]);
            return LoadFromAsyncEnumerableCore(items, headers, memberStyles, options ?? new LoadOptions(), cancellationToken, mapper.WriteRow);
        }

        private async Task<LoadResult> LoadFromAsyncEnumerableCore<T>(IAsyncEnumerable<T> items, string[] headers, CellStyle?[] memberStyles, LoadOptions options, CancellationToken cancellationToken, Action<T, RowWriter, CellStyle?[]> writeItem)
        {
            // Pull the items asynchronously into a small hand-off: LoadCore is synchronous by design (XmlWriter
            // into a zip entry), so each await happens between rows, never in the middle of one.
            var enumerator = items.GetAsyncEnumerator(cancellationToken);
            try
            {
                var hasCurrent = await enumerator.MoveNextAsync().ConfigureAwait(false);
                var first = true;
                var rowsBuffered = new List<T>(256);
                var firstRow = CurrentRow + 1;
                var dataRows = 0;

                // Write the header (and the first batch) through LoadCore, then continue batch by batch.
                while (true)
                {
                    rowsBuffered.Clear();
                    while (hasCurrent && rowsBuffered.Count < 256)
                    {
                        rowsBuffered.Add(enumerator.Current);
                        hasCurrent = await enumerator.MoveNextAsync().ConfigureAwait(false);
                    }
                    cancellationToken.ThrowIfCancellationRequested();

                    var batchOptions = first ? WithoutTable(options) : WithoutHeaderOrTable(options);
                    var index = -1;
                    var batch = LoadCore(headers, memberStyles, batchOptions,
                        () => ++index < rowsBuffered.Count,
                        (row, styles) =>
                        {
                            // A null item keeps its row with every column empty, like the synchronous path.
                            var item = rowsBuffered[index];
                            if (item == null) WriteEmptyRow(row, styles);
                            else writeItem(item, row, styles);
                        });
                    dataRows += batch.DataRows;
                    first = false;
                    if (!hasCurrent) break;
                }

                // Table and AutoFilter apply to the whole load, so add them once at the end.
                var lastColumn = options.StartColumn + headers.Length - 1;
                var lastRow = Math.Max(firstRow, CurrentRow);
                if (options.TableName != null)
                {
                    if (dataRows == 0) { Row().Skip(options.StartColumn - 1).WriteEmpty(); lastRow = CurrentRow; }
                    AddTable(options.TableName, firstRow, options.StartColumn, lastRow, lastColumn, headers).StyleName = options.TableStyle;
                }
                else if (options.AutoFilter && options.PrintHeaders && headers.Length > 0)
                {
                    AutoFilter(firstRow, options.StartColumn, lastColumn);
                }
                return new LoadResult(firstRow, lastRow, options.StartColumn, lastColumn, dataRows);
            }
            finally
            {
                await enumerator.DisposeAsync().ConfigureAwait(false);
            }
        }

        /// <summary>Copy of the options for the first batch: header as requested, table deferred to the end.</summary>
        private static LoadOptions WithoutTable(LoadOptions o)
        {
            var copy = new LoadOptions { PrintHeaders = o.PrintHeaders, HeaderStyle = o.HeaderStyle, StartColumn = o.StartColumn };
            foreach (var pair in o.ColumnStyles) copy.ColumnStyles[pair.Key] = pair.Value;
            return copy;
        }

        /// <summary>Copy of the options for later batches: no header, no table.</summary>
        private static LoadOptions WithoutHeaderOrTable(LoadOptions o)
        {
            var copy = WithoutTable(o);
            copy.PrintHeaders = false;
            return copy;
        }
    }
}
#endif
