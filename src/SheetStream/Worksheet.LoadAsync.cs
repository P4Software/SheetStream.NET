#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
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
        /// <see cref="LoadFromCollection{T}(IEnumerable{T}, LoadOptions?)"/>.
        /// </summary>
        /// <remarks>Only available on .NET 8 and later; netstandard2.0 has no built-in IAsyncEnumerable.</remarks>
        /// <example>
        /// <code>
        /// await sheet.LoadFromAsyncEnumerable(db.Orders.AsNoTracking().AsAsyncEnumerable(), new LoadOptions { TableName = "Orders" }, ct);
        /// </code>
        /// </example>
        public async Task<LoadResult> LoadFromAsyncEnumerable<T>(IAsyncEnumerable<T> items, LoadOptions? options = null, CancellationToken cancellationToken = default)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            options ??= new LoadOptions();
            var members = MemberMap.For(typeof(T));
            var memberStyles = new CellStyle?[members.Length];
            for (var i = 0; i < members.Length; i++)
            {
                var format = members[i].NumberFormat;
                if (format != null) memberStyles[i] = Workbook.Styles.Create(o => o.NumberFormat = format);
            }
            var headers = members.Select(m => m.Header).ToArray();

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
                        i => rowsBuffered[index] == null ? null : members[i].Getter(rowsBuffered[index]!));
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
