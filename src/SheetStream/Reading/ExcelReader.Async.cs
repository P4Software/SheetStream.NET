#if NET8_0_OR_GREATER
using System;
using System.Buffers;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace SheetStream.Reading
{
    /// <summary>
    /// Asynchronous reading: the counterpart of <c>Worksheet.LoadFromAsyncEnumerable</c>. An upload can be
    /// buffered with <see cref="OpenAsync(Stream, bool, CancellationToken)"/> and its rows pulled with
    /// <see cref="ReadRowsAsync(string, CancellationToken)"/> without ever blocking a request thread on I/O.
    /// Only available on .NET 8 and later; netstandard2.0 has no built-in <see cref="IAsyncEnumerable{T}"/>.
    /// </summary>
    public sealed partial class ExcelReader
    {
        /// <summary>
        /// How much of a non-seekable stream <see cref="OpenAsync(Stream, bool, CancellationToken)"/> keeps in
        /// memory before spilling the rest into a temporary file: 64 MB. Below the limit the whole upload sits
        /// in a <see cref="MemoryStream"/>; above it the bytes already buffered plus everything that follows go
        /// to a file opened with <see cref="FileOptions.DeleteOnClose"/>, so a multi-hundred-megabyte upload
        /// never doubles the process' managed heap and the file disappears when the reader is disposed.
        /// </summary>
        internal const long InMemoryBufferLimit = 64L * 1024 * 1024;

        /// <summary>
        /// Opens a workbook from a stream, buffering asynchronously when the stream cannot seek. A seekable stream
        /// (a <see cref="FileStream"/> or <see cref="MemoryStream"/>) is wrapped directly, exactly like
        /// <see cref="Open(Stream, bool)"/>. A non-seekable one (an ASP.NET Core request body, an
        /// <c>HttpClient</c> response) must be copied first because the zip directory sits at the end of the
        /// file; the copy uses <c>CopyToAsync</c>-style reads so no thread is blocked on the network, and it
        /// stays in memory up to <see cref="InMemoryBufferLimit"/> (64 MB) before spilling into a temporary file
        /// that is deleted with the reader.
        /// </summary>
        /// <remarks>
        /// Once the bytes are in place the package directory, workbook, shared strings and styles are parsed
        /// synchronously: <see cref="ZipArchive"/> has no asynchronous API and those parts are small compared
        /// with the sheets. The row data itself is read asynchronously by <see cref="ReadRowsAsync(string, CancellationToken)"/>.
        /// </remarks>
        /// <param name="stream">The .xlsx content.</param>
        /// <param name="leaveOpen">True to leave <paramref name="stream"/> open when the reader is disposed (or, for a non-seekable stream, after it has been copied).</param>
        /// <param name="cancellationToken">Cancels the copy of a non-seekable stream.</param>
        public static Task<ExcelReader> OpenAsync(Stream stream, bool leaveOpen = false, CancellationToken cancellationToken = default)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (stream.CanSeek) return Task.FromResult(new ExcelReader(stream, leaveOpen));
            return OpenBufferedAsync(stream, leaveOpen, cancellationToken);
        }

        private static async Task<ExcelReader> OpenBufferedAsync(Stream stream, bool leaveOpen, CancellationToken cancellationToken)
        {
            var copy = await BufferAsync(stream, InMemoryBufferLimit, cancellationToken).ConfigureAwait(false);
            try
            {
                // Same order as the synchronous Open: the source is released before the package is parsed.
                if (!leaveOpen) await stream.DisposeAsync().ConfigureAwait(false);
                return new ExcelReader(copy, false);
            }
            catch
            {
                await copy.DisposeAsync().ConfigureAwait(false); // Also deletes a spilled temp file.
                throw;
            }
        }

        /// <summary>
        /// Opens a workbook file for asynchronous reading. The file is opened for overlapped I/O
        /// (<see cref="FileOptions.Asynchronous"/>) so that <see cref="ReadRowsAsync(string, CancellationToken)"/>
        /// pulls sheet bytes with true asynchronous reads instead of thread-pool emulation. Opening itself is not
        /// fully asynchronous: the zip directory, workbook, shared strings and styles are read synchronously
        /// because <see cref="ZipArchive"/> has no asynchronous API; the returned task only surfaces I/O errors
        /// (a missing file, say) the way an async caller expects. Like <see cref="Open(string)"/>, the file is
        /// shared for reading so Excel can keep it open.
        /// </summary>
        public static Task<ExcelReader> OpenAsync(string path, CancellationToken cancellationToken = default)
        {
            if (path == null) throw new ArgumentNullException(nameof(path));
            if (cancellationToken.IsCancellationRequested) return Task.FromCanceled<ExcelReader>(cancellationToken);
            try
            {
                var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1 << 16, FileOptions.Asynchronous);
                try { return Task.FromResult(new ExcelReader(fs, false)); }
                catch { fs.Dispose(); throw; }
            }
            catch (Exception ex) when (!(ex is ArgumentException))
            {
                // Argument errors are programming mistakes and throw synchronously; everything else (file not
                // found, access denied, a corrupt package) is reported through the task like any async I/O call.
                return Task.FromException<ExcelReader>(ex);
            }
        }

        /// <summary>
        /// Copies a non-seekable stream into a seekable one positioned at zero: a <see cref="MemoryStream"/>
        /// while the content fits in <paramref name="memoryLimit"/> bytes, otherwise a temporary file opened
        /// with <see cref="FileOptions.DeleteOnClose"/>. Internal so tests can exercise the spill with a small limit.
        /// </summary>
        internal static async Task<Stream> BufferAsync(Stream source, long memoryLimit, CancellationToken cancellationToken)
        {
            var memory = new MemoryStream();
            // 80 KB keeps the scratch buffer below the large-object threshold; the pool makes it free to rent.
            var buffer = ArrayPool<byte>.Shared.Rent(81920);
            try
            {
                int read;
                while ((read = await source.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false)) > 0)
                {
                    if (memory.Length + read > memoryLimit)
                    {
                        return await SpillToFileAsync(memory, buffer.AsMemory(0, read), source, cancellationToken).ConfigureAwait(false);
                    }
                    memory.Write(buffer, 0, read);
                }
                memory.Position = 0;
                return memory;
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        /// <summary>Moves what is already buffered plus the rest of <paramref name="source"/> into a self-deleting temp file.</summary>
        private static async Task<Stream> SpillToFileAsync(MemoryStream buffered, ReadOnlyMemory<byte> pending, Stream source, CancellationToken cancellationToken)
        {
            // GetRandomFileName + CreateNew instead of GetTempFileName: the latter caps at 65,535 files per
            // folder and creates the file twice. DeleteOnClose ties the file's life to the reader's.
            var path = Path.Combine(Path.GetTempPath(), "sheetstream-" + Path.GetRandomFileName());
            var file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 1 << 16, FileOptions.Asynchronous | FileOptions.DeleteOnClose);
            try
            {
                await file.WriteAsync(buffered.GetBuffer().AsMemory(0, (int)buffered.Length), cancellationToken).ConfigureAwait(false);
                await file.WriteAsync(pending, cancellationToken).ConfigureAwait(false);
                await source.CopyToAsync(file, cancellationToken).ConfigureAwait(false);
                await file.FlushAsync(cancellationToken).ConfigureAwait(false);
                file.Position = 0;
                return file;
            }
            catch
            {
                await file.DisposeAsync().ConfigureAwait(false);
                throw;
            }
        }

        // ------------------------------------------------------------------ rows

        /// <summary>
        /// Streams the rows of a sheet by name (case-insensitive) without blocking a thread: the sheet XML is
        /// parsed with the asynchronous <see cref="XmlReader"/> methods and cancellation is checked before every
        /// row. Values are identical to <see cref="ReadRows(string)"/>; the two share the cell parsing code.
        /// Breaking out of the <c>await foreach</c> early disposes the XML reader and the zip entry stream.
        /// </summary>
        /// <example>
        /// <code>
        /// using var reader = await ExcelReader.OpenAsync(Request.Body, cancellationToken: ct);
        /// await foreach (var row in reader.ReadRowsAsync("Products", ct))
        ///     Console.WriteLine($"{row.GetString(0)} = {row.GetDouble(2)}");
        /// </code>
        /// </example>
        public IAsyncEnumerable<ExcelRow> ReadRowsAsync(string sheetName, CancellationToken cancellationToken = default)
            => ReadRowsAsync(FindSheet(sheetName), cancellationToken);

        /// <summary>Streams the rows of a sheet by zero-based position; see <see cref="ReadRowsAsync(string, CancellationToken)"/>.</summary>
        public IAsyncEnumerable<ExcelRow> ReadRowsAsync(int sheetIndex, CancellationToken cancellationToken = default)
            => ReadRowsCoreAsync(SheetEntry(sheetIndex), cancellationToken);

        /// <summary>
        /// The asynchronous twin of <see cref="ReadRowsCore"/>: same cursor discipline, same helpers for placing
        /// and typing cells, with every <c>Read</c> replaced by its awaited form. The token passed here is
        /// combined by the compiler with one supplied later through <c>WithCancellation</c>.
        /// </summary>
        private async IAsyncEnumerable<ExcelRow> ReadRowsCoreAsync(ZipArchiveEntry entry, [EnumeratorCancellation] CancellationToken cancellationToken)
        {
            // A plain using is enough: XmlReader is IDisposable only, and the async iterator runs this finally
            // block both on normal completion and when the consumer disposes the enumerator early.
            using (var xr = XmlReader.Create(entry.Open(), SheetSettings(async: true)))
            {
                var lastRow = 0;
                var values = new List<object?>();
                while (await xr.ReadAsync().ConfigureAwait(false))
                {
                    if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "row") continue;
                    // Between rows is the one place cancellation is honoured: a row is parsed whole or not at all.
                    cancellationToken.ThrowIfCancellationRequested();
                    var rowIndex = ParseInt(xr.GetAttribute("r")) ?? lastRow + 1;
                    lastRow = rowIndex;
                    values.Clear();
                    if (!xr.IsEmptyElement)
                    {
                        var depth = xr.Depth;
                        var nextColumn = 1;
                        await xr.ReadAsync().ConfigureAwait(false);
                        while (!xr.EOF && !(xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth))
                        {
                            if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "c")
                            {
                                var reference = xr.GetAttribute("r");
                                var column = reference != null ? ColumnOf(reference) : nextColumn;
                                nextColumn = column + 1;
                                var value = await ReadCellAsync(xr).ConfigureAwait(false); // Leaves the reader on </c> or on an empty <c/>.
                                await xr.ReadAsync().ConfigureAwait(false);
                                if (value != null) PlaceCell(values, column, value);
                                continue;
                            }
                            await xr.ReadAsync().ConfigureAwait(false);
                        }
                    }
                    yield return new ExcelRow(rowIndex, values.ToArray());
                }
            }
        }

        /// <summary>
        /// Asynchronous twin of <see cref="ReadCell"/>: collects the raw parts of a <c>c</c> element and hands them
        /// to the shared <see cref="CellValue"/>. Returns a <see cref="ValueTask{T}"/> because the XML buffer
        /// usually already holds the cell, so most calls complete synchronously without allocating a task.
        /// </summary>
        private async ValueTask<object?> ReadCellAsync(XmlReader xr)
        {
            var type = xr.GetAttribute("t");
            var styleIndex = ParseInt(xr.GetAttribute("s")) ?? 0;
            if (xr.IsEmptyElement) return null;

            string? v = null;
            string? inline = null;
            var depth = xr.Depth;
            await xr.ReadAsync().ConfigureAwait(false);
            while (!xr.EOF && !(xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth))
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    if (xr.LocalName == "v")
                    {
                        if (xr.IsEmptyElement) { v = ""; await xr.ReadAsync().ConfigureAwait(false); }
                        else v = await xr.ReadElementContentAsStringAsync().ConfigureAwait(false); // Advances past </v>.
                        continue;
                    }
                    if (xr.LocalName == "is")
                    {
                        if (xr.IsEmptyElement) inline = "";
                        else inline = await ReadRichTextAsync(xr).ConfigureAwait(false); // Leaves the reader on </is>.
                        await xr.ReadAsync().ConfigureAwait(false);
                        continue;
                    }
                    await xr.SkipAsync().ConfigureAwait(false); // <f> and anything else: skip the whole element.
                    continue;
                }
                await xr.ReadAsync().ConfigureAwait(false);
            }

            return CellValue(type, styleIndex, v, inline);
        }

        /// <summary>Asynchronous twin of <see cref="ReadRichText"/> for inline strings; same contract, same decoding.</summary>
        private static async ValueTask<string> ReadRichTextAsync(XmlReader xr)
        {
            var sb = new StringBuilder();
            var depth = xr.Depth;
            await xr.ReadAsync().ConfigureAwait(false);
            while (!xr.EOF && !(xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth))
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    if (xr.LocalName == "rPh") { await xr.SkipAsync().ConfigureAwait(false); continue; }
                    if (xr.LocalName == "t")
                    {
                        if (xr.IsEmptyElement) await xr.ReadAsync().ConfigureAwait(false);
                        else sb.Append(await xr.ReadElementContentAsStringAsync().ConfigureAwait(false)); // Advances past </t>.
                        continue;
                    }
                }
                await xr.ReadAsync().ConfigureAwait(false);
            }
            return Decode(sb.ToString()); // The reader is now on the container's end element.
        }

        // ------------------------------------------------------------------ objects and tables

        /// <summary>
        /// Maps rows to objects asynchronously, with the rules of <see cref="Read{T}(string)"/>: the first row is
        /// the header, headers match members case-insensitively, a value that cannot be converted throws a
        /// <see cref="FormatException"/> naming the row and column, and blank rows are skipped. A mapper
        /// registered in <see cref="SheetMappers"/> takes precedence over reflection.
        /// </summary>
        /// <example>
        /// <code>
        /// await foreach (var product in reader.ReadAsync&lt;Product&gt;("Products", ct))
        ///     await db.Products.AddAsync(product, ct);
        /// </code>
        /// </example>
        public IAsyncEnumerable<T> ReadAsync<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(string sheetName, CancellationToken cancellationToken = default) where T : new()
        {
            if (SheetMappers.TryGet<T>(out var mapper)) return ReadAsync(sheetName, mapper, cancellationToken);
            return ReadReflectedAsync<T>(FindSheet(sheetName), cancellationToken); // Wrong sheet name fails here, not on the first MoveNextAsync.
        }

        /// <summary>
        /// Maps rows to objects asynchronously through an explicit <see cref="ISheetRowMapper{T}"/>, like
        /// <see cref="Read{T}(string, ISheetRowMapper{T})"/>. No reflection is involved, so this is the overload
        /// for trimmed and Native AOT applications.
        /// </summary>
        public IAsyncEnumerable<T> ReadAsync<T>(string sheetName, ISheetRowMapper<T> mapper, CancellationToken cancellationToken = default)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            if (!mapper.CanRead) throw new NotSupportedException($"The mapper for {typeof(T).Name} cannot create instances (the type has no parameterless constructor), so it can only be used for writing.");
            return ReadMappedAsync(FindSheet(sheetName), mapper, cancellationToken);
        }

        private async IAsyncEnumerable<T> ReadReflectedAsync<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(int sheetIndex, [EnumeratorCancellation] CancellationToken cancellationToken) where T : new()
        {
            Binding[]? bindings = null;
            await foreach (var row in ReadRowsAsync(sheetIndex, cancellationToken).ConfigureAwait(false))
            {
                if (bindings == null)
                {
                    bindings = Bind(typeof(T), row);
                    continue;
                }
                if (MapRow(row, bindings, out T item)) yield return item;
            }
        }

        private async IAsyncEnumerable<T> ReadMappedAsync<T>(int sheetIndex, ISheetRowMapper<T> mapper, [EnumeratorCancellation] CancellationToken cancellationToken)
        {
            int[]? columns = null;
            await foreach (var row in ReadRowsAsync(sheetIndex, cancellationToken).ConfigureAwait(false))
            {
                if (columns == null)
                {
                    columns = BindColumns(mapper, row);
                    continue;
                }
                if (MapRow(row, columns, mapper, out var item)) yield return item;
            }
        }

        /// <summary>
        /// Reads a sheet into a <see cref="DataTable"/> asynchronously, with the column rules of
        /// <see cref="ReadDataTable(string, bool)"/>. The table is built in memory, so the gain over the
        /// synchronous call is only that the sheet bytes are read without blocking a thread.
        /// </summary>
        public async Task<DataTable> ReadDataTableAsync(string sheetName, bool firstRowIsHeader = true, CancellationToken cancellationToken = default)
        {
            var table = new DataTable(sheetName);
            var first = true;
            await foreach (var row in ReadRowsAsync(sheetName, cancellationToken).ConfigureAwait(false))
            {
                if (first && firstRowIsHeader)
                {
                    first = false;
                    AddHeaderColumns(table, row);
                    continue;
                }
                first = false;
                AddDataRow(table, row);
            }
            return table;
        }
    }
}
#endif
