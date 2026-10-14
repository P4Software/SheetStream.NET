using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.Net.Http.Headers;

namespace SheetStream.AspNetCore
{
    /// <summary>
    /// An <see cref="IResult"/> that builds a workbook and sends it as a download. Works from minimal APIs and,
    /// on .NET 7 and later, from MVC controller actions that return <see cref="IResult"/>.
    /// </summary>
    /// <remarks>
    /// The workbook is built into a temporary file (deleted on close) and then copied to the response
    /// asynchronously. That keeps memory flat for large exports, sends an accurate Content-Length, and
    /// avoids the synchronous writes ASP.NET Core forbids on the response body by default.
    /// </remarks>
    public sealed class ExcelResult : IResult
    {
        /// <summary>The media type of .xlsx files.</summary>
        public const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        private readonly Func<Workbook, CancellationToken, Task> _build;

        /// <summary>Creates a result that runs <paramref name="build"/> to fill the workbook.</summary>
        public ExcelResult(string fileName, Func<Workbook, CancellationToken, Task> build, PackageOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("File name cannot be empty.", nameof(fileName));
            FileName = fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ? fileName : fileName + ".xlsx";
            _build = build ?? throw new ArgumentNullException(nameof(build));
            Options = options;
        }

        /// <summary>The download file name, always ending in ".xlsx".</summary>
        public string FileName { get; }

        /// <summary>Optional package settings (compression level, shared strings).</summary>
        public PackageOptions? Options { get; }

        /// <inheritdoc />
        public async Task ExecuteAsync(HttpContext httpContext)
        {
            if (httpContext == null) throw new ArgumentNullException(nameof(httpContext));
            var ct = httpContext.RequestAborted;
            var path = Path.GetTempFileName();
            await using var temp = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 1 << 16,
                FileOptions.DeleteOnClose | FileOptions.Asynchronous);

            using (var package = new ExcelPackage(temp, leaveOpen: true, Options))
            {
                await _build(package.Workbook, ct).ConfigureAwait(false);
                package.Save();
            }

            var response = httpContext.Response;
            response.ContentType = ContentType;
            response.ContentLength = temp.Length;
            // filename* carries non-ASCII names (accents, ñ) correctly; filename is the ASCII fallback.
            var disposition = new ContentDispositionHeaderValue("attachment");
            disposition.SetHttpFileName(FileName);
            response.Headers[HeaderNames.ContentDisposition] = disposition.ToString();

            temp.Position = 0;
            await temp.CopyToAsync(response.Body, 1 << 16, ct).ConfigureAwait(false);
        }
    }

    /// <summary>Factory methods for <see cref="ExcelResult"/>.</summary>
    /// <example>
    /// <code>
    /// app.MapGet("/orders.xlsx", (AppDb db) => ExcelResults.Excel("orders.xlsx", async (wb, ct) =>
    ///     await wb.Worksheets.Add("Orders").LoadFromAsyncEnumerable(db.Orders.AsAsyncEnumerable(), new LoadOptions { TableName = "Orders" }, ct)));
    /// </code>
    /// </example>
    public static class ExcelResults
    {
        /// <summary>Returns a download built synchronously, e.g. from an in-memory list.</summary>
        public static ExcelResult Excel(string fileName, Action<Workbook> build, PackageOptions? options = null)
        {
            if (build == null) throw new ArgumentNullException(nameof(build));
            return new ExcelResult(fileName, (wb, _) => { build(wb); return Task.CompletedTask; }, options);
        }

        /// <summary>Returns a download built asynchronously, e.g. from a database query.</summary>
        public static ExcelResult Excel(string fileName, Func<Workbook, CancellationToken, Task> build, PackageOptions? options = null)
            => new ExcelResult(fileName, build, options);

        /// <summary>Lets callers write <c>Results.Extensions.Excel(...)</c> alongside the built-in results.</summary>
        public static ExcelResult Excel(this IResultExtensions extensions, string fileName, Action<Workbook> build, PackageOptions? options = null)
        {
            if (extensions == null) throw new ArgumentNullException(nameof(extensions));
            return Excel(fileName, build, options);
        }

        /// <summary>Async form of <c>Results.Extensions.Excel(...)</c>.</summary>
        public static ExcelResult Excel(this IResultExtensions extensions, string fileName, Func<Workbook, CancellationToken, Task> build, PackageOptions? options = null)
        {
            if (extensions == null) throw new ArgumentNullException(nameof(extensions));
            return Excel(fileName, build, options);
        }
    }
}
