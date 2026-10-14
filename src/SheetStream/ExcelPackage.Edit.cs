using System;
using System.IO;
using SheetStream.Internal;

namespace SheetStream
{
    public sealed partial class ExcelPackage
    {
        // In-place editing: the package writes to a temporary file beside the target and swaps it in once the
        // save has succeeded, so a failure anywhere leaves the original file untouched.
        private string? _inPlaceTarget;
        private string? _inPlaceTemp;

        /// <summary>
        /// Opens an existing workbook for editing and writes the edited copy to <paramref name="outputPath"/>.
        /// Open a sheet with <c>Workbook.Worksheets.Edit(name)</c>, read and change cells through
        /// <see cref="Worksheet.Cells"/>, then <see cref="Save"/>. The whole workbook is re-written: edited
        /// sheets from their cells, every other part (other sheets, charts, pictures, styles, names) copied as
        /// it was. Meant for small workbooks; see <see cref="PackageOptions.MaxEditableCells"/>. When both paths
        /// name the same file the workbook is edited in place, see <see cref="Edit(string, PackageOptions?)"/>.
        /// </summary>
        /// <example>
        /// <code>
        /// using var package = ExcelPackage.Edit("budget.xlsx", "budget-2027.xlsx");
        /// var sheet = package.Workbook.Worksheets.Edit("Summary");
        /// var current = sheet.Cells["B4"].Value;
        /// sheet.Cells["B4"].Value = 1250;
        /// sheet.Cells[sheet.Cells.Dimension.LastRow + 1, 1].Value = "Added by code";
        /// package.Save();
        /// </code>
        /// </example>
        public static ExcelPackage Edit(string inputPath, string outputPath, PackageOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("Input path cannot be empty.", nameof(inputPath));
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Output path cannot be empty.", nameof(outputPath));
            if (string.Equals(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
                return Edit(inputPath, options);
            return OpenTemplate(inputPath, outputPath, options);
        }

        /// <summary>
        /// Opens an existing workbook for editing in place. The edited copy is written to a temporary file next
        /// to <paramref name="path"/> and replaces the original only when <see cref="Save"/> (or the disposing
        /// <c>using</c> block) completes; if saving throws, the original is left untouched and the temporary
        /// file is removed. The whole workbook is re-written, see <see cref="Edit(string, string, PackageOptions?)"/>.
        /// </summary>
        /// <example>
        /// <code>
        /// using var package = ExcelPackage.Edit("budget.xlsx");
        /// package.Workbook.Worksheets.Edit("Summary").Cells["B4"].Value = 1250;
        /// package.Save();
        /// </code>
        /// </example>
        public static ExcelPackage Edit(string path, PackageOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Path cannot be empty.", nameof(path));
            var target = Path.GetFullPath(path);
            if (!File.Exists(target)) throw new FileNotFoundException($"The workbook '{path}' does not exist. Editing in place needs an existing file; use ExcelPackage.Create to make a new one.", target);

            // The whole input is read into memory first, so the file handle is released before anything is written.
            TemplateSource source;
            using (var input = new FileStream(target, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                source = TemplateSource.Load(input);

            // The temporary file lives in the same folder so the final swap is a rename on the same volume.
            var temp = Path.Combine(Path.GetDirectoryName(target) ?? "", Path.GetFileName(target) + "." + Guid.NewGuid().ToString("N") + ".tmp");
            var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1 << 16);
            try
            {
                var package = new ExcelPackage(stream, leaveOpen: false, options);
                package.Workbook.AttachTemplate(source);
                package._inPlaceTarget = target;
                package._inPlaceTemp = temp;
                return package;
            }
            catch
            {
                stream.Dispose();
                TryDelete(temp);
                throw;
            }
        }

        /// <summary>
        /// Opens a workbook from a stream for editing and writes the edited copy to <paramref name="output"/>.
        /// See <see cref="Edit(string, string, PackageOptions?)"/>.
        /// </summary>
        public static ExcelPackage Edit(Stream input, Stream output, bool leaveOpen = false, PackageOptions? options = null)
        {
            if (input == null) throw new ArgumentNullException(nameof(input));
            if (output == null) throw new ArgumentNullException(nameof(output));
            return OpenTemplate(input, output, leaveOpen, options);
        }

        /// <summary>
        /// Called after a successful save of an in-place edit: closes the temporary file and swaps it in for
        /// the original. Replace is atomic on the same volume, so readers see either the old or the new file.
        /// </summary>
        private void CompleteInPlaceSave()
        {
            if (_inPlaceTemp == null) return;
            var temp = _inPlaceTemp;
            var target = _inPlaceTarget!;
            _output.Dispose(); // The temporary file must be closed before it can be moved.
            try
            {
                try
                {
                    File.Replace(temp, target, null, ignoreMetadataErrors: true);
                }
                catch (PlatformNotSupportedException)
                {
                    // Some file systems cannot do the three-way replace; fall back to delete and move.
                    File.Delete(target);
                    File.Move(temp, target);
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                TryDelete(temp);
                throw new IOException($"The edited workbook could not replace '{target}'. The file may be open in Excel or read-only; the original is unchanged. Close it and try again.", ex);
            }
            finally
            {
                _inPlaceTemp = null;
            }
        }

        /// <summary>Removes the temporary file of an in-place edit that never completed (a failed or skipped save).</summary>
        private void DiscardInPlaceTemp()
        {
            if (_inPlaceTemp == null) return;
            TryDelete(_inPlaceTemp);
            _inPlaceTemp = null;
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
