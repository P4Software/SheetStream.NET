using System;
using System.IO;
using SheetStream.Internal;

namespace SheetStream
{
    public sealed partial class ExcelPackage
    {
        // A package that writes to a path works in a temporary file beside the target and swaps it in once the
        // save has succeeded, so a failure anywhere (or no save at all) leaves the target as it was.
        private string? _pendingTarget;
        private string? _pendingTemp;

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
        /// to <paramref name="path"/> and replaces the original only when <see cref="Save"/> completes. You must
        /// call <see cref="Save"/>: leaving the <c>using</c> block without it (because your code threw, say)
        /// discards the changes, so a half-finished edit never overwrites the file. If saving throws, the
        /// original is likewise left untouched and the temporary file is removed. The whole workbook is
        /// re-written, see <see cref="Edit(string, string, PackageOptions?)"/>.
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
                source = TemplateSource.Load(input, (options ?? PackageOptions.Default).SourceLimits);

            return ToFile(target, options, source);
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
        /// Called after a successful save of a package that writes to a path: closes the temporary file and
        /// puts it in place of the target. Replacing an existing file is atomic on the same volume, so readers
        /// see either the old file or the new one, never part of each.
        /// </summary>
        private void CompletePendingFile()
        {
            if (_pendingTemp == null) return;
            var temp = _pendingTemp;
            var target = _pendingTarget!;
            _output.Dispose(); // The temporary file must be closed before it can be moved.
            try
            {
                if (!File.Exists(target))
                {
                    File.Move(temp, target);
                }
                else
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
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                // The finished workbook stays in its temporary file, so the caller can clear the obstacle and
                // call Save again. Disposing the package without doing so removes the temporary file.
                throw new IOException($"The workbook was written but could not be put in place at '{target}'. The file may be open in Excel or read-only; whatever was there is unchanged. Close it and call Save again.", ex);
            }
            _pendingTemp = null;
        }

        /// <summary>Removes the temporary file of a package that never completed (a failed or skipped save).</summary>
        private void DiscardPendingFile()
        {
            if (_pendingTemp == null) return;
            TryDelete(_pendingTemp);
            _pendingTemp = null;
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
