using System;
using System.IO.Compression;

namespace SheetStream
{
    /// <summary>Tuning knobs for an <see cref="ExcelPackage"/>. All defaults favour correctness and small files.</summary>
    public sealed class PackageOptions
    {
        /// <summary>
        /// Zip compression applied to every part. <see cref="CompressionLevel.Optimal"/> gives the smallest
        /// file; <see cref="CompressionLevel.Fastest"/> trades roughly 20-30% size for noticeably faster writes
        /// on very large exports.
        /// </summary>
        public CompressionLevel CompressionLevel { get; set; } = CompressionLevel.Optimal;

        /// <summary>
        /// When <c>true</c> (default) string cells are de-duplicated through the shared string table, which
        /// is what Excel itself does and produces the smallest files. When <c>false</c> strings are written
        /// inline in the sheet, so memory stays flat regardless of how many unique strings the export has.
        /// </summary>
        public bool UseSharedStrings { get; set; } = true;

        /// <summary>Size of the write buffer placed in front of each zip entry, in bytes.</summary>
        public int BufferSize { get; set; } = 64 * 1024;

        /// <summary>
        /// Largest number of cells a sheet may hold to be opened with <c>Worksheets.Edit</c>. Editing keeps the
        /// whole sheet in memory, so this guards against loading a data dump by mistake: a sheet over the limit
        /// throws and the message points at streaming or templates instead. Default 200,000 cells.
        /// </summary>
        public int MaxEditableCells { get; set; } = 200_000;

        /// <summary>
        /// Text measurer used by <see cref="Worksheet.AutoFitColumns(double, double)"/> on every sheet of the
        /// package: <c>(text, bold) =&gt; width</c>, where the width is in Excel's column unit, the width of one
        /// digit of the default font (Calibri 11: 7 px at 96 dpi), so <c>"0000000000"</c> measures 10.
        /// SheetStream then adds Excel's cell padding and scales for font size. <c>null</c> (default) uses the
        /// built-in Calibri 11 glyph table, which reproduces Excel's own AutoFit to a few tenths of a unit for
        /// Latin text. Plug in real font metrics (SkiaSharp, System.Drawing, ...) when the workbook uses another
        /// font or non-Latin scripts; the delegate runs once per observed cell, so keep it cheap.
        /// A per-sheet measurer passed to <see cref="Worksheet.AutoFitColumns(Func{string, bool, double}, double, double)"/>
        /// takes precedence.
        /// </summary>
        public Func<string, bool, double>? TextMeasurer { get; set; }

        internal static PackageOptions Default => new PackageOptions();
    }
}
