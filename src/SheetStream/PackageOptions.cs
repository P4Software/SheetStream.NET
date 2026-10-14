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

        internal static PackageOptions Default => new PackageOptions();
    }
}
