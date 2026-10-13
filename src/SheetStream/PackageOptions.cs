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

        internal static PackageOptions Default => new PackageOptions();
    }
}
