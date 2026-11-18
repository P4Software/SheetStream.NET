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
        /// How text cells are stored: <see cref="SheetStream.StringStorage.Auto"/> (default) shares the values of
        /// columns that repeat and are long or varied (including columns that repeat among many thousands of values, which it
        /// finds with two more looks later in the sheet), and writes every other column (unique text, and short values
        /// from a small set) inside its cells,
        /// <see cref="SheetStream.StringStorage.Shared"/> always uses the shared string table (what SheetStream did
        /// before this setting existed) and <see cref="SheetStream.StringStorage.Inline"/> never does. Every choice
        /// reads back identically in Excel and in <see cref="Reading.ExcelReader"/>; only the file layout, the
        /// speed and the memory differ (some third-party readers show <c>_xHHHH_</c> escapes of inline text, see
        /// <see cref="SheetStream.StringStorage"/>). <see cref="UseSharedStrings"/> set to <c>false</c> overrides this setting
        /// and stores everything inline.
        /// </summary>
        public StringStorage StringStorage { get; set; } = StringStorage.Auto;

        /// <summary>
        /// Set to <c>false</c> to write every text inline in its cell, so memory stays flat however many unique
        /// strings the export has; it is the same as <see cref="StringStorage"/> set to
        /// <see cref="SheetStream.StringStorage.Inline"/> and wins over it. The default (<c>true</c>) leaves the
        /// choice to <see cref="StringStorage"/>.
        /// </summary>
        public bool UseSharedStrings { get; set; } = true;

        /// <summary>
        /// Memory budget of the shared string table, in characters (each entry also counts a small fixed
        /// overhead). While the table is under the budget it behaves as always; once it is full, strings it
        /// has not seen before are written inline in the sheet and strings already in it keep being shared.
        /// Repeated values (status codes, names) therefore still de-duplicate, while an export of millions of
        /// unique strings stops growing the table instead of holding them all until the save. The default of
        /// 16,000,000 keeps the table to a few tens of megabytes. Raise it for the smallest possible file when
        /// memory is no concern; it has no effect when text is stored inline
        /// (<see cref="UseSharedStrings"/> <c>false</c> or <see cref="StringStorage"/> set to
        /// <see cref="SheetStream.StringStorage.Inline"/>).
        /// </summary>
        public long MaxSharedStringChars { get; set; } = 16_000_000;

        /// <summary>
        /// When <c>true</c>, every cell is written with its address (<c>r="B2"</c>), the way Excel itself writes
        /// a sheet. The default is <c>false</c>: a cell that sits in the column right after the previous cell
        /// of its row is written without its address, which the Open XML specification allows (a cell without
        /// one belongs in the next column, and the first cell of a row in column A). That makes the file
        /// smaller and the write faster. We read a sheet with gaps, written with the default, through Excel,
        /// <see cref="Reading.ExcelReader"/>, EPPlus 4.5.3.3, ClosedXML 0.105.1, ExcelDataReader 3.9.0 and NPOI 2.7.6's
        /// XSSF reader (checked 2026-11-16): each puts every value in the right cell. Three readers we checked do not:
        /// MiniExcel 1.46.0 ignores the addresses of a sheet whose first cell has none, so a value after a gap lands in
        /// the wrong column (a sheet with no gaps reads correctly); NPOI 2.7.6's event (SAX) reader hands over a
        /// <c>null</c> address for such a cell; and the Excel ADO / OleDb provider (ACE 12.0 and 16.0) fails with
        /// "System resource exceeded". Your own Open XML SDK code that reads <c>Cell.CellReference</c> gets <c>null</c> for
        /// a cell without an address, as the attribute is absent. We have not checked Microsoft Spreadsheet Compare, which
        /// has been reported to need the address. If a tool you or your users rely on needs the address, set this to
        /// <c>true</c>. The cost is a larger file and a somewhat slower write, in proportion to the number of
        /// cells. It applies to every cell SheetStream writes, whichever way: <c>Row()</c>, <c>Cells[,]</c>,
        /// AutoFit, <c>Worksheets.Edit</c> and the rows filled into a template. Rows of a template that are copied
        /// unchanged keep the form the template had.
        /// </summary>
        public bool WriteCellReferences { get; set; }
        /// <summary>
        /// Size of the write buffer placed in front of the zip entries, in bytes. A package has one such buffer, used
        /// by whichever part is being written; a sheet's rows are composed in it directly. Values under 4096 are
        /// treated as 4096. The default of 64 KB suits nearly everything.
        /// </summary>
        public int BufferSize { get; set; } = 64 * 1024;

        /// <summary>
        /// Largest number of cells a sheet may hold to be opened with <c>Worksheets.Edit</c>. Editing keeps the
        /// whole sheet in memory, so this guards against loading a data dump by mistake: a sheet over the limit
        /// throws and the message points at streaming or templates instead. Default 200,000 cells.
        /// </summary>
        public int MaxEditableCells { get; set; } = 200_000;

        private Reading.ReaderOptions _sourceLimits = new Reading.ReaderOptions();

        /// <summary>
        /// Limits applied to the workbook that <see cref="ExcelPackage.OpenTemplate(string, string, PackageOptions?)"/>
        /// or <see cref="ExcelPackage.Edit(string, PackageOptions?)"/> loads, so a crafted file cannot exhaust
        /// memory: the same limits <see cref="Reading.ExcelReader"/> applies, with the same defaults. A sheet
        /// opened with <c>Worksheets.Edit</c> is held in memory and so counts as a loaded part, capped by
        /// <see cref="Reading.ReaderOptions.MaxLoadedPartBytes"/> as well as by <see cref="MaxEditableCells"/>;
        /// sheets that are filled or copied are streamed and may be any size.
        /// <see cref="Reading.ReaderOptions.MaxRowTextLength"/> is not used here.
        /// </summary>
        public Reading.ReaderOptions SourceLimits
        {
            get => _sourceLimits;
            set => _sourceLimits = value ?? throw new ArgumentNullException(nameof(value));
        }

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
