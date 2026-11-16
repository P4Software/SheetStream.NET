namespace SheetStream
{
    /// <summary>
    /// How text cells are stored in the file. All three choices produce workbooks that Excel and
    /// <see cref="Reading.ExcelReader"/> read identically: the cell shows the same text. They differ only in where the
    /// text sits inside the .xlsx and therefore in speed, memory and file size. Some third-party readers do not treat the
    /// two layouts alike: openpyxl (and pandas through it) shows the <c>_xHHHH_</c> escapes the file uses for a control
    /// character, and for a literal <c>_xHHHH_</c> in inline text, as characters of the text.
    /// </summary>
    /// <remarks>
    /// A workbook can store a text cell in one of two ways. <b>Shared</b> keeps each distinct text once in a
    /// workbook-wide list (the shared string table) and the cell holds only its number, which is what Excel itself
    /// does. <b>Inline</b> writes the text inside the cell. Shared wins when long or varied values repeat (customer
    /// names, descriptions, codes): the file can be much smaller, and Excel opens a sheet full of repeated texts
    /// faster. Inline wins when values hardly ever repeat (ids, e-mail addresses, free text), and for short values
    /// from a small set (status, country), which the zip compressor folds away: there is no list to build and keep in
    /// memory, and the write is faster.
    /// </remarks>
    public enum StringStorage
    {
        /// <summary>
        /// The default. SheetStream decides column by column while it writes: a column whose values repeat and are
        /// worth sharing keeps using the shared string table; every other column (ids, e-mail addresses, free text,
        /// short values from a small set) is written inline, with no lookup. The decision is made from the first 128 to
        /// 512 cells of the column. A column that looks unique then gets two more looks, of 512 cells each, at its 512th and
        /// 8,192nd cell, because it may repeat among thousands of values that are too far apart to show in its first cells (and
        /// too far apart for the zip compressor to fold away): if the looks estimate that its different values do not fit the
        /// compressor's 32 KB window, it is shared from there on. A shared column that stops repeating is written inline from then
        /// on. Memory stays bounded: the table is only fed by the cells that were sampled and by the columns that repeat, and never
        /// grows beyond <see cref="PackageOptions.MaxSharedStringChars"/>. A sheet of fewer than 128 rows is stored exactly
        /// as <see cref="Shared"/> does.
        /// </summary>
        Auto = 0,

        /// <summary>
        /// Every text goes through the shared string table, whatever it looks like: the smallest file, at the
        /// price of holding every distinct text in memory until the package is saved (up to
        /// <see cref="PackageOptions.MaxSharedStringChars"/>; beyond it new texts are written inline). This is
        /// how SheetStream stored text before <see cref="Auto"/> existed.
        /// </summary>
        Shared = 1,

        /// <summary>
        /// Every text is written inside its cell and no shared string table is built (the file still carries an
        /// empty one). Memory stays flat and the write is the fastest, but a value that repeats is stored every
        /// time it appears. The same as setting <see cref="PackageOptions.UseSharedStrings"/> to <c>false</c>.
        /// </summary>
        Inline = 2,
    }
}
