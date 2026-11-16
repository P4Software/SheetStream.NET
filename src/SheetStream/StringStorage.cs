namespace SheetStream
{
    /// <summary>
    /// How text cells are stored in the file. All three choices produce workbooks that Excel and every other
    /// reader (including <see cref="Reading.ExcelReader"/>) read identically: the cell shows the same text. They
    /// differ only in where the text sits inside the .xlsx and therefore in speed, memory and file size.
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
        /// 512 cells of the column. A shared column that stops repeating is written inline from then on. Memory stays
        /// bounded: the table is only fed by the cells that were sampled and by the columns that repeat, and never
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
