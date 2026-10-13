namespace SheetStream.Styles
{
    /// <summary>
    /// A registered cell style. Obtain instances from <see cref="StyleSheet.Create(System.Action{StyleOptions})"/>
    /// and pass them to the cell-writing methods. Instances are immutable and cheap to share.
    /// </summary>
    public sealed class CellStyle
    {
        internal CellStyle(int id)
        {
            Id = id;
        }

        /// <summary>
        /// Zero-based index into the workbook's <c>cellXfs</c> collection. Index 0 is the default style.
        /// </summary>
        public int Id { get; }

        /// <summary>True for the workbook's default style, which is never written as an attribute.</summary>
        public bool IsDefault => Id == 0;

        /// <inheritdoc />
        public override string ToString() => "CellStyle #" + Id;
    }

    /// <summary>
    /// A differential format: the partial style (font colour and weight, fill, border, number format) that a
    /// conditional formatting rule lays over a cell when the rule matches. Create one with
    /// <see cref="StyleSheet.CreateDifferential(System.Action{StyleOptions})"/>.
    /// </summary>
    public sealed class DifferentialStyle
    {
        internal DifferentialStyle(int id)
        {
            Id = id;
        }

        /// <summary>Zero-based index into the workbook's <c>dxfs</c> collection.</summary>
        public int Id { get; }

        /// <inheritdoc />
        public override string ToString() => "DifferentialStyle #" + Id;
    }
}
