namespace SheetStream.Styles
{
    /// <summary>Line style of a cell border edge. Values map one-to-one onto ST_BorderStyle.</summary>
    public enum BorderStyle
    {
        /// <summary>No border line.</summary>
        None = 0,
        /// <summary>Hairline (the thinnest Excel can draw).</summary>
        Hair,
        /// <summary>Thin solid line (the usual grid-like border).</summary>
        Thin,
        /// <summary>Medium solid line.</summary>
        Medium,
        /// <summary>Thick solid line.</summary>
        Thick,
        /// <summary>Dashed line.</summary>
        Dashed,
        /// <summary>Dotted line.</summary>
        Dotted,
        /// <summary>Double line.</summary>
        Double,
    }

    /// <summary>Horizontal text alignment inside a cell. Values map onto ST_HorizontalAlignment.</summary>
    public enum HorizontalAlignment
    {
        /// <summary>Excel default: text left, numbers right.</summary>
        General = 0,
        /// <summary>Left aligned.</summary>
        Left,
        /// <summary>Centred.</summary>
        Center,
        /// <summary>Right aligned.</summary>
        Right,
        /// <summary>Repeats the content to fill the cell width.</summary>
        Fill,
        /// <summary>Justified across the cell width.</summary>
        Justify,
    }

    /// <summary>Vertical text alignment inside a cell. Values map onto ST_VerticalAlignment.</summary>
    public enum VerticalAlignment
    {
        /// <summary>Excel default.</summary>
        Bottom = 0,
        /// <summary>Top aligned.</summary>
        Top,
        /// <summary>Vertically centred.</summary>
        Center,
        /// <summary>Justified vertically.</summary>
        Justify,
    }
}
