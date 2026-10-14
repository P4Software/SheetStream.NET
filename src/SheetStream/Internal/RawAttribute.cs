namespace SheetStream.Internal
{
    /// <summary>
    /// An XML attribute copied verbatim from an edited workbook (a row's <c>ht</c>, <c>x14ac:dyDescent</c>, a
    /// cell's <c>cm</c>...). Kept as the four parts <see cref="System.Xml.XmlWriter"/> needs so namespaced
    /// attributes survive the round trip without the library having to know what each one means.
    /// </summary>
    internal readonly struct RawAttribute
    {
        public RawAttribute(string prefix, string localName, string namespaceUri, string value)
        {
            Prefix = prefix;
            LocalName = localName;
            NamespaceUri = namespaceUri;
            Value = value;
        }

        public string Prefix { get; }
        public string LocalName { get; }
        public string NamespaceUri { get; }
        public string Value { get; }

        /// <summary>Writes the attribute, unqualified when it has no namespace.</summary>
        public void WriteTo(System.Xml.XmlWriter w)
        {
            if (NamespaceUri.Length == 0) w.WriteAttributeString(LocalName, Value);
            else w.WriteAttributeString(Prefix, LocalName, NamespaceUri, Value);
        }
    }
}
