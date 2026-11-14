using System.IO;
using System.Xml;

namespace SheetStream.Internal
{
    /// <summary>
    /// Wraps an <see cref="XmlReader"/> so that no caller can walk a document nested deeper than
    /// <see cref="MaxDepth"/>. The parser keeps a record for every open element, so a part made of millions of
    /// nested start tags would otherwise grow that list without bound however carefully the text is limited.
    /// </summary>
    /// <remarks>
    /// The check lives in <see cref="Read"/>, and the inherited <see cref="XmlReader.Skip"/>,
    /// <see cref="XmlReader.ReadSubtree"/> and content helpers are all built on <see cref="Read"/>, so they
    /// are covered without each call site having to remember. The template and edit code reads workbook parts
    /// in many places; a wrapper makes the limit a property of the reader instead of a rule for every loop.
    /// </remarks>
    internal sealed class DepthGuardedXmlReader : XmlReader
    {
        /// <summary>Deepest nesting accepted. Real SpreadsheetML nests about a dozen levels.</summary>
        public const int MaxDepth = 256;

        private readonly XmlReader _inner;
        private readonly System.Action? _failed;

        /// <param name="inner">The reader to guard; disposed with this one.</param>
        /// <param name="failed">Called once, the first time reading fails for any reason, before the exception leaves.</param>
        public DepthGuardedXmlReader(XmlReader inner, System.Action? failed = null)
        {
            _inner = inner;
            _failed = failed;
        }

        public override bool Read()
        {
            bool more;
            try
            {
                more = _inner.Read();
            }
            catch
            {
                _failed?.Invoke();
                throw;
            }
            if (_inner.Depth > MaxDepth)
            {
                _failed?.Invoke();
                throw new InvalidDataException($"The workbook XML is nested more than {MaxDepth} levels deep. The file is damaged or not a genuine workbook.");
            }
            return more;
        }

        public override XmlNodeType NodeType => _inner.NodeType;
        public override string LocalName => _inner.LocalName;
        public override string NamespaceURI => _inner.NamespaceURI;
        public override string Prefix => _inner.Prefix;
        public override string Name => _inner.Name;
        /// <summary>
        /// The parser reads a long text node lazily, so asking for the value can pull more of the part through
        /// the limits and fail exactly as <see cref="Read"/> can.
        /// </summary>
        public override string Value
        {
            get
            {
                try
                {
                    return _inner.Value;
                }
                catch
                {
                    _failed?.Invoke();
                    throw;
                }
            }
        }
        public override bool HasValue => _inner.HasValue;
        public override int Depth => _inner.Depth;
        public override string BaseURI => _inner.BaseURI;
        public override bool IsEmptyElement => _inner.IsEmptyElement;
        public override bool IsDefault => _inner.IsDefault;
        public override int AttributeCount => _inner.AttributeCount;
        public override bool EOF => _inner.EOF;
        public override ReadState ReadState => _inner.ReadState;
        public override XmlNameTable NameTable => _inner.NameTable;
        public override XmlReaderSettings? Settings => _inner.Settings;
        public override XmlSpace XmlSpace => _inner.XmlSpace;
        public override string XmlLang => _inner.XmlLang;

        public override string? GetAttribute(string name) => _inner.GetAttribute(name);
        public override string? GetAttribute(string name, string? namespaceURI) => _inner.GetAttribute(name, namespaceURI);
        public override string GetAttribute(int i) => _inner.GetAttribute(i);
        public override bool MoveToAttribute(string name) => _inner.MoveToAttribute(name);
        public override bool MoveToAttribute(string name, string? ns) => _inner.MoveToAttribute(name, ns);
        public override void MoveToAttribute(int i) => _inner.MoveToAttribute(i);
        public override bool MoveToFirstAttribute() => _inner.MoveToFirstAttribute();
        public override bool MoveToNextAttribute() => _inner.MoveToNextAttribute();
        public override bool MoveToElement() => _inner.MoveToElement();
        public override bool ReadAttributeValue() => _inner.ReadAttributeValue();
        public override string? LookupNamespace(string prefix) => _inner.LookupNamespace(prefix);
        public override void ResolveEntity() => _inner.ResolveEntity();

        public override bool CanReadValueChunk => _inner.CanReadValueChunk;
        public override int ReadValueChunk(char[] buffer, int index, int count)
        {
            try
            {
                return _inner.ReadValueChunk(buffer, index, count);
            }
            catch
            {
                _failed?.Invoke();
                throw;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
