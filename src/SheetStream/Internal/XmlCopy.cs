using System;
using System.Xml;

namespace SheetStream.Internal
{
    /// <summary>
    /// Streams XML from a reader to a writer node by node, so template parts are copied without being
    /// loaded into a DOM. Callers can rewrite attribute values and drop whole elements on the way.
    /// </summary>
    internal static class XmlCopy
    {
        /// <summary>Settings for reading untrusted template parts: no DTDs, no external resolution.</summary>
        public static readonly XmlReaderSettings ReaderSettings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = false,
            CloseInput = true,
        };

        /// <summary>
        /// Rewrites one attribute value. Receives the element's local name, the attribute's local name and the
        /// current value, and returns the value to write, or null to drop the attribute.
        /// </summary>
        public delegate string? AttributeRewriter(string element, string attribute, string value);

        /// <summary>
        /// Writes the start tag of the element the reader is on, copying (and optionally rewriting) its
        /// attributes. Leaves the reader on the start tag. Returns true when the element was empty, in which
        /// case the end tag has already been written.
        /// </summary>
        public static bool CopyStartElement(XmlReader r, XmlWriter w, AttributeRewriter? rewrite = null)
        {
            var element = r.LocalName;
            var empty = r.IsEmptyElement;
            w.WriteStartElement(r.Prefix, r.LocalName, r.NamespaceURI);
            if (r.MoveToFirstAttribute())
            {
                do
                {
                    var value = r.Value;
                    // Namespace declarations are copied verbatim; rewriting applies to ordinary attributes only.
                    if (rewrite != null && r.Prefix != "xmlns" && r.LocalName != "xmlns")
                    {
                        var replaced = rewrite(element, r.LocalName, value);
                        if (replaced == null) continue;
                        value = replaced;
                    }
                    w.WriteAttributeString(r.Prefix, r.LocalName, r.NamespaceURI, value);
                }
                while (r.MoveToNextAttribute());
                r.MoveToElement();
            }
            if (empty) w.WriteEndElement();
            return empty;
        }

        /// <summary>
        /// Copies the current node and its whole subtree, rewriting attributes as it goes, and leaves the reader
        /// on the node after it. Elements for which <paramref name="drop"/> returns true are skipped entirely.
        /// </summary>
        public static void CopySubtree(XmlReader r, XmlWriter w, AttributeRewriter? rewrite = null, Func<XmlReader, bool>? drop = null)
        {
            if (r.NodeType != XmlNodeType.Element)
            {
                CopyNode(r, w);
                r.Read();
                return;
            }
            if (drop != null && drop(r))
            {
                r.Skip();
                return;
            }
            var depth = r.Depth;
            if (CopyStartElement(r, w, rewrite))
            {
                r.Read();
                return;
            }
            r.Read();
            while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
            {
                if (r.NodeType == XmlNodeType.Element) CopySubtree(r, w, rewrite, drop);
                else { CopyNode(r, w); r.Read(); }
            }
            w.WriteFullEndElement();
            r.Read();
        }

        /// <summary>Copies a single non-element node (text, whitespace, comment, processing instruction).</summary>
        public static void CopyNode(XmlReader r, XmlWriter w)
        {
            switch (r.NodeType)
            {
                case XmlNodeType.Text: w.WriteString(r.Value); break;
                case XmlNodeType.CDATA: w.WriteCData(r.Value); break;
                case XmlNodeType.Whitespace:
                case XmlNodeType.SignificantWhitespace: w.WriteWhitespace(r.Value); break;
                case XmlNodeType.Comment: w.WriteComment(r.Value); break;
                case XmlNodeType.ProcessingInstruction: w.WriteProcessingInstruction(r.Name, r.Value); break;
                case XmlNodeType.EndElement: w.WriteFullEndElement(); break;
                // XmlDeclaration and DocumentType are handled by the caller (WriteStartDocument) or refused.
            }
        }

        /// <summary>Moves the reader to the document's root element.</summary>
        public static void MoveToRoot(XmlReader r)
        {
            while (r.Read())
                if (r.NodeType == XmlNodeType.Element) return;
            throw new System.IO.InvalidDataException("The template part has no root element.");
        }
    }
}
