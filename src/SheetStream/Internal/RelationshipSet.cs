using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace SheetStream.Internal
{
    /// <summary>
    /// The relationships of one part (a sheet, a drawing). Ids are allocated while the referencing XML is
    /// written, and the matching <c>.rels</c> part is written afterwards, once the source part is closed,
    /// which keeps the archive's one-open-entry rule intact.
    /// </summary>
    internal sealed class RelationshipSet
    {
        private readonly List<(string Id, string Type, string Target, bool External)> _items = new List<(string, string, string, bool)>();

        /// <summary>Number of relationships allocated so far.</summary>
        public int Count => _items.Count;

        /// <summary>Allocates the next <c>rIdN</c> for a relationship and returns it.</summary>
        /// <param name="type">Relationship type URI.</param>
        /// <param name="target">Target relative to the source part's folder, or an absolute URI when external.</param>
        /// <param name="external">True for targets outside the package (web hyperlinks).</param>
        public string Add(string type, string target, bool external = false)
        {
            var id = "rId" + (_items.Count + 1).ToString(CultureInfo.InvariantCulture);
            _items.Add((id, type, target, external));
            return id;
        }

        /// <summary>Writes the complete <c>.rels</c> document.</summary>
        public void WriteTo(XmlWriter w)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("Relationships", OpenXmlNamespaces.PackageRelationships);
            foreach (var item in _items)
            {
                w.WriteStartElement("Relationship");
                w.WriteAttributeString("Id", item.Id);
                w.WriteAttributeString("Type", item.Type);
                w.WriteAttributeString("Target", item.Target);
                if (item.External) w.WriteAttributeString("TargetMode", "External");
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndDocument();
        }
    }
}
