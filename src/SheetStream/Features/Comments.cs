using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>A cell comment (Excel "note").</summary>
    internal sealed class CellComment
    {
        public CellComment(int row, int column, string text, string author)
        {
            Row = row;
            Column = column;
            Text = text;
            Author = author;
        }

        public int Row { get; }
        public int Column { get; }
        public string Text { get; }
        public string Author { get; }
    }

    /// <summary>
    /// Writes a sheet's comments part and the VML drawing Excel still requires to position the note boxes.
    /// VML is legacy markup, so this emits the minimal shape Excel itself writes for a default note.
    /// </summary>
    internal static class CommentsWriter
    {
        public static void Write(ExcelPackage package, int commentsNumber, int sheetIndex, IReadOnlyList<CellComment> comments, out string commentsTarget, out string vmlTarget)
        {
            var n = commentsNumber.ToString(CultureInfo.InvariantCulture);
            var commentsPart = "xl/comments" + n + ".xml";
            var vmlPart = "xl/drawings/vmlDrawing" + n + ".vml";
            commentsTarget = "../comments" + n + ".xml";
            vmlTarget = "../drawings/vmlDrawing" + n + ".vml";

            var authors = new List<string>();
            var authorIndex = new Dictionary<string, int>();
            foreach (var c in comments)
            {
                if (authorIndex.ContainsKey(c.Author)) continue;
                authorIndex[c.Author] = authors.Count;
                authors.Add(c.Author);
            }

            using (var w = package.OpenPart(commentsPart))
            {
                w.WriteStartDocument(true);
                w.WriteStartElement("comments", OpenXmlNamespaces.Main);
                w.WriteStartElement("authors");
                foreach (var a in authors) w.WriteElementString("author", XmlTextUtil.Sanitize(a));
                w.WriteEndElement();
                w.WriteStartElement("commentList");
                foreach (var c in comments)
                {
                    w.WriteStartElement("comment");
                    w.WriteAttributeString("ref", CellReference.ToA1(c.Row, c.Column));
                    w.WriteAttributeString("authorId", authorIndex[c.Author].ToString(CultureInfo.InvariantCulture));
                    w.WriteStartElement("text");
                    w.WriteStartElement("t");
                    w.WriteAttributeString("xml", "space", null, "preserve");
                    w.WriteString(XmlTextUtil.Sanitize(c.Text));
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                }
                w.WriteEndElement();
                w.WriteEndElement();
                w.WriteEndDocument();
            }
            package.RegisterPart(commentsPart, OpenXmlNamespaces.CtComments);

            // VML is not XML-namespace clean in Excel's eyes; write it as text in the exact shape Excel emits.
            var vml = new StringBuilder();
            vml.Append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
            // Shape ids are allocated in blocks of 1024 per drawing; idmap names this sheet's block.
            vml.Append("<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"").Append(sheetIndex.ToString(CultureInfo.InvariantCulture)).Append("\"/></o:shapelayout>");
            vml.Append("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">");
            vml.Append("<v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>");
            var shapeId = sheetIndex * 1024 + 1;
            foreach (var c in comments)
            {
                var r0 = c.Row - 1;
                var c0 = c.Column - 1;
                var top = r0 > 0 ? r0 - 1 : 0;
                vml.Append("<v:shape id=\"_x0000_s").Append(shapeId++.ToString(CultureInfo.InvariantCulture)).Append("\" type=\"#_x0000_t202\" ");
                vml.Append("style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden\" ");
                vml.Append("fillcolor=\"#ffffe1\" o:insetmode=\"auto\">");
                vml.Append("<v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/>");
                vml.Append("<v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox>");
                vml.Append("<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/>");
                vml.Append("<x:Anchor>").Append(c0 + 1).Append(", 15, ").Append(top).Append(", 10, ").Append(c0 + 3).Append(", 15, ").Append(top + 4).Append(", 4</x:Anchor>");
                vml.Append("<x:AutoFill>False</x:AutoFill><x:Row>").Append(r0).Append("</x:Row><x:Column>").Append(c0).Append("</x:Column></x:ClientData>");
                vml.Append("</v:shape>");
            }
            vml.Append("</xml>");
            package.WriteBinaryPart(vmlPart, Encoding.UTF8.GetBytes(vml.ToString()));
            package.RegisterDefault("vml", OpenXmlNamespaces.CtVmlDrawing);
        }
    }
}
