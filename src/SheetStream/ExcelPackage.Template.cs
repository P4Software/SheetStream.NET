using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    public sealed partial class ExcelPackage
    {
        /// <summary>
        /// Opens a template workbook (designed in Excel: logo, styled headers, formulas, charts) and writes a
        /// filled copy to <paramref name="outputPath"/>. Open sheets with <c>Workbook.Worksheets.Open(name)</c>
        /// and stream rows into them; everything you do not touch is copied unchanged.
        /// </summary>
        /// <example>
        /// <code>
        /// using var package = ExcelPackage.OpenTemplate("invoice-template.xlsx", "invoice-42.xlsx");
        /// var sheet = package.Workbook.Worksheets.Open("Lines");
        /// var money = sheet.TemplateStyle("D5");
        /// var row = 5;
        /// foreach (var line in lines) sheet.Row(row++).WriteCell(line.Sku).WriteCell(line.Qty).Skip(1).WriteCell(line.Total, money);
        /// package.Save();
        /// </code>
        /// </example>
        public static ExcelPackage OpenTemplate(string templatePath, string outputPath, PackageOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(templatePath)) throw new ArgumentException("Template path cannot be empty.", nameof(templatePath));
            if (string.Equals(Path.GetFullPath(templatePath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("The output must be a different file from the template, so the template is never overwritten.", nameof(outputPath));
            TemplateSource source;
            using (var input = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                source = TemplateSource.Load(input, (options ?? PackageOptions.Default).SourceLimits);
            return ToFile(Path.GetFullPath(outputPath), options, source);
        }

        /// <summary>Opens a template from a stream and writes the filled copy to <paramref name="output"/>.</summary>
        public static ExcelPackage OpenTemplate(Stream template, Stream output, bool leaveOpen = false, PackageOptions? options = null)
        {
            if (output == null) throw new ArgumentNullException(nameof(output));
            try
            {
                var source = TemplateSource.Load(template, (options ?? PackageOptions.Default).SourceLimits);
                var package = new ExcelPackage(output, leaveOpen: true, options);
                package.Workbook.AttachTemplate(source);
                package._leaveOpen = leaveOpen;
                return package;
            }
            catch
            {
                // No package reaches the caller, so an output stream it handed over is released here.
                if (!leaveOpen) output.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Save path for template mode: seals the filled sheets, writes the merged shared strings and styles,
        /// rewrites the parts that filling makes stale (workbook calc settings, the calculation chain, table
        /// ranges) and copies every other part of the template byte for byte.
        /// </summary>
        private void SaveTemplate(TemplateSource template)
        {
            if (Workbook.DefinedNames.Count > 0 || Workbook.Protection != null)
                throw new NotSupportedException("Defined names and workbook protection come from the template. Add them to the template in Excel.");

            Workbook.Worksheets.SealAll();

            var written = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var tableBands = new Dictionary<string, (int First, int Last)>(StringComparer.OrdinalIgnoreCase);
            using (var archive = template.OpenArchive())
            {
                foreach (var sheet in Workbook.Worksheets)
                {
                    written.Add(sheet.TemplatePart!);
                    if (sheet.WrittenBand is (int, int) band)
                        foreach (var rel in template.ReadRelationships(archive, sheet.TemplatePart!))
                            if (rel.Type.EndsWith("/table", StringComparison.Ordinal)) tableBands[rel.Target] = band;
                }

                if (template.SharedStringsPart != null)
                {
                    using (var r = template.OpenRequiredXml(archive, template.SharedStringsPart))
                    using (var w = OpenPart(template.SharedStringsPart))
                        Workbook.SharedStrings.WriteTo(w, r);
                    written.Add(template.SharedStringsPart);
                }

                using (var r = template.OpenRequiredXml(archive, template.StylesPart!))
                using (var w = OpenPart(template.StylesPart!))
                    Workbook.Styles.WriteMerged(r, w);
                written.Add(template.StylesPart!);

                using (var r = template.OpenRequiredXml(archive, template.WorkbookPart))
                using (var w = OpenPart(template.WorkbookPart))
                    WriteTemplateWorkbook(r, w);
                written.Add(template.WorkbookPart);

                // The calculation chain lists formula cells in evaluation order; after filling it is stale, and a
                // stale chain makes Excel offer a repair. Drop it with its relationship and content type.
                var workbookRels = TemplateSource.RelsPartOf(template.WorkbookPart);
                using (var r = template.OpenRequiredXml(archive, workbookRels))
                using (var w = OpenPart(workbookRels))
                    CopyWithout(r, w, x => x.LocalName == "Relationship" && (x.GetAttribute("Type") ?? "").EndsWith("/calcChain", StringComparison.Ordinal));
                written.Add(workbookRels);

                using (var r = template.OpenRequiredXml(archive, "[Content_Types].xml"))
                using (var w = OpenPart("[Content_Types].xml"))
                    CopyWithout(r, w, x => x.LocalName == "Override" && template.CalcChainPart != null
                        && string.Equals((x.GetAttribute("PartName") ?? "").TrimStart('/'), template.CalcChainPart, StringComparison.OrdinalIgnoreCase));
                written.Add("[Content_Types].xml");
                if (template.CalcChainPart != null) written.Add(template.CalcChainPart);

                foreach (var pair in tableBands)
                {
                    using (var r = template.OpenXml(archive, pair.Key))
                    {
                        if (r == null) continue;
                        using (var w = OpenPart(pair.Key)) CopyTable(r, w, pair.Value);
                    }
                    written.Add(pair.Key);
                }

                // Everything else is copied as it is. The names come from a file this library did not write,
                // so each is checked before it goes into the output, and the copy as a whole is capped: a zip
                // directory can point thousands of names at one compressed block, which would otherwise make
                // a file of a few hundred kilobytes unpack and repack for hours.
                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var remaining = template.Limits.MaxStreamedPartBytes;
                foreach (var entry in archive.Entries)
                {
                    var name = entry.FullName;
                    if (name.EndsWith("/", StringComparison.Ordinal)) continue; // A folder entry carries nothing.
                    if (!TemplateSource.IsWritablePartName(name))
                        throw new InvalidDataException($"The workbook contains a part whose name is not a valid part name ('{(name.Length > 60 ? name.Substring(0, 60) + "..." : name)}'). The file is damaged or not a genuine workbook.");
                    if (!names.Add(name))
                        throw new InvalidDataException($"The workbook contains more than one part named '{name}'. The file is damaged or not a genuine workbook.");
                    if (written.Contains(name)) continue;
                    using (var source = entry.Open()) remaining -= WriteRawPart(name, source, remaining);
                }
            }

            _archive!.Dispose();
            _archive = null;
            _output.Flush();
        }

        /// <summary>
        /// Copies a part as raw bytes (images, printer settings, custom XML, untouched sheets). Nothing is
        /// held, but the copy takes as long as the part unpacks, so it is capped like any other streamed part.
        /// </summary>
        /// <returns>The number of bytes copied.</returns>
        private long WriteRawPart(string entryName, Stream source, long maxBytes)
        {
            if (_openPart != null) throw new InvalidOperationException($"Cannot write part '{entryName}' while another part is still being written.");
            var entry = _archive!.CreateEntry(entryName, Options.CompressionLevel);
            using (var target = entry.Open())
            {
                var buffer = new byte[81920];
                long total = 0;
                int read;
                while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    total += read;
                    if (total > maxBytes)
                        throw new InvalidDataException($"The parts of the workbook that are copied unchanged unpack to more than {Workbook.Template!.Limits.MaxStreamedPartBytes:N0} bytes in all (reached in '{entryName}'). The file is damaged or was built to waste processing time. If it is genuine, raise ReaderOptions.MaxStreamedPartBytes.");
                    target.Write(buffer, 0, read);
                }
                return total;
            }
        }

        /// <summary>
        /// Copies <c>workbook.xml</c>, setting <c>fullCalcOnLoad</c> so template formulas that read the new rows
        /// (totals, lookups) show correct values the moment the file opens.
        /// </summary>
        private static void WriteTemplateWorkbook(XmlReader r, XmlWriter w)
        {
            // Elements that follow calcPr in CT_Workbook; a missing calcPr is inserted before the first of them.
            var afterCalc = new HashSet<string> { "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst" };
            w.WriteStartDocument(true);
            XmlCopy.MoveToRoot(r);
            if (r.IsEmptyElement) throw new InvalidDataException("The workbook part is empty: it lists no sheets. The file is damaged.");
            XmlCopy.CopyStartElement(r, w);
            r.Read();
            var wroteCalc = false;
            while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == 0))
            {
                if (r.NodeType != XmlNodeType.Element || r.Depth != 1) { XmlCopy.CopyNode(r, w); r.Read(); continue; }
                if (r.LocalName == "calcPr")
                {
                    w.WriteStartElement(r.Prefix, r.LocalName, r.NamespaceURI);
                    if (r.MoveToFirstAttribute())
                    {
                        do { if (r.LocalName != "fullCalcOnLoad") w.WriteAttributeString(r.Prefix, r.LocalName, r.NamespaceURI, r.Value); }
                        while (r.MoveToNextAttribute());
                        r.MoveToElement();
                    }
                    w.WriteAttributeString("fullCalcOnLoad", "1");
                    w.WriteEndElement();
                    r.Skip();
                    wroteCalc = true;
                    continue;
                }
                if (!wroteCalc && afterCalc.Contains(r.LocalName))
                {
                    WriteCalcPr(w);
                    wroteCalc = true;
                }
                XmlCopy.CopySubtree(r, w);
            }
            if (!wroteCalc) WriteCalcPr(w);
            w.WriteEndElement();
            w.WriteEndDocument();
        }

        private static void WriteCalcPr(XmlWriter w)
        {
            w.WriteStartElement("calcPr", OpenXmlNamespaces.Main);
            w.WriteAttributeString("fullCalcOnLoad", "1");
            w.WriteEndElement();
        }

        /// <summary>Copies a whole XML part, leaving out the elements <paramref name="drop"/> selects.</summary>
        private static void CopyWithout(XmlReader r, XmlWriter w, Func<XmlReader, bool> drop)
        {
            w.WriteStartDocument(true);
            XmlCopy.MoveToRoot(r);
            // The elements to drop are children. If the root itself matches, the part is not what its name
            // says it is, and copying "nothing" would leave a document without a root element.
            if (drop(r))
                throw new InvalidDataException($"A part of the workbook starts with a '{r.LocalName}' element where its container element belongs. The file is damaged.");
            XmlCopy.CopySubtree(r, w, null, drop);
            w.WriteEndDocument();
        }

        /// <summary>Copies a table part, stretching its range and filter over the rows written below its header.</summary>
        private static void CopyTable(XmlReader r, XmlWriter w, (int First, int Last) band)
        {
            w.WriteStartDocument(true);
            XmlCopy.MoveToRoot(r);
            // A totals row sits at the table's bottom; stretching would run the data over it, so leave such tables alone.
            var hasTotals = (TemplateSource.Int(r.GetAttribute("totalsRowCount")) ?? 0) > 0;
            XmlCopy.CopySubtree(r, w, hasTotals ? null : (element, attribute, value) =>
                element == "table" && attribute == "ref" ? Worksheet.ExtendTemplateRange("table", "ref", value, band)
                : element == "autoFilter" && attribute == "ref" ? Worksheet.ExtendTemplateRange("autoFilter", "ref", value, band)
                : value);
            w.WriteEndDocument();
        }
    }
}
