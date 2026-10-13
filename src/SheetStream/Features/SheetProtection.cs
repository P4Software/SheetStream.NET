using System.Globalization;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>
    /// What users may still do on a protected sheet. Everything is blocked by default except selecting
    /// cells, matching Excel's "Protect Sheet" dialog. Cells styled with <c>Protection.Locked = false</c>
    /// stay editable.
    /// </summary>
    public sealed class SheetProtectionOptions
    {
        /// <summary>Select locked cells. Default true.</summary>
        public bool AllowSelectLockedCells { get; set; } = true;
        /// <summary>Select unlocked cells. Default true.</summary>
        public bool AllowSelectUnlockedCells { get; set; } = true;
        /// <summary>Format cells.</summary>
        public bool AllowFormatCells { get; set; }
        /// <summary>Change column widths and hide columns.</summary>
        public bool AllowFormatColumns { get; set; }
        /// <summary>Change row heights and hide rows.</summary>
        public bool AllowFormatRows { get; set; }
        /// <summary>Insert columns.</summary>
        public bool AllowInsertColumns { get; set; }
        /// <summary>Insert rows.</summary>
        public bool AllowInsertRows { get; set; }
        /// <summary>Insert hyperlinks.</summary>
        public bool AllowInsertHyperlinks { get; set; }
        /// <summary>Delete columns.</summary>
        public bool AllowDeleteColumns { get; set; }
        /// <summary>Delete rows.</summary>
        public bool AllowDeleteRows { get; set; }
        /// <summary>Sort.</summary>
        public bool AllowSort { get; set; }
        /// <summary>Use existing AutoFilter drop-downs.</summary>
        public bool AllowAutoFilter { get; set; }
        /// <summary>Use pivot tables.</summary>
        public bool AllowPivotTables { get; set; }
        /// <summary>Edit drawings, charts and comments.</summary>
        public bool AllowEditObjects { get; set; }
        /// <summary>Edit scenarios.</summary>
        public bool AllowEditScenarios { get; set; }
    }

    /// <summary>Sheet protection state written as <c>sheetProtection</c>.</summary>
    internal sealed class SheetProtection
    {
        private readonly string? _hash;
        private readonly string? _salt;
        private readonly SheetProtectionOptions _options;

        public SheetProtection(string? password, SheetProtectionOptions options)
        {
            _options = options;
            if (!string.IsNullOrEmpty(password)) (_hash, _salt) = PasswordHasher.Hash(password!);
        }

        /// <summary>
        /// Writes the element. In CT_SheetProtection most flags mean "this action is PROTECTED" and default to
        /// true, so an allowed action is written as "0"; the two select flags default to false and are
        /// written as "1" only when selection is blocked.
        /// </summary>
        public void WriteTo(XmlWriter w)
        {
            w.WriteStartElement("sheetProtection");
            if (_hash != null)
            {
                w.WriteAttributeString("algorithmName", "SHA-512");
                w.WriteAttributeString("hashValue", _hash);
                w.WriteAttributeString("saltValue", _salt);
                w.WriteAttributeString("spinCount", PasswordHasher.SpinCount.ToString(CultureInfo.InvariantCulture));
            }
            w.WriteAttributeString("sheet", "1");
            if (!_options.AllowEditObjects) w.WriteAttributeString("objects", "1");
            if (!_options.AllowEditScenarios) w.WriteAttributeString("scenarios", "1");
            Allow(w, "formatCells", _options.AllowFormatCells);
            Allow(w, "formatColumns", _options.AllowFormatColumns);
            Allow(w, "formatRows", _options.AllowFormatRows);
            Allow(w, "insertColumns", _options.AllowInsertColumns);
            Allow(w, "insertRows", _options.AllowInsertRows);
            Allow(w, "insertHyperlinks", _options.AllowInsertHyperlinks);
            Allow(w, "deleteColumns", _options.AllowDeleteColumns);
            Allow(w, "deleteRows", _options.AllowDeleteRows);
            if (!_options.AllowSelectLockedCells) w.WriteAttributeString("selectLockedCells", "1");
            Allow(w, "sort", _options.AllowSort);
            Allow(w, "autoFilter", _options.AllowAutoFilter);
            Allow(w, "pivotTables", _options.AllowPivotTables);
            if (!_options.AllowSelectUnlockedCells) w.WriteAttributeString("selectUnlockedCells", "1");
            w.WriteEndElement();
        }

        private static void Allow(XmlWriter w, string attribute, bool allowed)
        {
            if (allowed) w.WriteAttributeString(attribute, "0");
        }
    }
}
