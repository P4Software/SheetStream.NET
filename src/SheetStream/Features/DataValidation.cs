using System;
using System.Globalization;
using System.Text;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream
{
    /// <summary>Comparison used by numeric, date and text-length validations.</summary>
    public enum ValidationOperator
    {
        /// <summary>Value must lie between the two bounds (inclusive).</summary>
        Between,
        /// <summary>Value must lie outside the two bounds.</summary>
        NotBetween,
        /// <summary>Value must equal the bound.</summary>
        Equal,
        /// <summary>Value must differ from the bound.</summary>
        NotEqual,
        /// <summary>Value must be greater than the bound.</summary>
        GreaterThan,
        /// <summary>Value must be less than the bound.</summary>
        LessThan,
        /// <summary>Value must be greater than or equal to the bound.</summary>
        GreaterThanOrEqual,
        /// <summary>Value must be less than or equal to the bound.</summary>
        LessThanOrEqual,
    }

    /// <summary>How Excel reacts to a value that fails validation.</summary>
    public enum ValidationErrorStyle
    {
        /// <summary>Rejects the value (default).</summary>
        Stop,
        /// <summary>Warns and lets the user decide.</summary>
        Warning,
        /// <summary>Informs and accepts the value.</summary>
        Information,
    }

    /// <summary>
    /// A data validation rule over one range: a drop-down list, number or date bounds, text length, or a
    /// custom formula. Create through <see cref="Worksheet.AddDataValidation(string)"/>, then pick exactly one
    /// rule method and optionally add a prompt and error message.
    /// </summary>
    /// <example>
    /// <code>
    /// sheet.AddDataValidation("C2:C500").List("Open", "Closed", "On hold").WithError("Invalid status", "Pick a value from the list.");
    /// sheet.AddDataValidation("D2:D500").WholeNumber(ValidationOperator.Between, 1, 100);
    /// </code>
    /// </example>
    public sealed class DataValidation
    {
        private string? _type;
        private ValidationOperator _operator = ValidationOperator.Between;
        private string? _formula1;
        private string? _formula2;

        private readonly Action<DataValidation> _register;

        /// <param name="range">Cells the rule applies to.</param>
        /// <param name="register">
        /// Adds the validation to its sheet. Called only once a rule is set, so a rule call that throws (a
        /// comma in a list value, say) leaves nothing half-built behind to fail again at save time.
        /// </param>
        internal DataValidation(RangeReference range, Action<DataValidation> register)
        {
            Range = range;
            _register = register;
        }

        internal RangeReference Range { get; }

        /// <summary>Allow the cell to be left empty. Defaults to true, matching Excel.</summary>
        public bool AllowBlank { get; set; } = true;

        /// <summary>Hide the in-cell drop-down arrow of a list validation.</summary>
        public bool HideDropDown { get; set; }

        /// <summary>Title of the input prompt shown when the cell is selected.</summary>
        public string? PromptTitle { get; set; }

        /// <summary>Text of the input prompt shown when the cell is selected.</summary>
        public string? Prompt { get; set; }

        /// <summary>Title of the error alert.</summary>
        public string? ErrorTitle { get; set; }

        /// <summary>Text of the error alert.</summary>
        public string? Error { get; set; }

        /// <summary>Severity of the error alert.</summary>
        public ValidationErrorStyle ErrorStyle { get; set; }

        /// <summary>
        /// Restricts the cells to a fixed drop-down list. Excel limits an inline list to 255 characters
        /// including separators; for longer lists put the values on a sheet and use <see cref="ListFromRange"/>.
        /// </summary>
        public DataValidation List(params string[] values)
        {
            if (values == null || values.Length == 0) throw new ArgumentException("A list validation needs at least one value.", nameof(values));
            var sb = new StringBuilder("\"");
            for (var i = 0; i < values.Length; i++)
            {
                var v = values[i] ?? throw new ArgumentException("List values cannot be null.", nameof(values));
                if (v.IndexOf(',') >= 0) throw new ArgumentException($"List value '{v}' contains a comma, which Excel uses as the separator. Put such values on a sheet and use ListFromRange.", nameof(values));
                if (i > 0) sb.Append(',');
                sb.Append(v.Replace("\"", "\"\""));
            }
            sb.Append('"');
            if (sb.Length - 2 > 255) throw new ArgumentException("The inline list is longer than Excel's 255-character limit. Put the values on a sheet and use ListFromRange.", nameof(values));
            return Set("list", ValidationOperator.Between, sb.ToString(), null);
        }

        /// <summary>
        /// Restricts the cells to the values found in a range, e.g. <c>"'Lists'!A1:A20"</c> or <c>"H2:H30"</c> on
        /// the same sheet. A defined name such as <c>"=StatusList"</c> is passed through as written.
        /// </summary>
        public DataValidation ListFromRange(string rangeOrName)
        {
            if (string.IsNullOrWhiteSpace(rangeOrName)) throw new ArgumentException("Range cannot be empty.", nameof(rangeOrName));
            var t = rangeOrName.Trim().TrimStart('=');
            string formula;
            try { formula = t.IndexOf('!') >= 0 ? RangeReference.QualifyAbsolute(t, "") : RangeReference.Parse(t).ToAbsolute(); }
            catch (ArgumentException) { formula = t; } // A defined name.
            return Set("list", ValidationOperator.Between, formula, null);
        }

        /// <summary>Restricts the cells to whole numbers.</summary>
        public DataValidation WholeNumber(ValidationOperator op, long value1, long? value2 = null)
            => Set("whole", op, Num(value1), value2.HasValue ? Num(value2.Value) : null);

        /// <summary>Restricts the cells to decimal numbers.</summary>
        public DataValidation Decimal(ValidationOperator op, double value1, double? value2 = null)
            => Set("decimal", op, Num(value1), value2.HasValue ? Num(value2.Value) : null);

        /// <summary>Restricts the cells to dates.</summary>
        public DataValidation Date(ValidationOperator op, DateTime value1, DateTime? value2 = null)
            => Set("date", op, Num(value1.Date.ToOADate()), value2.HasValue ? Num(value2.Value.Date.ToOADate()) : null);

        /// <summary>Restricts the number of characters typed into the cells.</summary>
        public DataValidation TextLength(ValidationOperator op, int value1, int? value2 = null)
            => Set("textLength", op, Num(value1), value2.HasValue ? Num(value2.Value) : null);

        /// <summary>
        /// Accepts a value only when the formula is TRUE, written relative to the range's top-left cell,
        /// e.g. <c>"ISNUMBER(B2)"</c> for a range starting at B2.
        /// </summary>
        public DataValidation Custom(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentException("Formula cannot be empty.", nameof(formula));
            return Set("custom", ValidationOperator.Between, formula.Trim().TrimStart('='), null);
        }

        /// <summary>Sets the error alert shown when a value fails.</summary>
        public DataValidation WithError(string title, string message, ValidationErrorStyle style = ValidationErrorStyle.Stop)
        {
            ErrorTitle = title;
            Error = message;
            ErrorStyle = style;
            return this;
        }

        /// <summary>Sets the input prompt shown when a cell of the range is selected.</summary>
        public DataValidation WithPrompt(string title, string message)
        {
            PromptTitle = title;
            Prompt = message;
            return this;
        }

        private DataValidation Set(string type, ValidationOperator op, string formula1, string? formula2)
        {
            if (_type != null) throw new InvalidOperationException("This validation already has a rule. Call AddDataValidation again for another rule.");
            var needsTwo = op == ValidationOperator.Between || op == ValidationOperator.NotBetween;
            if (type != "list" && type != "custom" && needsTwo && formula2 == null)
                throw new ArgumentException("Between and NotBetween need a second bound.");
            _type = type;
            _operator = op;
            _formula1 = formula1;
            _formula2 = needsTwo ? formula2 : null;
            _register(this);
            return this;
        }

        private static string Num(double v) => v.ToString("R", CultureInfo.InvariantCulture);
        private static string Num(long v) => v.ToString(CultureInfo.InvariantCulture);

        private static string ToXml(ValidationOperator op) => op switch
        {
            ValidationOperator.NotBetween => "notBetween",
            ValidationOperator.Equal => "equal",
            ValidationOperator.NotEqual => "notEqual",
            ValidationOperator.GreaterThan => "greaterThan",
            ValidationOperator.LessThan => "lessThan",
            ValidationOperator.GreaterThanOrEqual => "greaterThanOrEqual",
            ValidationOperator.LessThanOrEqual => "lessThanOrEqual",
            _ => "between",
        };

        /// <summary>Writes one <c>dataValidation</c> element (CT_DataValidation: formula1, formula2).</summary>
        internal void WriteTo(XmlWriter w)
        {
            if (_type == null) throw new InvalidOperationException($"The data validation on {Range} has no rule. Call List, WholeNumber, Decimal, Date, TextLength or Custom.");
            w.WriteStartElement("dataValidation");
            w.WriteAttributeString("type", _type);
            if (_type != "list" && _type != "custom" && _operator != ValidationOperator.Between) w.WriteAttributeString("operator", ToXml(_operator));
            if (ErrorStyle == ValidationErrorStyle.Warning) w.WriteAttributeString("errorStyle", "warning");
            else if (ErrorStyle == ValidationErrorStyle.Information) w.WriteAttributeString("errorStyle", "information");
            if (AllowBlank) w.WriteAttributeString("allowBlank", "1");
            // Excel's attribute is inverted: showDropDown="1" HIDES the arrow.
            if (HideDropDown) w.WriteAttributeString("showDropDown", "1");
            w.WriteAttributeString("showInputMessage", "1");
            w.WriteAttributeString("showErrorMessage", "1");
            if (ErrorTitle != null) w.WriteAttributeString("errorTitle", Limit(ErrorTitle, 32));
            if (Error != null) w.WriteAttributeString("error", Limit(Error, 255));
            if (PromptTitle != null) w.WriteAttributeString("promptTitle", Limit(PromptTitle, 32));
            if (Prompt != null) w.WriteAttributeString("prompt", Limit(Prompt, 255));
            w.WriteAttributeString("sqref", Range.ToString());
            w.WriteElementString("formula1", XmlTextUtil.Sanitize(_formula1!));
            if (_formula2 != null) w.WriteElementString("formula2", _formula2);
            w.WriteEndElement();
        }

        /// <summary>Excel rejects the file when titles exceed 32 or messages exceed 255 characters, so trim.</summary>
        private static string Limit(string text, int max) => XmlTextUtil.Sanitize(text.Length <= max ? text : text.Substring(0, max));
    }
}
