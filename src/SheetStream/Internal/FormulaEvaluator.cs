using System;
using System.Globalization;

namespace SheetStream.Internal
{
    /// <summary>
    /// Evaluates the small subset of Excel formulas that business reports use for totals, so their result can
    /// be written as the cell's cached value. This is deliberately not a formula engine: it knows numbers,
    /// same-sheet cell and range references, <c>+ - * /</c>, unary minus, a percent suffix, parentheses and
    /// the functions SUM, AVERAGE, COUNT, COUNTA, MIN, MAX, ROUND and ABS. Everything else, and every
    /// reference to a cell that has not been written yet, makes evaluation fail so the formula is written
    /// without a cached value. Failing is always preferred over guessing: a viewer that shows a blank is
    /// better than one that shows a wrong total.
    /// </summary>
    internal static class FormulaEvaluator
    {
        /// <summary>
        /// Tries to compute <paramref name="formula"/> (without the leading <c>=</c>) for the cell being
        /// written at <paramref name="row"/>/<paramref name="column"/>. Only cells already written (rows above,
        /// or cells to the left on the same row) may be referenced.
        /// </summary>
        /// <returns>False when the formula uses anything outside the supported subset, references a cell that is
        /// not yet known, divides by zero or produces a non-finite number.</returns>
        public static bool TryEvaluate(string formula, FormulaValueStore values, int row, int column, string sheetName, out double result)
        {
            var parser = new Parser(formula, values, row, column, sheetName);
            result = parser.Run();
            if (parser.Failed || double.IsNaN(result) || double.IsInfinity(result))
            {
                result = 0;
                return false;
            }
            return true;
        }

        private enum Function { Sum, Average, Count, CountA, Min, Max, Round, Abs }

        /// <summary>
        /// Recursive-descent parser that evaluates as it parses. Unsupported input sets <see cref="Failed"/>
        /// instead of throwing, because a formula per row on a large export would otherwise pay for an
        /// exception per row.
        /// </summary>
        private sealed class Parser
        {
            private readonly string _s;
            private readonly FormulaValueStore _values;
            private readonly int _row;
            private readonly int _column;
            private readonly string _sheetName;
            private int _pos;

            public Parser(string formula, FormulaValueStore values, int row, int column, string sheetName)
            {
                _s = formula;
                _values = values;
                _row = row;
                _column = column;
                _sheetName = sheetName;
            }

            public bool Failed { get; private set; }

            public double Run()
            {
                SkipSpaces();
                var v = ParseExpression();
                SkipSpaces();
                if (_pos != _s.Length) Fail();
                return v;
            }

            private double Fail()
            {
                Failed = true;
                return 0;
            }

            // expression := term (('+' | '-') term)*
            private double ParseExpression()
            {
                var left = ParseTerm();
                while (!Failed)
                {
                    SkipSpaces();
                    if (_pos >= _s.Length) break;
                    var op = _s[_pos];
                    if (op != '+' && op != '-') break;
                    _pos++;
                    var right = ParseTerm();
                    left = op == '+' ? left + right : left - right;
                }
                return left;
            }

            // term := factor (('*' | '/') factor)*
            private double ParseTerm()
            {
                var left = ParseFactor();
                while (!Failed)
                {
                    SkipSpaces();
                    if (_pos >= _s.Length) break;
                    var op = _s[_pos];
                    if (op != '*' && op != '/') break;
                    _pos++;
                    var right = ParseFactor();
                    if (op == '*') left *= right;
                    else if (right == 0) return Fail(); // Excel shows #DIV/0!; an error is not something to cache.
                    else left /= right;
                }
                return left;
            }

            // factor := ('-' | '+') factor | primary '%'*
            private double ParseFactor()
            {
                SkipSpaces();
                if (_pos < _s.Length && (_s[_pos] == '-' || _s[_pos] == '+'))
                {
                    var negative = _s[_pos] == '-';
                    _pos++;
                    var v = ParseFactor();
                    return negative ? -v : v;
                }
                var value = ParsePrimary();
                while (!Failed)
                {
                    SkipSpaces();
                    if (_pos < _s.Length && _s[_pos] == '%')
                    {
                        _pos++;
                        value /= 100;
                    }
                    else break;
                }
                return value;
            }

            // primary := number | '(' expression ')' | function '(' arguments ')' | cell reference
            private double ParsePrimary()
            {
                SkipSpaces();
                if (_pos >= _s.Length) return Fail();
                var c = _s[_pos];
                if ((c >= '0' && c <= '9') || c == '.') return ParseNumber();
                if (c == '(')
                {
                    _pos++;
                    var v = ParseExpression();
                    SkipSpaces();
                    if (_pos >= _s.Length || _s[_pos] != ')') return Fail();
                    _pos++;
                    return v;
                }

                // A reference (possibly sheet-qualified), or a function name followed by '('.
                if (c == '\'' || c == '$' || IsLetter(c) || c == '_')
                {
                    var start = _pos;
                    if (c != '\'' && c != '$')
                    {
                        var name = ReadIdentifier();
                        SkipSpaces();
                        if (_pos < _s.Length && _s[_pos] == '(')
                        {
                            _pos++;
                            return ParseFunctionCall(name);
                        }
                        _pos = start;
                    }
                    if (!TryParseReference(out var range, out var isRange)) return Fail();
                    if (isRange) return Fail(); // A range only makes sense inside an aggregate function.
                    return CellValueForArithmetic(range.FromRow, range.FromColumn);
                }
                return Fail(); // Strings, comparisons, names, '&', '^' and anything else.
            }

            private double ParseNumber()
            {
                var start = _pos;
                while (_pos < _s.Length && ((_s[_pos] >= '0' && _s[_pos] <= '9') || _s[_pos] == '.')) _pos++;
                if (_pos < _s.Length && (_s[_pos] == 'e' || _s[_pos] == 'E'))
                {
                    // Only an exponent when digits follow, otherwise "2E" could swallow the start of a reference like E5.
                    var save = _pos;
                    _pos++;
                    if (_pos < _s.Length && (_s[_pos] == '+' || _s[_pos] == '-')) _pos++;
                    if (_pos < _s.Length && _s[_pos] >= '0' && _s[_pos] <= '9')
                    {
                        while (_pos < _s.Length && _s[_pos] >= '0' && _s[_pos] <= '9') _pos++;
                    }
                    else _pos = save;
                }
                return double.TryParse(_s.Substring(start, _pos - start), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : Fail();
            }

            private string ReadIdentifier()
            {
                var start = _pos;
                while (_pos < _s.Length && (IsLetter(_s[_pos]) || (_s[_pos] >= '0' && _s[_pos] <= '9') || _s[_pos] == '_' || _s[_pos] == '.' || _s[_pos] == '$')) _pos++;
                return _s.Substring(start, _pos - start);
            }

            private double ParseFunctionCall(string name)
            {
                Function fn;
                switch (name.ToUpperInvariant())
                {
                    case "SUM": fn = Function.Sum; break;
                    case "AVERAGE": fn = Function.Average; break;
                    case "COUNT": fn = Function.Count; break;
                    case "COUNTA": fn = Function.CountA; break;
                    case "MIN": fn = Function.Min; break;
                    case "MAX": fn = Function.Max; break;
                    case "ROUND": fn = Function.Round; break;
                    case "ABS": fn = Function.Abs; break;
                    default: return Fail();
                }

                // Aggregate state; the function decides which fields matter.
                double sum = 0, min = double.MaxValue, max = double.MinValue;
                var count = 0;
                var argCount = 0;
                double arg0 = 0, arg1 = 0;

                SkipSpaces();
                if (_pos < _s.Length && _s[_pos] == ')') return Fail(); // Excel rejects empty argument lists too.
                while (!Failed)
                {
                    argCount++;
                    if (TryParseRangeArgument(out var range))
                    {
                        if (fn == Function.Round || fn == Function.Abs)
                        {
                            // Scalar functions take a single cell, read like an arithmetic operand.
                            if (range.RowCount != 1 || range.ColumnCount != 1) return Fail();
                            var scalar = CellValueForArithmetic(range.FromRow, range.FromColumn);
                            if (argCount == 1) arg0 = scalar; else arg1 = scalar;
                        }
                        else if (!Aggregate(range, fn, ref sum, ref min, ref max, ref count)) return Fail();
                    }
                    else
                    {
                        var v = ParseExpression();
                        if (Failed) return 0;
                        // A literal or expression argument is a number for every function (SUM(1, A1) adds 1).
                        sum += v;
                        count++;
                        if (v < min) min = v;
                        if (v > max) max = v;
                        if (argCount == 1) arg0 = v; else arg1 = v;
                    }
                    SkipSpaces();
                    if (_pos >= _s.Length) return Fail();
                    if (_s[_pos] == ',') { _pos++; continue; }
                    if (_s[_pos] == ')') { _pos++; break; }
                    return Fail();
                }
                if (Failed) return 0;

                switch (fn)
                {
                    case Function.Sum: return sum;
                    case Function.Average: return count == 0 ? Fail() : sum / count; // AVERAGE of nothing is #DIV/0!.
                    case Function.Count: return count;
                    case Function.CountA: return count;
                    case Function.Min: return count == 0 ? 0 : min; // Excel returns 0 for MIN/MAX over empty cells.
                    case Function.Max: return count == 0 ? 0 : max;
                    case Function.Abs: return argCount == 1 ? Math.Abs(arg0) : Fail();
                    default: return argCount == 2 ? Round(arg0, arg1) : Fail();
                }
            }

            /// <summary>Folds the cells of a range into the aggregate state. False when a cell is unknown or not yet written.</summary>
            private bool Aggregate(RangeReference range, Function fn, ref double sum, ref double min, ref double max, ref int count)
            {
                if (!IsWritten(range.ToRow, range.ToColumn)) return false;
                for (var r = range.FromRow; r <= range.ToRow; r++)
                {
                    for (var c = range.FromColumn; c <= range.ToColumn; c++)
                    {
                        var state = _values.Get(r, c, out var v);
                        switch (state)
                        {
                            case FormulaCellState.Unknown: return false;
                            case FormulaCellState.Empty: continue;
                            case FormulaCellState.Text:
                                if (fn == Function.CountA) count++;
                                continue;
                            default:
                                count++;
                                sum += v;
                                if (v < min) min = v;
                                if (v > max) max = v;
                                continue;
                        }
                    }
                }
                return true;
            }

            /// <summary>
            /// Parses an argument that is a bare reference (<c>C2:C101</c>, <c>$B$7</c>, <c>'Sheet'!A1:A9</c>) followed
            /// by <c>,</c> or <c>)</c>. Anything else (an expression such as <c>B7*2</c>) is left for the scalar path.
            /// </summary>
            private bool TryParseRangeArgument(out RangeReference range)
            {
                var start = _pos;
                SkipSpaces();
                if (TryParseReference(out range, out _))
                {
                    SkipSpaces();
                    if (_pos < _s.Length && (_s[_pos] == ',' || _s[_pos] == ')')) return true;
                }
                _pos = start; // Not a reference argument: rewind and let the caller parse an expression.
                return false;
            }

            /// <summary>
            /// Parses <c>[sheet!]cell[:cell]</c>. A sheet qualifier must name the current sheet; another sheet's
            /// values are not available to a streaming writer, so that fails.
            /// </summary>
            private bool TryParseReference(out RangeReference range, out bool isRange)
            {
                range = default;
                isRange = false;
                var start = _pos;

                if (_pos < _s.Length && _s[_pos] == '\'')
                {
                    // Quoted sheet name; an embedded apostrophe is doubled.
                    _pos++;
                    var sb = new System.Text.StringBuilder();
                    while (true)
                    {
                        if (_pos >= _s.Length) return false;
                        if (_s[_pos] == '\'')
                        {
                            if (_pos + 1 < _s.Length && _s[_pos + 1] == '\'') { sb.Append('\''); _pos += 2; continue; }
                            _pos++;
                            break;
                        }
                        sb.Append(_s[_pos++]);
                    }
                    if (_pos >= _s.Length || _s[_pos] != '!') return false;
                    _pos++;
                    if (!IsCurrentSheet(sb.ToString())) return false;
                }
                else
                {
                    // Unquoted: "Sheet1!A1". Only when a '!' follows the identifier; otherwise it is the cell itself.
                    var save = _pos;
                    var ident = ReadIdentifier();
                    if (_pos < _s.Length && _s[_pos] == '!')
                    {
                        _pos++;
                        if (!IsCurrentSheet(ident)) return false;
                    }
                    else _pos = save;
                }

                if (!TryParseCell(out var r1, out var c1)) { _pos = start; return false; }
                var r2 = r1;
                var c2 = c1;
                if (_pos < _s.Length && _s[_pos] == ':')
                {
                    _pos++;
                    if (!TryParseCell(out r2, out c2)) { _pos = start; return false; }
                    isRange = true;
                }
                range = new RangeReference(Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
                return true;
            }

            /// <summary>Parses <c>B7</c> or <c>$B$7</c> (case-insensitive) into a row and column.</summary>
            private bool TryParseCell(out int row, out int column)
            {
                row = column = 0;
                if (_pos < _s.Length && _s[_pos] == '$') _pos++;
                var letters = 0;
                while (_pos < _s.Length && IsLetter(_s[_pos]))
                {
                    column = column * 26 + (char.ToUpperInvariant(_s[_pos]) - 'A' + 1);
                    letters++;
                    _pos++;
                }
                if (letters == 0 || letters > 3 || column > CellReference.MaxColumns) return false;
                if (_pos < _s.Length && _s[_pos] == '$') _pos++;
                var digits = 0;
                while (_pos < _s.Length && _s[_pos] >= '0' && _s[_pos] <= '9')
                {
                    row = row * 10 + (_s[_pos] - '0');
                    digits++;
                    _pos++;
                    if (row > CellReference.MaxRows) return false;
                }
                if (digits == 0 || row < 1) return false;
                // "A1B" would be a defined name, not a reference.
                if (_pos < _s.Length && (IsLetter(_s[_pos]) || _s[_pos] == '_' || _s[_pos] == '.')) return false;
                return true;
            }

            /// <summary>Value of one cell used as an arithmetic operand: numbers as-is, empty as 0, text and unknown fail.</summary>
            private double CellValueForArithmetic(int row, int column)
            {
                if (!IsWritten(row, column)) return Fail();
                switch (_values.Get(row, column, out var v))
                {
                    case FormulaCellState.Number: return v;
                    case FormulaCellState.Empty: return 0;
                    default: return Fail(); // Text in arithmetic is #VALUE!; a boolean would be coerced but is not tracked.
                }
            }

            /// <summary>True when the cell precedes the one being written: a row above, or the same row to the left.</summary>
            private bool IsWritten(int row, int column) => row < _row || (row == _row && column < _column);

            private bool IsCurrentSheet(string name) => string.Equals(name, _sheetName, StringComparison.OrdinalIgnoreCase);

            private void SkipSpaces()
            {
                while (_pos < _s.Length && (_s[_pos] == ' ' || _s[_pos] == '\t' || _s[_pos] == '\n' || _s[_pos] == '\r')) _pos++;
            }

            private static bool IsLetter(char c) => (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');

            /// <summary>Excel's ROUND: half away from zero, negative digits round to tens, hundreds...</summary>
            private static double Round(double value, double digits)
            {
                var n = (int)Math.Truncate(digits);
                if (n >= 0) return n > 15 ? value : Math.Round(value, n, MidpointRounding.AwayFromZero);
                var scale = Math.Pow(10, -n);
                return Math.Round(value / scale, MidpointRounding.AwayFromZero) * scale;
            }
        }
    }
}
