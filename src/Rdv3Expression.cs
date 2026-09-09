// Arithmetic and string expressions used by the table-operation pipeline.
// C# 5, ASCII source. Parsing and evaluation share this implementation.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

internal abstract class Rdv3Expression
{
    public abstract string Evaluate(string[] row);

    public static bool IsFunctionName(string name)
    {
        return name == "regexExtract" || name == "splitPart" || name == "substring";
    }

    public static Rdv3Expression Compile(string text, string[] columns)
    {
        Parser parser = new Parser(text, columns);
        Rdv3Expression expression = parser.ParseExpression();
        parser.Finish();
        return expression;
    }

    private sealed class Literal : Rdv3Expression
    {
        private readonly string value;
        private readonly bool quoted;
        public Literal(string v, bool isQuoted) { value = v; quoted = isQuoted; }
        public string Value { get { return value; } }
        public bool Quoted { get { return quoted; } }
        public override string Evaluate(string[] row) { return value; }
    }

    private sealed class Field : Rdv3Expression
    {
        private readonly int column;
        public Field(int c) { column = c; }
        public override string Evaluate(string[] row) { return row[column]; }
    }

    private sealed class Unary : Rdv3Expression
    {
        private readonly Rdv3Expression child;
        public Unary(Rdv3Expression value) { child = value; }
        public override string Evaluate(string[] row)
        {
            decimal value = Number(child.Evaluate(row));
            return (-value).ToString("G29", CultureInfo.InvariantCulture);
        }
    }

    private sealed class Binary : Rdv3Expression
    {
        private readonly char operation;
        private readonly Rdv3Expression left;
        private readonly Rdv3Expression right;
        public Binary(char op, Rdv3Expression a, Rdv3Expression b)
        {
            operation = op;
            left = a;
            right = b;
        }
        public override string Evaluate(string[] row)
        {
            string a = left.Evaluate(row);
            string b = right.Evaluate(row);
            decimal an;
            decimal bn;
            bool aNumber = Rdv3Input.TryNumber(a, out an);
            bool bNumber = Rdv3Input.TryNumber(b, out bn);
            if (operation == '+' && (!aNumber || !bNumber)) { return a + b; }
            if (!aNumber || !bNumber) { throw new Rdv3RecordError(Rdv3Text.Format(Rdv3Text.RecordNumber, Rdv3Input.Display(!aNumber ? a : b))); }
            decimal value;
            if (operation == '+') { value = an + bn; }
            else if (operation == '-') { value = an - bn; }
            else if (operation == '*') { value = an * bn; }
            else
            {
                if (bn == 0) { throw new Rdv3RecordError(Rdv3Text.Format(Rdv3Text.RecordDivisionZero, a, b)); }
                value = an / bn;
            }
            return value.ToString("G29", CultureInfo.InvariantCulture);
        }
    }

    private sealed class Function : Rdv3Expression
    {
        private readonly string name;
        private readonly Rdv3Expression source;
        private readonly Regex pattern;
        private readonly string separator;
        private readonly int position;
        private readonly int count;

        public Function(string functionName, Rdv3Expression value, string text, int at, int length)
        {
            name = functionName;
            source = value;
            separator = text;
            position = at;
            count = length;
            if (name == "regexExtract") { pattern = new Regex(text, RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(250)); }
        }

        public override string Evaluate(string[] row)
        {
            string value = source.Evaluate(row);
            if (value.Length == 0) { throw Failure(value); }
            string result;
            if (name == "regexExtract")
            {
                Match match = pattern.Match(value);
                if (!match.Success) { throw Failure(value); }
                result = match.Value;
            }
            else if (name == "splitPart")
            {
                string[] parts = value.Split(new string[] { separator }, StringSplitOptions.None);
                if (position >= parts.Length)
                {
                    throw Failure(value);
                }
                result = parts[position];
            }
            else
            {
                if (position > value.Length || count > value.Length - position)
                {
                    throw Failure(value);
                }
                result = value.Substring(position, count);
            }
            if (result.Length == 0) { throw Failure(value); }
            return result;
        }

        private Rdv3RecordError Failure(string value)
        {
            string condition = name == "regexExtract" ? pattern.ToString()
                : name == "splitPart" ? separator + ", " + position.ToString(CultureInfo.InvariantCulture)
                : position.ToString(CultureInfo.InvariantCulture) + ", " + count.ToString(CultureInfo.InvariantCulture);
            return new Rdv3RecordError(Rdv3Text.Format(Rdv3Text.RecordFunction, name,
                Rdv3Input.Display(value), Rdv3Input.Display(condition)));
        }
    }

    private static decimal Number(string text)
    {
        decimal value;
        if (!Rdv3Input.TryNumber(text, out value))
        {
            throw new Rdv3RecordError(Rdv3Text.Format(Rdv3Text.RecordNumber, Rdv3Input.Display(text)));
        }
        return value;
    }

    private sealed class Parser
    {
        private readonly string text;
        private readonly string[] columns;
        private int position;

        public Parser(string source, string[] names)
        {
            text = source ?? "";
            columns = names;
        }

        public Rdv3Expression ParseExpression()
        {
            Rdv3Expression value = ParseTerm();
            while (true)
            {
                Skip();
                if (!Take('+') && !Take('-')) { return value; }
                char operation = text[position - 1];
                value = new Binary(operation, value, ParseTerm());
            }
        }

        private Rdv3Expression ParseTerm()
        {
            Rdv3Expression value = ParseFactor();
            while (true)
            {
                Skip();
                if (!Take('*') && !Take('/')) { return value; }
                char operation = text[position - 1];
                value = new Binary(operation, value, ParseFactor());
            }
        }

        private Rdv3Expression ParseFactor()
        {
            Skip();
            if (Take('-')) { return new Unary(ParseFactor()); }
            if (Take('('))
            {
                Rdv3Expression value = ParseExpression();
                Skip();
                if (!Take(')')) { throw Error(Rdv3Text.ExpressionParen); }
                return value;
            }
            if (position < text.Length && text[position] == '\'') { return new Literal(ParseString(), true); }
            int start = position;
            while (position < text.Length && !char.IsWhiteSpace(text[position])
                   && "+-*/(),".IndexOf(text[position]) < 0) { position++; }
            if (start == position) { throw Error(Rdv3Text.ExpressionExpected); }
            string token = text.Substring(start, position - start);
            Skip();
            if (Take('(')) { return ParseFunction(token); }
            decimal number;
            if (decimal.TryParse(token, NumberStyles.Number, CultureInfo.InvariantCulture, out number))
            {
                return new Literal(number.ToString("G29", CultureInfo.InvariantCulture), false);
            }
            for (int i = 0; i < columns.Length; i++)
            {
                if (columns[i] == token) { return new Field(i); }
            }
            throw Error(Rdv3Text.Format(Rdv3Text.ExpressionColumn, token));
        }

        private Rdv3Expression ParseFunction(string name)
        {
            if (!IsFunctionName(name)) { throw Error(Rdv3Text.Format(Rdv3Text.ExpressionFunction, name)); }
            List<Rdv3Expression> arguments = new List<Rdv3Expression>();
            Skip();
            if (!Take(')'))
            {
                while (true)
                {
                    arguments.Add(ParseExpression());
                    Skip();
                    if (Take(')')) { break; }
                    if (!Take(',')) { throw Error(Rdv3Text.Format(Rdv3Text.ExpressionArgumentSeparator, name)); }
                }
            }
            return MakeFunction(name, arguments);
        }

        private Rdv3Expression MakeFunction(string name, List<Rdv3Expression> arguments)
        {
            int wanted = (name == "regexExtract") ? 2 : 3;
            if (arguments.Count != wanted)
            {
                throw Error(Rdv3Text.Format(Rdv3Text.ExpressionArguments, name, wanted, arguments.Count));
            }
            if (name == "regexExtract")
            {
                string regex = Quoted(name, "pattern", arguments[1]);
                if (regex.Length == 0) { throw Error(Rdv3Text.ExpressionEmptyPattern); }
                try { return new Function(name, arguments[0], regex, 0, 0); }
                catch (ArgumentException ex)
                {
                    throw Error(Rdv3Text.Format(Rdv3Text.ExpressionBadPattern, regex, ex.Message));
                }
            }
            if (name == "splitPart")
            {
                string separator = Quoted(name, "separator", arguments[1]);
                if (separator.Length == 0) { throw Error(Rdv3Text.ExpressionEmptySeparator); }
                return new Function(name, arguments[0], separator,
                                    Whole(name, "position", arguments[2], false), 0);
            }
            return new Function(name, arguments[0], "",
                                Whole(name, "start", arguments[1], false),
                                Whole(name, "length", arguments[2], true));
        }

        private string Quoted(string functionName, string argumentName, Rdv3Expression expression)
        {
            Literal literal = expression as Literal;
            if (literal == null || !literal.Quoted)
            {
                throw Error(Rdv3Text.Format(Rdv3Text.ExpressionQuotedArgument, functionName, argumentName));
            }
            return literal.Value;
        }

        private int Whole(string functionName, string argumentName, Rdv3Expression expression, bool positive)
        {
            Literal literal = expression as Literal;
            decimal value = 0;
            bool valid = literal != null && !literal.Quoted
                && decimal.TryParse(literal.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out value)
                && value == decimal.Truncate(value) && value >= 0 && value <= int.MaxValue;
            if (valid && (!positive || value > 0)) { return (int)value; }
            throw Error(Rdv3Text.Format(Rdv3Text.ExpressionWholeArgument, functionName, argumentName, positive ? 1 : 0));
        }

        private string ParseString()
        {
            position++;
            StringBuilder sb = new StringBuilder();
            while (position < text.Length)
            {
                char ch = text[position++];
                if (ch != '\'') { sb.Append(ch); continue; }
                if (position < text.Length && text[position] == '\'')
                {
                    sb.Append('\'');
                    position++;
                    continue;
                }
                return sb.ToString();
            }
            throw Error(Rdv3Text.ExpressionQuote);
        }

        public void Finish()
        {
            Skip();
            if (position != text.Length) { throw Error(Rdv3Text.ExpressionUnexpected); }
        }

        private bool Take(char wanted)
        {
            if (position < text.Length && text[position] == wanted) { position++; return true; }
            return false;
        }

        private void Skip()
        {
            while (position < text.Length && char.IsWhiteSpace(text[position])) { position++; }
        }

        private InvalidDataException Error(string message)
        {
            return new InvalidDataException(Rdv3Text.Format(Rdv3Text.ExpressionLocation, text, position + 1, message));
        }
    }
}
