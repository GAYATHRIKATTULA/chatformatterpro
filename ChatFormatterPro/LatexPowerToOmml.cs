using System;
using System.Collections.Generic;
using System.Text;

namespace ChatFormatterPro // change namespace if your project uses different one
{
    /// <summary>
    /// Converts a small LaTeX subset to Word OMML equations.
    /// Current supported:
    ///  - powers: x^2, x^{2}, (a+b)^2, (3^4)^2
    ///  - parentheses: ( )
    ///  - braces: { }
    ///  - plain letters/numbers/operators as text (+ - * /)
    /// </summary>
    public static class LatexPowerToOmml
    {
        private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static string ToOmml(string latex)
        {
            if (string.IsNullOrWhiteSpace(latex))
                return WrapOMath(Runs(""));

            // normalize common wrappers
            latex = latex.Trim();

            // remove \( \) or \[ \] if present
            latex = latex.Replace(@"\(", "").Replace(@"\)", "")
                         .Replace(@"\[", "").Replace(@"\]", "").Trim();

            var p = new Parser(latex);
            var expr = p.ParseExpression();

            return WrapOMath(expr);
        }

        // Wrap inner OMML into a full <m:oMath> element
        private static string WrapOMath(string inner)
        {
            return
$@"<m:oMath xmlns:m=""{M}"" xmlns:w=""{W}"">
{inner}
</m:oMath>";
        }

        // Create OMML runs for plain text
        private static string Runs(string text)
        {
            text ??= "";
            return $@"<m:r><m:t>{EscapeXml(text)}</m:t></m:r>";
        }

        private static string EscapeXml(string s)
        {
            return s.Replace("&", "&amp;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;")
                    .Replace("\"", "&quot;")
                    .Replace("'", "&apos;");
        }

        // Build superscript structure: base^(sup)
        private static string Sup(string baseInner, string supInner)
        {
            return
$@"<m:sSup>
  <m:e>
    {baseInner}
  </m:e>
  <m:sup>
    {supInner}
  </m:sup>
</m:sSup>";
        }

        // Put parentheses around an expression (as text runs)
        private static string Parens(string inner)
        {
            return Runs("(") + inner + Runs(")");
        }

        // ---------------- PARSER ----------------
        private sealed class Parser
        {
            private readonly string _s;
            private int _i;

            public Parser(string s) { _s = s; _i = 0; }

            public string ParseExpression()
            {
                // Expression is a sequence of terms (we keep operators as plain runs)
                // and apply exponent parsing when ^ is found.
                var parts = new List<string>();

                while (!End() && Peek() != ')' && Peek() != '}')
                {
                    // skip spaces
                    if (char.IsWhiteSpace(Peek()))
                    {
                        _i++;
                        continue;
                    }

                    var atom = ParseAtomOrGroup();

                    // power handling: atom ^ exponent
                    while (!End() && Peek() == '^')
                    {
                        _i++; // consume ^
                        SkipSpaces();

                        var exponent = ParseExponentGroupOrAtom();
                        atom = Sup(atom, exponent);
                        SkipSpaces();
                    }

                    parts.Add(atom);
                }

                // join parts (sequence)
                return string.Join("", parts);
            }

            private string ParseAtomOrGroup()
            {
                SkipSpaces();
                if (End()) return Runs("");

                char c = Peek();

                // parentheses group
                if (c == '(')
                {
                    _i++; // consume '('
                    var inner = ParseExpression();
                    Expect(')');
                    return Parens(inner);
                }

                // braces group treated similar (without printing braces)
                if (c == '{')
                {
                    _i++; // consume '{'
                    var inner = ParseExpression();
                    Expect('}');
                    return inner;
                }

                // single token: letters/numbers/operators
                return Runs(ParseToken());
            }

            private string ParseExponentGroupOrAtom()
            {
                SkipSpaces();
                if (End()) return Runs("");

                if (Peek() == '{')
                {
                    _i++; // consume '{'
                    var inner = ParseExpression();
                    Expect('}');
                    return inner;
                }
                if (Peek() == '(')
                {
                    _i++; // consume '('
                    var inner = ParseExpression();
                    Expect(')');
                    // exponent with parentheses should keep parentheses visually
                    return Parens(inner);
                }

                // single token exponent (like 2, x)
                return Runs(ParseToken());
            }

            private string ParseToken()
            {
                SkipSpaces();
                if (End()) return "";

                char c = Peek();

                // if operator, return it as single
                if (IsOperator(c))
                {
                    _i++;
                    return c.ToString();
                }

                // read letters/digits sequence (e.g., 3, 12, x, abc)
                var sb = new StringBuilder();
                while (!End())
                {
                    c = Peek();
                    if (char.IsLetterOrDigit(c))
                    {
                        sb.Append(c);
                        _i++;
                        continue;
                    }
                    break;
                }

                // if nothing read (unknown char), consume one char
                if (sb.Length == 0)
                {
                    sb.Append(Peek());
                    _i++;
                }

                return sb.ToString();
            }

            private bool IsOperator(char c)
            {
                return c == '+' || c == '-' || c == '*' || c == '/' || c == '=' || c == ',' || c == '.';
            }

            private void SkipSpaces()
            {
                while (!End() && char.IsWhiteSpace(Peek()))
                    _i++;
            }

            private void Expect(char ch)
            {
                SkipSpaces();
                if (End() || Peek() != ch)
                    return; // be tolerant (don’t crash export)
                _i++;
            }

            private char Peek() => _s[_i];
            private bool End() => _i >= _s.Length;
        }
    }
}
