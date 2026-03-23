// MathLatexHelper.cs
// PASTE ORDER: Replace this file FIRST, before MainWindow.xaml.cs.
//
// This file has TWO public responsibilities:
//
//   1. NormalizeInformalMath(string fullEditorText)
//      Called by MainWindow on paste / "Render Math" button.
//      Converts informal school math notation → clean LaTeX text.
//      The result goes back into the editor so the user can see it.
//      Works line-by-line on the full editor content.
//
//   2. NormalizeForParsing(string latexFragment)
//      Called by DocxExporter on individual math strings.
//      Braces ungrouped exponents, converts simple fractions.
//      Input is already LaTeX (output of NormalizeInformalMath, or
//      LaTeX typed directly by the user).
//
//   ConvertPowersToLatex(string) is kept as a legacy alias for NormalizeForParsing.
//
// WHAT THIS FILE MUST NEVER DO:
//   - Convert ^ or \frac to Unicode characters (², ³, ½ …)
//   - Be called more than once on the same string
//   - Modify text that is already valid LaTeX

using System;
using System.Text;
using System.Text.RegularExpressions;

namespace ChatFormatterPro
{
    public static class MathLatexHelper
    {
        // =====================================================================
        // PUBLIC API — STAGE 1: Editor normalization (informal → LaTeX)
        // =====================================================================

        /// <summary>
        /// Normalizes the full text of the editor so that informal math
        /// notation is converted to clean LaTeX that the DOCX exporter
        /// can parse correctly.
        ///
        /// Call this:
        ///   - When the user pastes content (PreviewTextInput / DataObject.Pasting)
        ///   - When the user clicks "Render Math in LaTeX"
        ///
        /// Processes line by line. Plain-text lines with no math are untouched.
        /// Already-valid LaTeX is detected and skipped so it is never double-processed.
        ///
        /// Example transforms (line level):
        ///   x² + y² = z²           →  x^{2} + y^{2} = z^{2}
        ///   x^2 + y^2 = z^2        →  x^{2} + y^{2} = z^{2}
        ///   (a+b)^2                 →  (a+b)^{2}
        ///   √50                     →  \sqrt{50}
        ///   sqrt(50)                →  \sqrt{50}
        ///   7/2 + 5/3               →  \frac{7}{2} + \frac{5}{3}
        ///   (x+1)/(x-1) = 2         →  (x+1)/(x-1) = 2    [complex — left for parser]
        ///   x₁ + x₂ = 10           →  x_{1} + x_{2} = 10
        ///   x != 5                  →  x \neq 5
        ///   x <= 9                  →  x \le 9
        ///   a >= b                  →  a \ge b
        ///   a ± b                   →  a \pm b
        ///   3 × 4 = 12              →  3 \times 4 = 12
        ///   20 ÷ 5 = 4              →  20 \div 5 = 4
        ///   πr²                     →  \pi r^{2}
        ///   \frac{7}{2}             →  \frac{7}{2}   (already LaTeX — untouched)
        ///   \sqrt{50}               →  \sqrt{50}     (already LaTeX — untouched)
        /// </summary>
        public static string NormalizeInformalMath(string fullText)
        {
            if (string.IsNullOrEmpty(fullText))
                return fullText ?? string.Empty;

            var lines = fullText.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            var sb = new StringBuilder(fullText.Length);

            for (int i = 0; i < lines.Length; i++)
            {
                sb.Append(NormalizeInformalMathLine(lines[i]));
                if (i < lines.Length - 1)
                    sb.Append('\n');
            }

            return sb.ToString();
        }

        // =====================================================================
        // PUBLIC API — STAGE 2: Pre-parse normalization (LaTeX → braced LaTeX)
        // =====================================================================

        /// <summary>
        /// Prepares a LaTeX fragment for the OMML parser.
        /// Input should already be LaTeX (output of NormalizeInformalMath,
        /// or LaTeX typed directly).
        ///
        /// Steps:
        ///   1. Strip outer math wrappers \( \) \[ \] $$ $
        ///   2. Brace ungrouped exponents:  a^2 → a^{2},  a^-2 → a^{-2}
        ///   3. Convert simple a/b fractions → \frac{a}{b}  (if no \frac present)
        /// </summary>
        public static string NormalizeForParsing(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return input ?? string.Empty;

            input = StripMathWrappers(input);
            input = BraceExponents(input);

            if (!input.Contains(@"\frac"))
                input = BraceSlashFractions(input);

            return input.Trim();
        }

        /// <summary>Legacy alias — calls NormalizeForParsing.</summary>
        public static string ConvertPowersToLatex(string input)
            => NormalizeForParsing(input);

        // =====================================================================
        // PRIVATE — line-level informal math normalizer
        // =====================================================================

        /// <summary>
        /// Applies all informal-math-to-LaTeX conversions to a single line.
        /// Order matters: each step is designed not to conflict with the next.
        /// Already-valid LaTeX commands are detected and skipped at each step.
        /// </summary>
        private static string NormalizeInformalMathLine(string line)
        {
            if (string.IsNullOrEmpty(line))
                return line;

            // ── Step 1: Protect existing LaTeX ──────────────────────────────
            // If the line already contains \frac, \sqrt, \pi, etc., it is
            // already LaTeX. We still apply the safe symbol conversions
            // (operators, comparison signs) but skip the structural ones
            // (Unicode superscripts, subscripts, sqrt spelling) to avoid
            // corrupting valid LaTeX.
            bool hasLatexCommands = ContainsLatexCommands(line);

            // ── Step 2: Unicode superscripts (e.g. x² → x^{2}) ─────────────
            // Only on lines without existing LaTeX structure, because ² inside
            // a LaTeX expression is already meaningful or is a formatting issue.
            if (!hasLatexCommands)
                line = ConvertUnicodeSuperscripts(line);

            // ── Step 3: Unicode subscripts (e.g. x₁ → x_{1}) ───────────────
            if (!hasLatexCommands)
                line = ConvertUnicodeSubscripts(line);

            // ── Step 4: π → \pi ─────────────────────────────────────────────
            // Safe on all lines. We only replace standalone π, not πr^{2}
            // (the r^{2} will be handled by step 5/6).
            // We skip this if \pi is already present.
            if (!line.Contains(@"\pi"))
                line = line.Replace("π", @"\pi ");

            // ── Step 5: √N → \sqrt{N}  and  √(expr) → \sqrt{expr} ───────────
            if (!hasLatexCommands)
            {
                // √ followed by digits
                line = Regex.Replace(line, @"√(\d+)", @"\sqrt{$1}");
                // √ followed by a letter
                line = Regex.Replace(line, @"√([A-Za-z])", @"\sqrt{$1}");
                // √(expr) — anything in parens
                line = Regex.Replace(line, @"√\(([^)]+)\)", @"\sqrt{$1}");
                // bare √ with nothing after (edge case)
                line = line.Replace("√", @"\sqrt ");
            }

            // ── Step 6: sqrt(expr) → \sqrt{expr} ────────────────────────────
            // Covers typed "sqrt(50)" — only when not already \sqrt
            if (!line.Contains(@"\sqrt"))
                line = Regex.Replace(line, @"\bsqrt\(([^)]+)\)", @"\sqrt{$1}");

            // ── Step 7: Caret exponents — brace them ────────────────────────
            // a^2 → a^{2},  (a+b)^2 → (a+b)^{2},  a^-2 → a^{-2}
            // We call BraceExponents here (same logic as NormalizeForParsing)
            // so the editor text shows braced exponents immediately.
            line = BraceExponents(line);

            // ── Step 8: Operator symbols → LaTeX commands ───────────────────
            // These are safe on all lines including those with LaTeX.
            // We convert Unicode math operators that the OMML parser needs
            // to see as LaTeX commands.

            // Multiply / divide
            if (!line.Contains(@"\times"))
                line = line.Replace("×", @"\times ");
            if (!line.Contains(@"\div"))
                line = line.Replace("÷", @"\div ");

            // Plus-minus
            if (!line.Contains(@"\pm"))
                line = line.Replace("±", @"\pm ");

            // ── Step 9: Comparison operators → LaTeX ────────────────────────
            // Order: != before anything else (two-char sequence).
            // Use word-boundary-aware replacement to avoid changing "!=" inside
            // code comments, but in math context it is safe.
            line = ReplaceComparisonOps(line);

            // ── Step 10: Simple slash fractions → \frac{}{} ─────────────────
            // Only when no \frac already present on the line,
            // and only for simple numeric or single-letter fractions.
            // Complex fractions like (x+1)/(x-1) are left for the OMML parser
            // which handles / natively.
            if (!line.Contains(@"\frac"))
                line = BraceSlashFractions(line);

            return line;
        }

        // ── Step 2 helper: Unicode superscripts → caret notation ─────────────

        private static readonly (char unicode, string caret)[] SuperscriptMap =
        {
            ('⁰', "^{0}"),  ('¹', "^{1}"),  ('²', "^{2}"),  ('³', "^{3}"),
            ('⁴', "^{4}"),  ('⁵', "^{5}"),  ('⁶', "^{6}"),  ('⁷', "^{7}"),
            ('⁸', "^{8}"),  ('⁹', "^{9}"),
            // negative and positive superscript signs
            ('⁻', "^{-"),   ('⁺', "^{+"),
        };

        /// <summary>
        /// Converts Unicode superscript characters to caret notation.
        ///
        /// Handles runs like x²³ → x^{23} and x⁻² → x^{-2}.
        ///
        /// Algorithm: scan character by character. When a superscript character
        /// is encountered, emit "^{" then collect the whole run, then close "}".
        /// A leading ⁻ or ⁺ is included inside the braces.
        /// </summary>
        private static string ConvertUnicodeSuperscripts(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;

            // Quick exit if no superscript characters present
            bool hasSup = false;
            foreach (char c in s)
            {
                if (IsSuperscriptChar(c)) { hasSup = true; break; }
            }
            if (!hasSup) return s;

            var sb = new StringBuilder(s.Length + 8);
            int i = 0;

            while (i < s.Length)
            {
                char c = s[i];

                if (!IsSuperscriptChar(c))
                {
                    sb.Append(c);
                    i++;
                    continue;
                }

                // Start of a superscript run
                sb.Append("^{");

                // Optional sign (⁻ or ⁺) — comes first
                if (c == '⁻' || c == '⁺')
                {
                    sb.Append(c == '⁻' ? '-' : '+');
                    i++;
                    if (i >= s.Length) { sb.Append('}'); break; }
                    c = s[i];
                }

                // Digit superscripts (collect whole run)
                while (i < s.Length && IsDigitSuperscriptChar(s[i]))
                {
                    sb.Append(SuperscriptDigitToChar(s[i]));
                    i++;
                }

                sb.Append('}');
            }

            return sb.ToString();
        }

        private static bool IsSuperscriptChar(char c)
            => c == '⁰' || c == '¹' || c == '²' || c == '³' || c == '⁴' ||
               c == '⁵' || c == '⁶' || c == '⁷' || c == '⁸' || c == '⁹' ||
               c == '⁻' || c == '⁺';

        private static bool IsDigitSuperscriptChar(char c)
            => c == '⁰' || c == '¹' || c == '²' || c == '³' || c == '⁴' ||
               c == '⁵' || c == '⁶' || c == '⁷' || c == '⁸' || c == '⁹';

        private static char SuperscriptDigitToChar(char c) => c switch
        {
            '⁰' => '0',
            '¹' => '1',
            '²' => '2',
            '³' => '3',
            '⁴' => '4',
            '⁵' => '5',
            '⁶' => '6',
            '⁷' => '7',
            '⁸' => '8',
            '⁹' => '9',
            _ => c
        };

        // ── Step 3 helper: Unicode subscripts → _{} notation ─────────────────

        private static readonly (char unicode, char digit)[] SubscriptMap =
        {
            ('₀','0'),('₁','1'),('₂','2'),('₃','3'),('₄','4'),
            ('₅','5'),('₆','6'),('₇','7'),('₈','8'),('₉','9'),
        };

        /// <summary>
        /// Converts Unicode subscript digits to _{N} notation.
        /// x₁ → x_{1},  x₁₂ → x_{12}
        /// Collects consecutive subscript digits into one _{...} group.
        /// </summary>
        private static string ConvertUnicodeSubscripts(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;

            bool hasSub = false;
            foreach (char c in s)
            {
                if (IsSubscriptDigit(c)) { hasSub = true; break; }
            }
            if (!hasSub) return s;

            var sb = new StringBuilder(s.Length + 8);
            int i = 0;

            while (i < s.Length)
            {
                char c = s[i];

                if (!IsSubscriptDigit(c))
                {
                    sb.Append(c);
                    i++;
                    continue;
                }

                sb.Append("_{");
                while (i < s.Length && IsSubscriptDigit(s[i]))
                {
                    sb.Append(SubscriptDigitToChar(s[i]));
                    i++;
                }
                sb.Append('}');
            }

            return sb.ToString();
        }

        private static bool IsSubscriptDigit(char c)
            => c == '₀' || c == '₁' || c == '₂' || c == '₃' || c == '₄' ||
               c == '₅' || c == '₆' || c == '₇' || c == '₈' || c == '₉';

        private static char SubscriptDigitToChar(char c) => c switch
        {
            '₀' => '0',
            '₁' => '1',
            '₂' => '2',
            '₃' => '3',
            '₄' => '4',
            '₅' => '5',
            '₆' => '6',
            '₇' => '7',
            '₈' => '8',
            '₉' => '9',
            _ => c
        };

        // ── Step 9 helper: Comparison operators ──────────────────────────────

        /// <summary>
        /// Converts !=, &lt;=, &gt;= to their LaTeX equivalents.
        ///
        /// Important: must check for already-converted forms first.
        /// Must not touch >= inside a number like ">=5" written as prose
        /// — in a math document this is almost always a comparison, so we
        /// convert it.  In a code comment context this would be wrong, but
        /// this app processes math content, not code.
        ///
        /// Order: != first (two chars), then <=, then >=.
        /// </summary>
        private static string ReplaceComparisonOps(string s)
        {
            // Already converted — skip
            if (s.Contains(@"\neq") && s.Contains(@"\le") && s.Contains(@"\ge"))
                return s;

            // != → \neq  (check before individual < > processing)
            if (!s.Contains(@"\neq"))
                s = s.Replace("!=", @" \neq ");

            // <= → \le
            if (!s.Contains(@"\le"))
                s = s.Replace("<=", @" \le ");

            // >= → \ge
            if (!s.Contains(@"\ge"))
                s = s.Replace(">=", @" \ge ");

            return s;
        }

        // ── Guard: detect existing LaTeX commands ────────────────────────────

        /// <summary>
        /// Returns true if the line already contains LaTeX structural commands.
        /// Used to skip destructive conversions on lines that are already LaTeX.
        /// </summary>
        private static bool ContainsLatexCommands(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            // Any backslash-command sequence indicates LaTeX
            return s.Contains('\\')
                && Regex.IsMatch(s, @"\\[a-zA-Z]+");
        }

        // =====================================================================
        // PRIVATE — BraceExponents (used in both Stage 1 and Stage 2)
        // =====================================================================

        /// <summary>
        /// Braces all ungrouped caret exponents.
        ///
        ///   a^2        →  a^{2}
        ///   a^23       →  a^{23}
        ///   a^-2       →  a^{-2}
        ///   a^n        →  a^{n}
        ///   a^{2}      →  a^{2}   (already braced — untouched)
        ///   (a+b)^2    →  (a+b)^{2}
        ///
        /// Uses a character-by-character scan because nested braces make
        /// regex replacement unreliable.
        /// </summary>
        private static string BraceExponents(string s)
        {
            if (string.IsNullOrEmpty(s) || !s.Contains('^'))
                return s;

            var sb = new StringBuilder(s.Length + 16);
            int i = 0;

            while (i < s.Length)
            {
                char c = s[i];

                if (c != '^')
                {
                    sb.Append(c);
                    i++;
                    continue;
                }

                sb.Append('^');
                i++;

                if (i >= s.Length) break;

                char next = s[i];

                // Already braced — copy verbatim
                if (next == '{')
                {
                    int depth = 0;
                    while (i < s.Length)
                    {
                        char ch = s[i];
                        sb.Append(ch);
                        if (ch == '{') { depth++; i++; }
                        else if (ch == '}') { depth--; i++; if (depth == 0) break; }
                        else i++;
                    }
                    continue;
                }

                // Optional sign
                string sign = string.Empty;
                if (next == '-' || next == '+')
                {
                    sign = next.ToString();
                    i++;
                    if (i >= s.Length) { sb.Append(sign); break; }
                    next = s[i];
                }

                // Digit sequence
                if (char.IsDigit(next))
                {
                    int start = i;
                    while (i < s.Length && char.IsDigit(s[i])) i++;
                    sb.Append('{');
                    sb.Append(sign);
                    sb.Append(s, start, i - start);
                    sb.Append('}');
                    continue;
                }

                // Single letter
                if (char.IsLetter(next))
                {
                    sb.Append('{');
                    sb.Append(sign);
                    sb.Append(next);
                    sb.Append('}');
                    i++;
                    continue;
                }

                // Anything else — emit sign and leave for parser
                sb.Append(sign);
            }

            return sb.ToString();
        }

        // =====================================================================
        // PRIVATE — BraceSlashFractions (used in both Stage 1 and Stage 2)
        // =====================================================================

        /// <summary>
        /// Converts simple num/den → \frac{num}{den}.
        ///
        /// Only for simple tokens (digits or single letter on each side).
        /// Complex fractions like (x+1)/(x-1) are intentionally left alone;
        /// the OMML parser in DocxExporter handles / natively for those.
        /// </summary>
        private static string BraceSlashFractions(string s)
        {
            return Regex.Replace(
                s,
                @"(?<!\{)(?<num>[A-Za-z]|\d+)\s*/\s*(?<den>[A-Za-z]|\d+)(?!\})",
                m => $@"\frac{{{m.Groups["num"].Value}}}{{{m.Groups["den"].Value}}}"
            );
        }

        // =====================================================================
        // PRIVATE — StripMathWrappers (Stage 2 only)
        // =====================================================================

        private static string StripMathWrappers(string s)
        {
            s = s.Trim();
            if (s.StartsWith(@"\(") && s.EndsWith(@"\)") && s.Length >= 4)
                return s.Substring(2, s.Length - 4).Trim();
            if (s.StartsWith(@"\[") && s.EndsWith(@"\]") && s.Length >= 4)
                return s.Substring(2, s.Length - 4).Trim();
            if (s.StartsWith("$$") && s.EndsWith("$$") && s.Length >= 4)
                return s.Substring(2, s.Length - 4).Trim();
            if (s.StartsWith("$") && s.EndsWith("$") && s.Length >= 2)
                return s.Substring(1, s.Length - 2).Trim();
            return s;
        }
    }
}
