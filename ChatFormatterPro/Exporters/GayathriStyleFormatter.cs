// GayathriStyleFormatter.cs
// PASTE THIS FILE: Replace your existing GayathriStyleFormatter.cs entirely.
//
// KEY CHANGE vs original:
//   REMOVED the caret-to-Unicode superscript conversion (the Regex.Replace for \^(?<exp>...))
//   from WorksheetNormalizeLine. That block was converting:
//       2^3  →  2³
//   BEFORE the OMML builder ever saw the text, destroying all proper Word equation output.
//
//   The ^ character must survive normalization untouched so MathLatexHelper and
//   DocxExporter's OMML parser can build real Microsoft Word equation superscripts.
//
//   All other behavior is identical to your original file.

using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ChatFormatterPro.Exporters
{
    public static class GayathriStyleFormatter
    {
        // --- Detect if content is "worksheet-like" ---
        private static bool LooksLikeWorksheet(string fullText)
        {
            if (string.IsNullOrWhiteSpace(fullText)) return false;

            string t = fullText.ToLowerInvariant();

            string[] keys =
            {
                "think & do", "worksheet", "pattern", "rule bank",
                "question 1", "question 2", "match the rule", "steps", "step-by-step",
                "h o t s", "hots", "level: grade"
            };

            return keys.Any(k => t.Contains(k));
        }

        // --- Universal (safe) cleanups for all ChatGPT content ---
        private static string UniversalNormalizeLine(string line)
        {
            if (line == null) return "";

            // Remove common ChatGPT prefixes at line start
            line = Regex.Replace(line, @"^\s*(>+)\s*", "");              // blockquote
            line = Regex.Replace(line, @"^\s*(->|=>|>>>+)\s*", "");      // arrows

            // Remove zero-width chars sometimes copied from web
            line = line.Replace("\u200B", "").Replace("\uFEFF", "");

            // Keep indentation as-is; only trim right side
            return line.TrimEnd();
        }

        // --- Worksheet-only formatting rules ---
        //
        // DESIGN RULES (unchanged from original):
        //   - Do NOT create Word bullets (•) for rule lines here.
        //   - DocxExporter.cs decides bullets (real Word numbering).
        //
        // MATH SAFETY RULE (new):
        //   - Do NOT convert ^ to Unicode superscripts (², ³ etc.) here.
        //     The caret must survive so the OMML pipeline can create real Word equations.
        //     If you want display-only Unicode (e.g. HTML export), do it in a separate
        //     HtmlExporter path, never before DOCX export.
        private static string WorksheetNormalizeLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return line;

            // 1) Remove markdown/unicode star markers anywhere (***, ∗∗∗, ✱✱ etc.)
            line = Regex.Replace(line, @"[\*\u2217\uFE61\u2731\u2733\uFE0E\uFE0F]{2,}", "");

            // 2) Remove leftover bold markers **...**
            //    We only remove the markers, not the content.
            //    (DocxExporter's BoldRegex handles **bold** → actual bold runs.)

            string t = line.TrimStart();

            // 3) Emoji → simple symbols (worksheet-only, safe)
            //    NOTE: We do NOT output "• " here for rule bank bullets.
            line = line.Replace("🟣", "◉ ")
                       .Replace("🔷", "◇ ")
                       .Replace("🔹", "• ")
                       .Replace("✅", "✓ ")
                       .Replace("☑️", "✓ ")
                       .Replace("✔️", "✓ ");

            t = line.TrimStart();

            // 4) RULE BANK heading
            if (Regex.IsMatch(t, @"^RULE\s*BANK\b", RegexOptions.IgnoreCase))
                return "◇ " + t.Trim();

            // 5) "Question N" heading
            if (Regex.IsMatch(t, @"^Question\s+\d+\b", RegexOptions.IgnoreCase))
                return "◇ " + t.Trim();

            // 6) Rule lines: keep as plain "R1: ..." (NO bullets here!)
            //    DocxExporter promotes these to real Word bullet list items.
            if (Regex.IsMatch(t, @"^R\s*\d+\s*:", RegexOptions.IgnoreCase))
                return t.Trim();

            // ── REMOVED FROM ORIGINAL ──────────────────────────────────────────
            // The original code had this block here:
            //
            //   line = Regex.Replace(line, @"\^(?<exp>[+\-]?\d+)", m => {
            //       string exp = m.Groups["exp"].Value;
            //       return exp.Replace("+","⁺").Replace("-","⁻")...;
            //   });
            //
            // THIS WAS THE PRIMARY BUG.  It converted 2^3 → 2³ universally,
            // including on lines that would become Word equations.  The OMML
            // parser in DocxExporter then received "2³" and emitted it as a
            // plain text run instead of a proper m:sSup superscript element.
            //
            // FIX: The ^ character is now left completely untouched here.
            //      MathLatexHelper.NormalizeForParsing() braces it correctly,
            //      and DocxExporter's OmmlBuilder produces the proper Word
            //      native equation object.
            // ───────────────────────────────────────────────────────────────────

            return line;
        }

        /// <summary>
        /// Normalize whole content (recommended entry point).
        /// Keeps non-worksheet content safe + applies worksheet rules when detected.
        /// </summary>
        public static string NormalizeContent(string content)
        {
            if (content == null) return "";

            bool worksheetMode = LooksLikeWorksheet(content);

            var lines = content.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            var sb = new StringBuilder(content.Length);

            foreach (var ln in lines)
            {
                var line = UniversalNormalizeLine(ln);
                if (worksheetMode)
                    line = WorksheetNormalizeLine(line);

                sb.AppendLine(line);
            }

            return sb.ToString().TrimEnd('\n');
        }

        /// <summary>
        /// Line-based API (kept for compatibility).
        /// Prefer NormalizeContent() for correct behavior.
        /// </summary>
        public static string NormalizeLine(string line) => UniversalNormalizeLine(line);
    }
}
