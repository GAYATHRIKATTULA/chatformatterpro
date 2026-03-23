#nullable enable
// DocxExporter.cs  — COMPLETE FILE
// ════════════════════════════════════════════════════════════════════════════
// PASTE ORDER: Replace your existing DocxExporter.cs entirely with this file.
//
// CHANGES IN THIS VERSION vs the previous session:
//
//   ── TABLE MATH FIX (the only changes) ──────────────────────────────────
//
//   CHANGED:  AppendWordTable()
//     Before: header cells → AppendTextWithBoldRuns (no math awareness)
//             data cells   → AppendMixedContent (math-aware but un-normalized)
//     After:  ALL cells    → AppendTableCellContent() (see below)
//
//   NEW:      AppendTableCellContent(Paragraph, string, MainDocumentPart, bool)
//     Runs the same two-step pre-normalization every paragraph line gets:
//       NormalizeForParsing + NormalizeLatex
//     Then routes to:
//       AppendMixedContent   — if cell has \(...\) or \[...\] delimiters
//       AppendMathToParagraph— if cell is bare math (\frac, x^{2}, etc.)
//       AppendTextWithBoldRuns — if cell is plain text
//     Header cells: plain-text runs are bolded via BoldifyPlainRuns().
//
//   NEW:      BoldifyPlainRuns(Paragraph)
//     Applies Bold to plain W.Run elements without touching M.OfficeMath children.
//
//   NOTHING ELSE CHANGED. All paragraph math, heading, bullet, image, and
//   equation logic is identical to the previous session's version.
//
// ════════════════════════════════════════════════════════════════════════════

using ChatFormatterPro;
using ChatFormatterPro.Exporters;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace ChatFormatterPro.Exporters
{
    public static class DocxExporter
    {
        // ══════════════════════════════════════════════════════════════════════
        // SETTINGS
        // ══════════════════════════════════════════════════════════════════════

        private static bool OneLineOneEquationBox = true;
        private static bool InlineMathOnlyInsideParentheses = true;
        private static bool ConvertHorizontalRules = true;

        // ══════════════════════════════════════════════════════════════════════
        // IMAGE IDs
        // ══════════════════════════════════════════════════════════════════════

        private static uint _imageDocPropId = 1;
        private static uint NextImageId() => _imageDocPropId++;

        // ══════════════════════════════════════════════════════════════════════
        // COMPILED REGEX
        // ══════════════════════════════════════════════════════════════════════

        private static readonly Regex InlineMathRegex =
            new Regex(@"\\\((.*?)\\\)|\\\[(.*?)\\\]",
                RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex BoldRegex =
            new Regex(@"\*\*(.+?)\*\*",
                RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex ImgTokenRegex =
            new Regex(@"\{\{IMG:(?<path>.+?)\}\}",
                RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex ParenthesesMathRegex =
            new Regex(
                @"\((?<m>[^()\r\n]{0,400}" +
                @"(\\frac|\\sqrt|\\pi|\\neq|\\le|\\ge|\\pm|\\times|\\div|\^|_|=|/|×|÷)" +
                @"[^()\r\n]{0,400})\)",
                RegexOptions.Compiled);

        // ══════════════════════════════════════════════════════════════════════
        // XML SAFETY
        // ══════════════════════════════════════════════════════════════════════

        private static string SanitizeForOpenXml(string? s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            var sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                if (char.IsControl(ch) && ch != '\t' && ch != '\r' && ch != '\n') continue;
                if (char.IsHighSurrogate(ch))
                {
                    if (i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
                    { sb.Append(ch); sb.Append(s[++i]); }
                    continue;
                }
                if (char.IsLowSurrogate(ch)) continue;
                sb.Append(ch);
            }
            return sb.ToString();
        }

        // ══════════════════════════════════════════════════════════════════════
        // LATEX NORMALISATION
        // Direction: Unicode → LaTeX only. Never LaTeX → Unicode.
        // ══════════════════════════════════════════════════════════════════════

        private static string NormalizeLatex(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim();

            s = Regex.Replace(s, @"\\left\s*", "");
            s = Regex.Replace(s, @"\\right\s*", "");

            s = s.Replace(@"\dfrac", @"\frac");
            s = s.Replace(@"\tfrac", @"\frac");
            s = s.Replace(@"\cdot", @"\times");

            s = s.Replace("×", @"\times");
            s = s.Replace("÷", @"\div");
            s = s.Replace("±", @"\pm");
            s = s.Replace("≈", @"\approx");
            s = s.Replace("≤", @"\le");
            s = s.Replace("≥", @"\ge");
            s = s.Replace("≠", @"\neq");
            s = s.Replace("π", @"\pi");
            s = s.Replace(@"\Rightarrow", "⇒");
            s = s.Replace(@"\square", "");

            s = Regex.Replace(s, @"\\sqrt\s*(\d+)", m => $@"\sqrt{{{m.Groups[1].Value}}}");
            s = Regex.Replace(s, @"\\sqrt\s*([a-zA-Z])\b", m => $@"\sqrt{{{m.Groups[1].Value}}}");

            s = Regex.Replace(s, @"\\frac\s*(\d+)\s*(\d+)",
                m => $@"\frac{{{m.Groups[1].Value}}}{{{m.Groups[2].Value}}}");
            s = Regex.Replace(s, @"\\frac\s*([a-zA-Z])\s*([a-zA-Z])\b",
                m => $@"\frac{{{m.Groups[1].Value}}}{{{m.Groups[2].Value}}}");

            return s.Trim();
        }

        // ══════════════════════════════════════════════════════════════════════
        // AppendMathToParagraph — schema-safe inline + display equation inserter
        //
        // M.OfficeMath is appended DIRECTLY to W.Paragraph (not inside W.Run).
        // This is the only correct OOXML structure for inline and display math.
        // ══════════════════════════════════════════════════════════════════════

        private static void AppendMathToParagraph(Paragraph p, string latex)
        {
            try
            {
                p.Append(BuildOfficeMath(latex));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"[DocxExporter] OMML build failed.\n  LaTeX: {latex}\n  {ex.Message}");
                p.Append(new Run(
                    new RunProperties(new Color { Val = "CC0000" }),
                    new Text($"[EQ? {SanitizeForOpenXml(latex)}]")
                    { Space = SpaceProcessingModeValues.Preserve }
                ));
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // LINE CLASSIFICATION HELPERS
        // ══════════════════════════════════════════════════════════════════════

        private static bool IsRuleLabelLine(string t)
        {
            if (string.IsNullOrWhiteSpace(t)) return false;
            t = t.TrimStart('\u200B', '\uFEFF').TrimStart();
            return Regex.IsMatch(t, @"^R\s*\d+\s*:", RegexOptions.IgnoreCase);
        }

        private static bool IsRuleBankLine(string t)
        {
            if (string.IsNullOrWhiteSpace(t)) return false;
            t = t.TrimStart('\u200B', '\uFEFF').TrimStart();
            if (Regex.IsMatch(t, @"^(•\s*)?R\s*\d+\s*:", RegexOptions.IgnoreCase)) return true;
            if (Regex.IsMatch(t, @"^[\*\u2217\uFE61\u2731\u2733\uFE0E\uFE0F]+\s*R\s*\d+\s*:",
                    RegexOptions.IgnoreCase)) return true;
            return false;
        }

        private static bool LooksEquationLine(string t)
        {
            if (string.IsNullOrWhiteSpace(t)) return false;
            if (t.Contains(@"\frac") || t.Contains(@"\sqrt") || t.Contains("√")) return true;
            if (t.Contains(")^")) return true;
            if (t.Length <= 80
                && !t.EndsWith(":")
                && !Regex.IsMatch(t, @"^[A-Z][a-z]"))
            {
                bool startsLikeMath = Regex.IsMatch(t, @"^[\(\[\-\+\d\\]");
                if (startsLikeMath && (t.Contains("=") || t.Contains("/"))) return true;
            }
            return false;
        }

        private static bool LooksMathish(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            if (Regex.IsMatch(s, @"\\(frac|sqrt|pi|approx|pm|le|ge|neq|left|right|times|div)\b"))
                return true;
            if (s.Contains("^") || s.Contains("_") || s.Contains("=") || s.Contains("/") ||
                s.Contains("×") || s.Contains("÷") || s.Contains("±") || s.Contains("≈") ||
                s.Contains("≤") || s.Contains("≥") || s.Contains("≠"))
                return true;
            if (Regex.IsMatch(s, @"([A-Za-z]\d|\d[A-Za-z])"))
                return true;
            return false;
        }

        private static bool TryParseOptionMathLine(string line, out string prefix, out string math)
        {
            prefix = ""; math = "";
            if (string.IsNullOrWhiteSpace(line)) return false;
            var m = Regex.Match(line.Trim(),
                @"^(?<prefix>(?:[A-Z]|Step\s+[A-Z])\)\s*)(?<math>\(.+\))$",
                RegexOptions.IgnoreCase);
            if (!m.Success) return false;
            prefix = m.Groups["prefix"].Value;
            math = m.Groups["math"].Value;
            return true;
        }

        private static bool IsHorizontalRuleLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return false;
            var t = line.Trim();
            if (t.Length < 3) return false;
            foreach (char ch in t)
                if (ch != '-' && ch != '_' && ch != '*' && ch != ' ' && ch != '\t') return false;
            return t.Count(c => c == '-') >= 3
                || t.Count(c => c == '*') >= 3
                || t.Count(c => c == '_') >= 3;
        }

        private enum DisplayBlockKind { SquareBracket, LatexBracket }

        private static bool IsDisplayBlockStart(string line, out DisplayBlockKind kind)
        {
            kind = DisplayBlockKind.SquareBracket;
            string t = (line ?? "").Trim();
            if (t == "[") { kind = DisplayBlockKind.SquareBracket; return true; }
            if (t == @"\[") { kind = DisplayBlockKind.LatexBracket; return true; }
            return false;
        }

        private static bool IsDisplayBlockEnd(string line, DisplayBlockKind kind)
        {
            string t = (line ?? "").Trim();
            return kind switch
            {
                DisplayBlockKind.SquareBracket => t == "]",
                DisplayBlockKind.LatexBracket => t == @"\]",
                _ => false
            };
        }

        private static bool TryGetSingleLineDisplayMath(string line, out string math)
        {
            math = "";
            string t = (line ?? "").Trim();
            if ((t.StartsWith(@"\(") && t.EndsWith(@"\)")) ||
                (t.StartsWith(@"\[") && t.EndsWith(@"\]")))
            {
                var m = InlineMathRegex.Match(t);
                if (m.Success && m.Index == 0 && m.Length == t.Length)
                {
                    string raw = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                    math = NormalizeLatex(raw);
                    return !string.IsNullOrWhiteSpace(math);
                }
            }
            if (t.StartsWith("[") && t.EndsWith("]") && t.Length >= 2)
            {
                string inner = t.Substring(1, t.Length - 2);
                if (LooksMathish(inner)) { math = NormalizeLatex(inner); return !string.IsNullOrWhiteSpace(math); }
            }
            return false;
        }

        private static bool TryGetLatexFromInlineMath(string line, out string latex)
        {
            latex = "";
            if (string.IsNullOrWhiteSpace(line)) return false;
            var t = line.Trim();
            if (t.StartsWith(@"\(") && t.EndsWith(@"\)") && t.Length >= 4)
            { latex = t.Substring(2, t.Length - 4).Trim(); return true; }
            if (t.StartsWith("$") && t.EndsWith("$") && t.Length >= 2)
            { latex = t.Substring(1, t.Length - 2).Trim(); return true; }
            return false;
        }

        private static string CleanAIMarkdownArtifacts(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return line;
            line = Regex.Replace(line, @"^\s*(>+)\s*", "");
            line = Regex.Replace(line, @"^\s*(->|=>|>>>+)\s*", "");
            line = Regex.Replace(line, @"^\s*•\s+", "");
            return line;
        }

        private static string ConvertCaretPowersToUnicode(string line)
        {
            if (string.IsNullOrWhiteSpace(line) || !line.Contains("^")) return line;
            static string ToSup(string s)
            {
                var sb = new StringBuilder(s.Length);
                foreach (char ch in s)
                    sb.Append(ch switch
                    {
                        '0' => '⁰',
                        '1' => '¹',
                        '2' => '²',
                        '3' => '³',
                        '4' => '⁴',
                        '5' => '⁵',
                        '6' => '⁶',
                        '7' => '⁷',
                        '8' => '⁸',
                        '9' => '⁹',
                        '-' => '⁻',
                        '+' => '⁺',
                        _ => ch
                    });
                return sb.ToString();
            }
            return Regex.Replace(line, @"\^(?<exp>[+-]?\d+)", m => ToSup(m.Groups["exp"].Value));
        }

        private static string ConvertSimpleSlashFractionToLatex(string s)
        {
            var t = s.Trim();
            if (t.Contains(@"\frac")) return t;
            int slash = t.IndexOf('/');
            if (slash < 0) return t;
            var left = t.Substring(0, slash).Trim();
            var right = t.Substring(slash + 1).Trim();
            if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right)) return t;
            var eq = right.IndexOf('=');
            if (eq >= 0)
            {
                var denom = right.Substring(0, eq).Trim();
                var rest = right.Substring(eq).Trim();
                if (!string.IsNullOrWhiteSpace(denom))
                    return $@"\frac{{{left}}}{{{denom}}} {rest}";
                return t;
            }
            return $@"\frac{{{left}}}{{{right}}}";
        }

        // ══════════════════════════════════════════════════════════════════════
        // NUMBERING
        // ══════════════════════════════════════════════════════════════════════

        private static int EnsureBulletNumbering(MainDocumentPart mainPart)
        {
            var numberingPart = mainPart.NumberingDefinitionsPart
                ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering ??= new Numbering();
            var numbering = numberingPart.Numbering;
            const int abstractNumId = 1, numId = 1;
            if (!numbering.Elements<AbstractNum>().Any(a => a.AbstractNumberId?.Value == abstractNumId))
            {
                var abs = new AbstractNum { AbstractNumberId = abstractNumId };
                var lvl = new Level { LevelIndex = 0 };
                lvl.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });
                lvl.Append(new LevelText { Val = "•" });
                lvl.Append(new LevelJustification { Val = LevelJustificationValues.Left });
                lvl.Append(new ParagraphProperties(new Indentation { Left = "720", Hanging = "360" }));
                abs.Append(lvl);
                numbering.Append(abs);
            }
            if (!numbering.Elements<NumberingInstance>().Any(n => n.NumberID?.Value == numId))
            {
                var inst = new NumberingInstance { NumberID = numId };
                inst.Append(new AbstractNumId { Val = abstractNumId });
                numbering.Append(inst);
            }
            numbering.Save();
            return numId;
        }

        private static void AppendBulletParagraph(Body body, int bulletNumId, string text)
        {
            var p = new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = bulletNumId }
                    )
                )
            );
            p.Append(new Run(
                new Text(SanitizeForOpenXml(text)) { Space = SpaceProcessingModeValues.Preserve }
            ));
            body.Append(p);
        }

        private static void AppendRuleBankParagraph(Body body, int bulletNumId, string text)
        {
            string t = text?.Trim() ?? "";
            t = Regex.Replace(t, @"^[•\-\*\s]+", "");
            string left = t, right = "";
            int arrowIndex = t.IndexOf('→');
            if (arrowIndex >= 0)
            {
                left = t.Substring(0, arrowIndex + 1).Trim();
                right = t.Substring(arrowIndex + 1).Trim();
            }
            var p = new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = bulletNumId }
                    )
                )
            );
            p.Append(new Run(
                new Text(SanitizeForOpenXml(left + " ")) { Space = SpaceProcessingModeValues.Preserve }
            ));
            if (!string.IsNullOrWhiteSpace(right))
            {
                string mathStr = MathLatexHelper.NormalizeForParsing(right);
                mathStr = NormalizeLatex(mathStr);
                AppendMathToParagraph(p, mathStr);
            }
            body.Append(p);
        }

        // ══════════════════════════════════════════════════════════════════════
        // EXPORT — main entry point
        // ══════════════════════════════════════════════════════════════════════

        public static void Export(string content, string filePath, string title)
        {
            content ??= string.Empty;

            if (File.Exists(filePath))
            {
                try { File.Delete(filePath); }
                catch (IOException)
                {
                    MessageBox.Show(
                        "Please close the previously opened Word file before exporting again.",
                        "File In Use", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            using var wordDoc = WordprocessingDocument.Create(
                filePath, WordprocessingDocumentType.Document);
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            _imageDocPropId = 1;

            if (!string.IsNullOrWhiteSpace(title))
            {
                body.Append(new Paragraph(
                    new Run(
                        new RunProperties(new Bold(), new FontSize { Val = "28" }),
                        new Text(SanitizeForOpenXml(title))
                        { Space = SpaceProcessingModeValues.Preserve }
                    )
                ));
                body.Append(new Paragraph(new Run(new Text(""))));
            }

            content = GayathriStyleFormatter.NormalizeContent(content);
            var lines = content.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            int bulletNumId = EnsureBulletNumbering(mainPart);
            bool inWhySection = false;

            for (int i = 0; i < lines.Length; i++)
            {
                string raw = lines[i] ?? string.Empty;
                string t = raw.Trim();

                if (t.Equals("WHY", StringComparison.OrdinalIgnoreCase))
                    inWhySection = true;

                if (inWhySection)
                {
                    if (Regex.IsMatch(t,
                        @"^(Question\s+\d+|Solution|Expression:|Steps:|Final Answer)",
                        RegexOptions.IgnoreCase))
                        inWhySection = false;
                    else if (!t.Equals("WHY", StringComparison.OrdinalIgnoreCase))
                    {
                        var wp = new Paragraph();
                        AppendTextWithBoldRuns(wp, raw);
                        body.Append(wp);
                        continue;
                    }
                }

                // PRIORITY 1: empty line
                if (string.IsNullOrWhiteSpace(t))
                { body.Append(new Paragraph(new Run(new Text("")))); continue; }

                // PRIORITY 2: full-line \( \) or \[ \] math
                if (TryGetLatexFromInlineMath(t, out var inlineTex))
                {
                    string tex = MathLatexHelper.NormalizeForParsing(inlineTex);
                    tex = NormalizeLatex(tex);
                    body.Append(BuildDisplayEquationParagraph(tex));
                    body.Append(new Paragraph(new Run(new Text(""))));
                    continue;
                }

                // PRIORITY 3: multi-line display block \[ ... \]
                if (IsDisplayBlockStart(raw, out var displayKind))
                {
                    var sb = new StringBuilder();
                    i++;
                    while (i < lines.Length && !IsDisplayBlockEnd(lines[i] ?? "", displayKind))
                    { sb.AppendLine(lines[i] ?? ""); i++; }
                    if (i < lines.Length) i++;
                    string math = MathLatexHelper.NormalizeForParsing(sb.ToString());
                    math = NormalizeLatex(math);
                    body.Append(!string.IsNullOrWhiteSpace(math)
                        ? BuildDisplayEquationParagraph(math)
                        : new Paragraph(new Run(new Text(""))));
                    i--;
                    continue;
                }

                // PRIORITY 4: rule bank bullet
                if (IsRuleLabelLine(t) || IsRuleBankLine(t))
                { AppendRuleBankParagraph(body, bulletNumId, t); continue; }

                // PRIORITY 5: option math  A) (expr)
                if (TryParseOptionMathLine(raw, out var optionPrefix, out var optionMath))
                {
                    var p = new Paragraph();
                    p.Append(new Run(new Text(optionPrefix) { Space = SpaceProcessingModeValues.Preserve }));
                    string oTex = MathLatexHelper.NormalizeForParsing(optionMath);
                    oTex = NormalizeLatex(oTex);
                    AppendMathToParagraph(p, oTex);
                    body.Append(p);
                    continue;
                }

                // PRIORITY 6: headings # ## ###
                if (Regex.IsMatch(raw, @"^\s*#{1,6}"))
                {
                    int level = Math.Clamp(Regex.Match(raw, @"^\s*(#+)").Groups[1].Value.Length, 1, 3);
                    string hText = Regex.Replace(raw, @"^\s*#{1,6}\s*", "").Trim();
                    string size = level switch { 1 => "40", 2 => "32", _ => "28" };
                    var p = new Paragraph();
                    p.ParagraphProperties = new ParagraphProperties(
                        new SpacingBetweenLines { Before = "120", After = "120" });
                    AppendMixedContent(p, hText, mainPart);
                    foreach (var run in p.Elements<Run>())
                    {
                        run.RunProperties ??= new RunProperties();
                        run.RunProperties.Bold = new Bold();
                        run.RunProperties.FontSize = new FontSize { Val = size };
                    }
                    body.Append(p);
                    continue;
                }

                // PRIORITY 7: bullet list items  - or *
                if (raw.TrimStart().StartsWith("- ") || raw.TrimStart().StartsWith("* "))
                {
                    string itemText = raw.TrimStart().Substring(2);
                    var p = new Paragraph(
                        new ParagraphProperties(
                            new NumberingProperties(
                                new NumberingLevelReference { Val = 0 },
                                new NumberingId { Val = bulletNumId }
                            )
                        )
                    );
                    AppendMixedContent(p, itemText, mainPart);
                    body.Append(p);
                    continue;
                }

                // PRIORITY 8: horizontal rule
                if (ConvertHorizontalRules && IsHorizontalRuleLine(raw))
                { body.Append(BuildHorizontalRuleParagraph()); continue; }

                // PRIORITY 9: image-only line
                var onlyImg = ImgTokenRegex.Match(t);
                if (onlyImg.Success && onlyImg.Value.Trim() == t)
                { body.Append(BuildImageParagraph(mainPart, onlyImg.Groups["path"].Value.Trim())); continue; }

                // PRIORITY 10: pipe table
                if (IsPipeTableLine(raw))
                {
                    var pipeLines = new List<string>();
                    while (i < lines.Length && IsPipeTableLine(lines[i] ?? ""))
                    { pipeLines.Add(lines[i]!); i++; }
                    i--;
                    var rows = new List<List<string>>();
                    for (int k = 0; k < pipeLines.Count; k++)
                    {
                        if (k == 1 && IsMarkdownSeparatorLine(pipeLines[k])) continue;
                        rows.Add(SplitPipeRow(pipeLines[k]));
                    }
                    AppendWordTable(body, rows, mainPart);
                    continue;
                }

                // PRIORITY 11: single-line display math [expr]
                if (TryGetSingleLineDisplayMath(raw, out string displayMath))
                { body.Append(BuildDisplayEquationParagraph(displayMath)); continue; }

                // PRIORITY 12: equation line (strict detection)
                if (!t.EndsWith(":") && !t.StartsWith("#") && !IsRuleBankLine(t) && LooksEquationLine(t))
                {
                    string eq = ConvertSimpleSlashFractionToLatex(t);
                    eq = MathLatexHelper.NormalizeForParsing(eq);
                    eq = NormalizeLatex(eq);
                    body.Append(BuildDisplayEquationParagraph(eq));
                    body.Append(new Paragraph(new Run(new Text(""))));
                    continue;
                }

                // PRIORITY 13 (default): plain text paragraph
                string line2 = CleanAIMarkdownArtifacts(raw);
                line2 = ConvertCaretPowersToUnicode(line2);
                var para = new Paragraph();
                AppendMixedContent(para, line2, mainPart);
                body.Append(para);
            }

            mainPart.Document.Save();
            FileOpener.Open(filePath);
        }

        // ══════════════════════════════════════════════════════════════════════
        // INLINE CONTENT BUILDER
        // ══════════════════════════════════════════════════════════════════════

        private static void AppendMixedContent(Paragraph p, string line, MainDocumentPart mainPart)
        {
            var imgMatches = ImgTokenRegex.Matches(line);
            if (imgMatches.Count > 0)
            {
                int last = 0;
                foreach (Match m in imgMatches)
                {
                    if (m.Index > last) AppendTextWithBoldRuns(p, line.Substring(last, m.Index - last));
                    var imgPara = BuildImageParagraph(mainPart, m.Groups["path"].Value.Trim());
                    foreach (var run in imgPara.Elements<Run>()) p.Append(run.CloneNode(true));
                    last = m.Index + m.Length;
                }
                if (last < line.Length) AppendTextWithBoldRuns(p, line.Substring(last));
                return;
            }

            var wrapperMatches = InlineMathRegex.Matches(line);
            if (wrapperMatches.Count > 0) { AppendUsingWrappers(p, line, wrapperMatches); return; }

            if (InlineMathOnlyInsideParentheses)
            {
                var pm = ParenthesesMathRegex.Matches(line);
                if (pm.Count > 0) { AppendUsingParenthesesMath(p, line, pm); return; }
            }

            AppendTextWithBoldRuns(p, line);
        }

        private static bool IsMathJoiner(string s)
        {
            if (string.IsNullOrEmpty(s)) return true;
            foreach (char ch in s)
            {
                if (char.IsWhiteSpace(ch)) continue;
                if (ch == ',' || ch == ';' || ch == ':' || ch == '.' || ch == '!' || ch == '?' ||
                    ch == ')' || ch == '(' || ch == ']' || ch == '[') continue;
                return false;
            }
            return true;
        }

        private static void AppendUsingWrappers(Paragraph p, string line, MatchCollection matches)
        {
            int last = 0;
            var mathSb = new StringBuilder();
            bool haveMath = false;

            void FlushMath()
            {
                if (mathSb.Length == 0) return;
                string latex = MathLatexHelper.NormalizeForParsing(mathSb.ToString());
                latex = NormalizeLatex(latex);
                AppendMathToParagraph(p, latex);
                mathSb.Clear();
                haveMath = false;
            }

            foreach (Match m in matches)
            {
                string before = m.Index > last ? line.Substring(last, m.Index - last) : "";
                if (!string.IsNullOrEmpty(before))
                {
                    if (OneLineOneEquationBox && haveMath && IsMathJoiner(before)) mathSb.Append(" ");
                    else { FlushMath(); AppendTextWithBoldRuns(p, before); }
                }
                string mathText = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                if (OneLineOneEquationBox) { if (mathSb.Length > 0) mathSb.Append(" "); mathSb.Append(mathText); haveMath = true; }
                else { string l2 = MathLatexHelper.NormalizeForParsing(mathText); l2 = NormalizeLatex(l2); AppendMathToParagraph(p, l2); }
                last = m.Index + m.Length;
            }
            string after = last < line.Length ? line.Substring(last) : "";
            if (!string.IsNullOrEmpty(after)) { FlushMath(); AppendTextWithBoldRuns(p, after); }
            else FlushMath();
        }

        private static void AppendUsingParenthesesMath(Paragraph p, string line, MatchCollection matches)
        {
            int last = 0;
            foreach (Match m in matches)
            {
                if (m.Index > last) AppendTextWithBoldRuns(p, line.Substring(last, m.Index - last));
                AppendTextWithBoldRuns(p, "(");
                string originalInside = m.Groups["m"].Value;
                string inside = MathLatexHelper.NormalizeForParsing(originalInside);
                inside = NormalizeLatex(inside);
                bool malformed = originalInside.Contains("{") || originalInside.Contains("}")
                    || originalInside.Contains(@"\")
                    || Regex.IsMatch(originalInside, @"[A-Za-z]{3,}\s+[A-Za-z]{3,}");
                bool realMath = LooksMathish(inside) &&
                    Regex.IsMatch(inside, @"[\^_=\/×÷±≈≤≥≠]|\\(frac|sqrt|pi|neq|le|ge|pm|times|div)\b");
                if (string.IsNullOrWhiteSpace(inside) || malformed || !realMath)
                    AppendTextWithBoldRuns(p, originalInside);
                else
                    AppendMathToParagraph(p, inside);
                AppendTextWithBoldRuns(p, ")");
                last = m.Index + m.Length;
            }
            if (last < line.Length) AppendTextWithBoldRuns(p, line.Substring(last));
        }

        // ══════════════════════════════════════════════════════════════════════
        // BOLD RUN BUILDER
        // ══════════════════════════════════════════════════════════════════════

        private static void AppendTextWithBoldRuns(Paragraph p, string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            text = SanitizeForOpenXml(text);
            int idx = 0;
            foreach (Match m in BoldRegex.Matches(text))
            {
                if (m.Index > idx)
                    p.Append(new Run(
                        new Text(SanitizeForOpenXml(text.Substring(idx, m.Index - idx)))
                        { Space = SpaceProcessingModeValues.Preserve }
                    ));
                p.Append(new Run(
                    new RunProperties(new Bold()),
                    new Text(SanitizeForOpenXml(m.Groups[1].Value))
                    { Space = SpaceProcessingModeValues.Preserve }
                ));
                idx = m.Index + m.Length;
            }
            if (idx < text.Length)
                p.Append(new Run(
                    new Text(SanitizeForOpenXml(text.Substring(idx)))
                    { Space = SpaceProcessingModeValues.Preserve }
                ));
        }

        // ══════════════════════════════════════════════════════════════════════
        // DISPLAY EQUATION PARAGRAPH
        // ══════════════════════════════════════════════════════════════════════

        private static Paragraph BuildDisplayEquationParagraph(string mathText)
        {
            var p = new Paragraph();
            p.ParagraphProperties = new ParagraphProperties(
                new Justification { Val = JustificationValues.Left },
                new SpacingBetweenLines { Before = "120", After = "120" }
            );
            AppendMathToParagraph(p, mathText);
            return p;
        }

        // ══════════════════════════════════════════════════════════════════════
        // HORIZONTAL RULE
        // ══════════════════════════════════════════════════════════════════════

        private static Paragraph BuildHorizontalRuleParagraph()
        {
            var p = new Paragraph();
            p.ParagraphProperties = new ParagraphProperties(
                new ParagraphBorders(
                    new BottomBorder { Val = BorderValues.Single, Size = 12, Color = "000000" }
                ),
                new SpacingBetweenLines { Before = "120", After = "120" }
            );
            p.Append(new Run(new Text("")));
            return p;
        }

        // ══════════════════════════════════════════════════════════════════════
        // TABLE  ── the only section that changed in this version ──
        // ══════════════════════════════════════════════════════════════════════

        private static bool IsPipeTableLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return false;
            line = line.Trim();
            return line.Contains("|") && line.Count(c => c == '|') >= 2;
        }

        private static bool IsMarkdownSeparatorLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return false;
            line = line.Trim();
            foreach (char ch in line)
                if (ch != '|' && ch != '-' && ch != ':' && ch != ' ' && ch != '\t') return false;
            return line.Contains("-");
        }

        private static List<string> SplitPipeRow(string line)
        {
            line = (line ?? "").Trim();
            if (line.StartsWith("|")) line = line.Substring(1);
            if (line.EndsWith("|")) line = line.Substring(0, line.Length - 1);
            // NOTE: We do NOT call SanitizeForOpenXml here anymore.
            // Sanitization is done per-run inside AppendTableCellContent so that
            // math delimiters (\( \) \frac etc.) survive into the math pipeline.
            var parts = line.Split('|').Select(x => x.Trim()).ToList();
            while (parts.Count > 0 && parts[0] == "") parts.RemoveAt(0);
            while (parts.Count > 0 && parts[^1] == "") parts.RemoveAt(parts.Count - 1);
            return parts;
        }

        // ── CHANGED ─────────────────────────────────────────────────────────
        // AppendWordTable: both header and data cells now use AppendTableCellContent
        // instead of the old header→AppendTextWithBoldRuns / data→AppendMixedContent split.
        // Table structure (borders, rows, columns) is unchanged.

        private static void AppendWordTable(Body body, List<List<string>> rows, MainDocumentPart mainPart)
        {
            if (rows == null || rows.Count == 0) return;

            var table = new Table();
            var tblProps = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 8 },
                    new BottomBorder { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder { Val = BorderValues.Single, Size = 8 },
                    new RightBorder { Val = BorderValues.Single, Size = 8 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 8 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 8 }
                )
            );
            table.AppendChild(tblProps);

            int colCount = rows.Max(r => r.Count);
            foreach (var r in rows)
                while (r.Count < colCount) r.Add("");

            for (int r = 0; r < rows.Count; r++)
            {
                var tr = new TableRow();

                for (int c = 0; c < colCount; c++)
                {
                    var tc = new TableCell();
                    var p = new Paragraph();
                    bool isHeader = (r == 0);

                    AppendTableCellContent(p, rows[r][c], mainPart, isHeader);

                    tc.Append(p);
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto }
                    ));
                    tr.Append(tc);
                }

                table.Append(tr);
            }

            body.Append(table);
            body.Append(new Paragraph(new Run(new Text(""))));
        }

        // ── NEW ──────────────────────────────────────────────────────────────
        // AppendTableCellContent: math pipeline for individual table cells.
        //
        // Four cell cases handled:
        //   Case A: cell has \(...\) or \[...\] delimiters
        //           → AppendMixedContent handles mixed text+math correctly
        //           → header cells: plain-text runs are bolded afterward
        //
        //   Case B: bare math without delimiters (\frac{1}{4}, x^{2}+1, etc.)
        //           → detected by LooksMathish after normalization
        //           → rendered as a single equation via AppendMathToParagraph
        //
        //   Case C: plain text (including ?, —, labels like "Simplify:")
        //           → AppendTextWithBoldRuns; header cells get Bold applied
        //
        //   Case D: empty cell
        //           → single empty run so the cell is not malformed in Word

        private static void AppendTableCellContent(
            Paragraph p,
            string cellText,
            MainDocumentPart mainPart,
            bool isHeader)
        {
            // Case D: empty cell
            if (string.IsNullOrEmpty(cellText))
            {
                p.Append(new Run(
                    new Text("") { Space = SpaceProcessingModeValues.Preserve }
                ));
                return;
            }

            // Pre-normalize: same two steps every equation-line in Export() goes through.
            // NormalizeForParsing: braces exponents, converts simple slash fractions
            // NormalizeLatex:      Unicode operators → \commands, \dfrac → \frac, etc.
            string normalized = MathLatexHelper.NormalizeForParsing(cellText);
            normalized = NormalizeLatex(normalized);

            // Case A: explicit math delimiters present → AppendMixedContent
            bool hasDelimiters = InlineMathRegex.IsMatch(normalized);
            if (hasDelimiters)
            {
                AppendMixedContent(p, normalized, mainPart);
                if (isHeader) BoldifyPlainRuns(p);
                return;
            }

            // Case B: bare math (no delimiters but content looks like math)
            // Guard: skip short strings that are just operators or single chars,
            // skip strings that look like prose labels (start with a capital word).
            bool looksLikeMath = LooksMathish(normalized)
                && !normalized.Trim().EndsWith(":")
                && !Regex.IsMatch(normalized.Trim(), @"^[A-Z][a-z]{2,}");

            if (looksLikeMath)
            {
                AppendMathToParagraph(p, normalized);
                return;
            }

            // Case C: plain text
            // Use the original cellText (not normalized) so that plain-text content
            // like "Step 1" or "?" is not accidentally altered by NormalizeLatex.
            AppendTextWithBoldRuns(p, cellText);
            if (isHeader) BoldifyPlainRuns(p);
        }

        // ── NEW ──────────────────────────────────────────────────────────────
        // BoldifyPlainRuns: applies Bold to every plain W.Run in the paragraph
        // that does not already have Bold set. M.OfficeMath children are not
        // W.Run objects and are never touched by this method.

        private static void BoldifyPlainRuns(Paragraph p)
        {
            foreach (var run in p.Elements<Run>())
            {
                run.RunProperties ??= new RunProperties();
                if (run.RunProperties.Bold == null)
                    run.RunProperties.Bold = new Bold();
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // IMAGES
        // ══════════════════════════════════════════════════════════════════════

        private static Paragraph BuildImageParagraph(MainDocumentPart mainPart, string imagePath)
        {
            var p = new Paragraph();
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
            {
                p.Append(new Run(
                    new Text(SanitizeForOpenXml($"[Missing image: {imagePath}]"))
                    { Space = SpaceProcessingModeValues.Preserve }
                ));
                return p;
            }

            string ext = Path.GetExtension(imagePath).ToLowerInvariant();
            string contentType = ext switch
            {
                ".png" => "image/png",
                ".jpg" => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                ".tif" => "image/tiff",
                ".tiff" => "image/tiff",
                _ => "image/png"
            };

            var imagePart = mainPart.AddImagePart(contentType);
            using (var stream = File.OpenRead(imagePath)) imagePart.FeedData(stream);
            string relId = mainPart.GetIdOfPart(imagePart);

            long widthPx = 800, heightPx = 450;
            try
            {
                using var fs = File.OpenRead(imagePath);
                var decoder = System.Windows.Media.Imaging.BitmapDecoder.Create(fs,
                    System.Windows.Media.Imaging.BitmapCreateOptions.PreservePixelFormat,
                    System.Windows.Media.Imaging.BitmapCacheOption.OnLoad);
                widthPx = decoder.Frames[0].PixelWidth;
                heightPx = decoder.Frames[0].PixelHeight;
            }
            catch { }

            const long emusPerInch = 914400, dpi = 96;
            long cx = widthPx * emusPerInch / dpi, cy = heightPx * emusPerInch / dpi;
            long maxCx = (long)(6.5 * emusPerInch);
            if (cx > maxCx) { double sc = (double)maxCx / cx; cx = maxCx; cy = (long)(cy * sc); }

            uint docPropId = NextImageId(), nvId = NextImageId();

            var drawing = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = cx, Cy = cy },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = (UInt32Value)docPropId, Name = "Picture" },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties
                                    { Id = (UInt32Value)nvId, Name = SanitizeForOpenXml(Path.GetFileName(imagePath)) },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = relId, CompressionState = A.BlipCompressionValues.Print },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = cx, Cy = cy }),
                                    new A.PresetGeometry(new A.AdjustValueList())
                                    { Preset = A.ShapeTypeValues.Rectangle })
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U }
            );

            p.ParagraphProperties = new ParagraphProperties(
                new Justification { Val = JustificationValues.Left },
                new SpacingBetweenLines { Before = "120", After = "120" }
            );
            p.Append(new Run(drawing));
            return p;
        }

        // ══════════════════════════════════════════════════════════════════════
        // OMML BUILDER — core math engine
        // AppendMathToParagraph is the only public-facing call.
        // BuildOfficeMath is the engine. ParseSequence is the recursive parser.
        // ══════════════════════════════════════════════════════════════════════

        private static M.OfficeMath BuildOfficeMath(string latex)
        {
            latex = MathLatexHelper.NormalizeForParsing(latex ?? "");
            latex = NormalizeLatex(latex);
            var ts = new TokenStream(latex);
            var math = new M.OfficeMath();
            foreach (var el in ParseSequence(ts))
                math.Append(el);
            return math;
        }

        private static OpenXmlElement BuildFrac(OpenXmlElement num, OpenXmlElement den)
            => new M.Fraction(
                new M.FractionProperties(),
                new M.Numerator(new M.Base(num)),
                new M.Denominator(new M.Base(den))
            );

        private static OpenXmlElement BuildSqrt(OpenXmlElement radicand)
        {
            var rad = new M.Radical();

            var radPr = new M.RadicalProperties();
            radPr.Append(new M.HideDegree());
            rad.Append(radPr);

            var baseEl = new M.Base();
            baseEl.Append(radicand.CloneNode(true));
            rad.Append(baseEl);

            return rad;
        }

        private static OpenXmlElement BuildRow(IEnumerable<OpenXmlElement> elements)
        {
            var box = new M.Box(new M.BoxProperties());
            var b = new M.Base();
            foreach (var el in elements) b.Append(el);
            box.Append(b);
            return box;
        }

        private static OpenXmlElement ElementsToElement(List<OpenXmlElement> els)
        {
            if (els.Count == 0) return BuildMRun("");
            if (els.Count == 1) return els[0];
            return BuildRow(els);
        }

        private static M.Superscript BuildSup(OpenXmlElement @base, OpenXmlElement supArg)
        {
            var s = new M.Superscript();
            var b = new M.Base(); b.Append(@base.CloneNode(true));
            var a = new M.SuperArgument(); a.Append(supArg.CloneNode(true));
            s.Append(b); s.Append(a);
            return s;
        }

        private static M.Subscript BuildSub(OpenXmlElement @base, OpenXmlElement subArg)
        {
            var s = new M.Subscript();
            var b = new M.Base(); b.Append(@base.CloneNode(true));
            var a = new M.SubArgument(); a.Append(subArg.CloneNode(true));
            s.Append(b); s.Append(a);
            return s;
        }

        private static M.Run BuildMRun(string text)
            => new M.Run(new M.Text(SanitizeForOpenXml(text)));

        private static List<OpenXmlElement> ParseSequence(TokenStream ts, char? stopAt = null)
        {
            var list = new List<OpenXmlElement>();

            while (!ts.End)
            {
                if (stopAt.HasValue && ts.Peek() == stopAt.Value) break;
                if (char.IsWhiteSpace(ts.Peek())) { ts.Read(); list.Add(BuildMRun(" ")); continue; }

                if (ts.StartsWith(@"\frac"))
                {
                    ts.Consume(@"\frac");
                    var n = ParseGroupOrAtomElements(ts);
                    var d = ParseGroupOrAtomElements(ts);
                    list.Add(BuildFrac(ElementsToElement(n), ElementsToElement(d)));
                    continue;
                }
                if (ts.StartsWith(@"\sqrt"))
                {
                    ts.Consume(@"\sqrt");
                    list.Add(BuildSqrt(ElementsToElement(ParseGroupOrAtomElements(ts))));
                    continue;
                }

                if (ts.StartsWith(@"\pi")) { ts.Consume(@"\pi"); list.Add(BuildMRun("π")); continue; }
                if (ts.StartsWith(@"\approx")) { ts.Consume(@"\approx"); list.Add(BuildMRun("≈")); continue; }
                if (ts.StartsWith(@"\times")) { ts.Consume(@"\times"); list.Add(BuildMRun("×")); continue; }
                if (ts.StartsWith(@"\cdot")) { ts.Consume(@"\cdot"); list.Add(BuildMRun("·")); continue; }
                if (ts.StartsWith(@"\div")) { ts.Consume(@"\div"); list.Add(BuildMRun("÷")); continue; }
                if (ts.StartsWith(@"\pm")) { ts.Consume(@"\pm"); list.Add(BuildMRun("±")); continue; }
                if (ts.StartsWith(@"\le")) { ts.Consume(@"\le"); list.Add(BuildMRun("≤")); continue; }
                if (ts.StartsWith(@"\ge")) { ts.Consume(@"\ge"); list.Add(BuildMRun("≥")); continue; }
                if (ts.StartsWith(@"\neq")) { ts.Consume(@"\neq"); list.Add(BuildMRun("≠")); continue; }
                if (ts.StartsWith(@"\infty")) { ts.Consume(@"\infty"); list.Add(BuildMRun("∞")); continue; }
                if (ts.StartsWith(@"\alpha")) { ts.Consume(@"\alpha"); list.Add(BuildMRun("α")); continue; }
                if (ts.StartsWith(@"\beta")) { ts.Consume(@"\beta"); list.Add(BuildMRun("β")); continue; }
                if (ts.StartsWith(@"\theta")) { ts.Consume(@"\theta"); list.Add(BuildMRun("θ")); continue; }
                if (ts.StartsWith(@"\Delta")) { ts.Consume(@"\Delta"); list.Add(BuildMRun("Δ")); continue; }
                if (ts.StartsWith(@"\delta")) { ts.Consume(@"\delta"); list.Add(BuildMRun("δ")); continue; }
                if (ts.StartsWith(@"\sigma")) { ts.Consume(@"\sigma"); list.Add(BuildMRun("σ")); continue; }
                if (ts.StartsWith(@"\mu")) { ts.Consume(@"\mu"); list.Add(BuildMRun("μ")); continue; }
                if (ts.StartsWith(@"\lambda")) { ts.Consume(@"\lambda"); list.Add(BuildMRun("λ")); continue; }

                if (!ts.End && ts.Peek() == '\\')
                {
                    ts.Read();
                    string cmd = ts.ReadWhile(char.IsLetter);
                    list.Add(BuildMRun(@"\" + cmd));
                    continue;
                }

                OpenXmlElement atom = ParseAtom(ts);

                while (!ts.End && (ts.Peek() == '^' || ts.Peek() == '_'))
                {
                    char op = ts.Read();
                    var argEl = ElementsToElement(ParseGroupOrAtomElements(ts));
                    atom = op == '^' ? (OpenXmlElement)BuildSup(atom, argEl) : BuildSub(atom, argEl);
                }

                while (!ts.End && ts.Peek() == '/')
                {
                    ts.Read();
                    OpenXmlElement right = ParseAtom(ts);
                    while (!ts.End && (ts.Peek() == '^' || ts.Peek() == '_'))
                    {
                        char op2 = ts.Read();
                        var a2 = ElementsToElement(ParseGroupOrAtomElements(ts));
                        right = op2 == '^' ? (OpenXmlElement)BuildSup(right, a2) : BuildSub(right, a2);
                    }
                    atom = BuildFrac(atom, right);
                }

                list.Add(atom);
            }
            return list;
        }

        private static OpenXmlElement ParseAtom(TokenStream ts)
        {
            if (!ts.End && (ts.Peek() == '-' || ts.Peek() == '+'))
            {
                char sign = ts.Read();
                var nextEl = ElementsToElement(ParseGroupOrAtomElements(ts));
                return BuildRow(new List<OpenXmlElement> { BuildMRun(sign.ToString()), nextEl });
            }
            if (ts.End) return BuildMRun("");

            if (ts.Peek() == '(')
            {
                ts.Read();
                var inner = ParseSequence(ts, stopAt: ')');
                if (!ts.End && ts.Peek() == ')') ts.Read();
                return BuildRow(new List<OpenXmlElement>
                    { BuildMRun("("), ElementsToElement(inner), BuildMRun(")") });
            }
            if (ts.Peek() == '[')
            {
                ts.Read();
                var inner = ParseSequence(ts, stopAt: ']');
                if (!ts.End && ts.Peek() == ']') ts.Read();
                return BuildRow(new List<OpenXmlElement>
                    { BuildMRun("["), ElementsToElement(inner), BuildMRun("]") });
            }
            if (ts.Peek() == '{')
                return ElementsToElement(ParseGroupElements(ts));

            if (char.IsLetter(ts.Peek())) return BuildMRun(ts.ReadWhile(char.IsLetter));
            if (char.IsDigit(ts.Peek())) return BuildMRun(ts.ReadWhile(char.IsDigit));
            return BuildMRun(ts.Read().ToString());
        }

        private static List<OpenXmlElement> ParseGroupElements(TokenStream ts)
        {
            if (ts.End || ts.Peek() != '{') return new List<OpenXmlElement> { BuildMRun("") };
            ts.Read();
            var inner = ParseSequence(ts, stopAt: '}');
            if (!ts.End && ts.Peek() == '}') ts.Read();
            return inner;
        }

        private static List<OpenXmlElement> ParseGroupOrAtomElements(TokenStream ts)
        {
            if (!ts.End && ts.Peek() == '{') return ParseGroupElements(ts);
            if (ts.StartsWith(@"\frac"))
            {
                ts.Consume(@"\frac");
                var n = ParseGroupOrAtomElements(ts); var d = ParseGroupOrAtomElements(ts);
                return new List<OpenXmlElement> { BuildFrac(ElementsToElement(n), ElementsToElement(d)) };
            }
            if (ts.StartsWith(@"\sqrt"))
            {
                ts.Consume(@"\sqrt");
                return new List<OpenXmlElement> { BuildSqrt(ElementsToElement(ParseGroupOrAtomElements(ts))) };
            }
            if (ts.StartsWith(@"\pi")) { ts.Consume(@"\pi"); return new List<OpenXmlElement> { BuildMRun("π") }; }
            if (ts.StartsWith(@"\pm")) { ts.Consume(@"\pm"); return new List<OpenXmlElement> { BuildMRun("±") }; }
            if (ts.StartsWith(@"\approx")) { ts.Consume(@"\approx"); return new List<OpenXmlElement> { BuildMRun("≈") }; }
            if (ts.StartsWith(@"\neq")) { ts.Consume(@"\neq"); return new List<OpenXmlElement> { BuildMRun("≠") }; }
            if (ts.StartsWith(@"\times")) { ts.Consume(@"\times"); return new List<OpenXmlElement> { BuildMRun("×") }; }
            if (ts.StartsWith(@"\div")) { ts.Consume(@"\div"); return new List<OpenXmlElement> { BuildMRun("÷") }; }
            if (ts.StartsWith(@"\le")) { ts.Consume(@"\le"); return new List<OpenXmlElement> { BuildMRun("≤") }; }
            if (ts.StartsWith(@"\ge")) { ts.Consume(@"\ge"); return new List<OpenXmlElement> { BuildMRun("≥") }; }
            return new List<OpenXmlElement> { ParseAtom(ts) };
        }

        private sealed class TokenStream
        {
            private readonly string _s;
            private int _i;
            public TokenStream(string s) { _s = s ?? ""; _i = 0; }
            public bool End => _i >= _s.Length;
            public char Peek() => End ? '\0' : _s[_i];
            public char Read() => End ? '\0' : _s[_i++];
            public string ReadWhile(Func<char, bool> pred)
            {
                int start = _i;
                while (!End && pred(Peek())) _i++;
                return _s.Substring(start, _i - start);
            }
            public bool StartsWith(string t)
            {
                if (t == null || _i + t.Length > _s.Length) return false;
                return string.Compare(_s, _i, t, 0, t.Length, StringComparison.Ordinal) == 0;
            }
            public void Consume(string t) { if (StartsWith(t)) _i += t.Length; }
        }
    }
}
