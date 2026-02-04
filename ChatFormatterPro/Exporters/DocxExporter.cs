using System;
using System.Collections.Generic;
using System.Linq; // ✅ REQUIRED for .Any() and LINQ
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using M = DocumentFormat.OpenXml.Math;

using ChatFormatterPro; // ✅ to access FileOpener

namespace ChatFormatterPro.Exporters
{
    public static class DocxExporter
    {
        // Finds \( ... \) or \[ ... \]
        private static readonly Regex InlineMathRegex =
            new Regex(@"\\\((.*?)\\\)|\\\[(.*?)\\\]", RegexOptions.Singleline);

        // Detects **bold**
        private static readonly Regex BoldRegex =
            new Regex(@"\*\*(.+?)\*\*", RegexOptions.Singleline);

        public static void Export(string content, string filePath, string title)
        {
            content ??= string.Empty;

            using var wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            // Title
            if (!string.IsNullOrWhiteSpace(title))
            {
                body.Append(
                    new Paragraph(
                        new Run(
                            new RunProperties(new Bold(), new FontSize { Val = "28" }),
                            new Text(title) { Space = SpaceProcessingModeValues.Preserve }
                        )
                    )
                );
                body.Append(new Paragraph(new Run(new Text(""))));
            }

            // Split into lines
            var lines = content.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');

            // ✅ Bullet numbering id (int)
            int bulletNumId = EnsureBulletNumbering(mainPart);

            for (int i = 0; i < lines.Length; i++)
            {
                var rawLine = lines[i];
                var line = (rawLine ?? string.Empty);

                // Blank line
                if (string.IsNullOrWhiteSpace(line))
                {
                    body.Append(new Paragraph(new Run(new Text(""))));
                    continue;
                }

                // ✅ TABLE detection (Markdown pipe table)
                if (IsPipeTableLine(line))
                {
                    var pipeLines = new List<string>();

                    // collect contiguous table lines
                    while (i < lines.Length && IsPipeTableLine(lines[i]))
                    {
                        pipeLines.Add(lines[i]);
                        i++;
                    }

                    // step back (because for-loop will i++)
                    i--;

                    // convert to rows (skip separator row if present)
                    var rows = new List<List<string>>();
                    for (int k = 0; k < pipeLines.Count; k++)
                    {
                        if (k == 1 && IsMarkdownSeparatorLine(pipeLines[k]))
                            continue;

                        rows.Add(SplitPipeRow(pipeLines[k]));
                    }

                    AppendWordTable(body, rows);
                    continue;
                }

                // Heading 1: # Title
                if (line.StartsWith("# "))
                {
                    var text = line.Substring(2).Trim();
                    var p = new Paragraph();
                    p.Append(new Run(
                        new RunProperties(new Bold(), new FontSize { Val = "40" }),
                        new Text(text) { Space = SpaceProcessingModeValues.Preserve }
                    ));
                    body.Append(p);
                    continue;
                }

                // Heading 2: ## Title (supports **bold**)
                if (line.StartsWith("## "))
                {
                    var text = line.Substring(3).Trim();
                    var p = new Paragraph();

                    // Apply heading font size
                    var r = new Run();
                    r.RunProperties = new RunProperties(
                        new FontSize { Val = "32" },
                        new Bold()
                    );
                    p.Append(r);

                    // preserve **bold**
                    AppendTextWithBoldRuns(p, text);

                    body.Append(p);
                    continue;
                }

                // Bullet: - item  OR  * item
                if (line.StartsWith("- ") || line.StartsWith("* "))
                {
                    var itemText = line.Substring(2);

                    var p = new Paragraph(
                        new ParagraphProperties(
                            new NumberingProperties(
                                new NumberingLevelReference() { Val = 0 },
                                new NumberingId() { Val = bulletNumId }
                            )
                        )
                    );

                    AppendMixedContent(p, itemText);
                    body.Append(p);
                    continue;
                }

                // Normal line
                {
                    var p = new Paragraph();
                    AppendMixedContent(p, line);
                    body.Append(p);
                }
            }

            // ✅ Save the document
            mainPart.Document.Save();

            // ✅ Auto-open after export
            FileOpener.Open(filePath);
        }

        // -------------------- TEXT + MATH MIX --------------------

        private static void AppendMixedContent(Paragraph p, string line)
        {
            int last = 0;
            var matches = InlineMathRegex.Matches(line);

            // No math markers -> text only (but preserve **bold**)
            if (matches.Count == 0)
            {
                AppendTextWithBoldRuns(p, line);
                return;
            }

            foreach (Match m in matches)
            {
                int start = m.Index;
                int len = m.Length;

                // text before math (preserve **bold**)
                if (start > last)
                {
                    var before = line.Substring(last, start - last);
                    AppendTextWithBoldRuns(p, before);
                }

                // math content (group1 for \( \), group2 for \[ \])
                var mathText = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;

                // Append as Word equation (OMML)
                var omml = BuildOfficeMath(mathText);
                p.Append(omml);

                last = start + len;
            }

            // remaining text (preserve **bold**)
            if (last < line.Length)
            {
                var after = line.Substring(last);
                AppendTextWithBoldRuns(p, after);
            }
        }

        // ✅ Supports: normal text + **bold text** mixed
        private static void AppendTextWithBoldRuns(Paragraph p, string text)
        {
            if (string.IsNullOrEmpty(text))
                return;

            int idx = 0;

            foreach (Match m in BoldRegex.Matches(text))
            {
                // normal part before bold
                if (m.Index > idx)
                {
                    var normal = text.Substring(idx, m.Index - idx);
                    p.Append(new Run(new Text(normal) { Space = SpaceProcessingModeValues.Preserve }));
                }

                // bold part
                var boldText = m.Groups[1].Value;
                p.Append(
                    new Run(
                        new RunProperties(new Bold()),
                        new Text(boldText) { Space = SpaceProcessingModeValues.Preserve }
                    )
                );

                idx = m.Index + m.Length;
            }

            // remaining normal part
            if (idx < text.Length)
            {
                var remaining = text.Substring(idx);
                p.Append(new Run(new Text(remaining) { Space = SpaceProcessingModeValues.Preserve }));
            }
        }

        // -------------------- BULLET NUMBERING --------------------

        private static int EnsureBulletNumbering(MainDocumentPart mainPart)
        {
            var numberingPart = mainPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering();
            }

            if (numberingPart.Numbering == null)
                numberingPart.Numbering = new Numbering();

            var numbering = numberingPart.Numbering;

            const int abstractNumId = 1;
            const int numId = 1;

            // AbstractNum (bullet style)
            if (!numbering.Elements<AbstractNum>().Any(a => a.AbstractNumberId?.Value == abstractNumId))
            {
                var abs = new AbstractNum() { AbstractNumberId = abstractNumId };

                var lvl = new Level() { LevelIndex = 0 };
                lvl.Append(new NumberingFormat() { Val = NumberFormatValues.Bullet });
                lvl.Append(new LevelText() { Val = "•" });
                lvl.Append(new LevelJustification() { Val = LevelJustificationValues.Left });
                lvl.Append(new ParagraphProperties(
                    new Indentation() { Left = "720", Hanging = "360" }
                ));

                abs.Append(lvl);
                numbering.Append(abs);
            }

            // NumberingInstance (w:num)
            if (!numbering.Elements<NumberingInstance>().Any(n => n.NumberID?.Value == numId))
            {
                var inst = new NumberingInstance() { NumberID = numId };
                inst.Append(new AbstractNumId() { Val = abstractNumId });
                numbering.Append(inst);
            }

            numbering.Save();
            return numId;
        }

        // -------------------- TABLE SUPPORT (Markdown pipe tables) --------------------

        private static bool IsPipeTableLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return false;
            line = line.Trim();

            // must contain at least 2 pipes to look like a row
            return line.Contains("|") && line.Count(c => c == '|') >= 2;
        }

        private static bool IsMarkdownSeparatorLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return false;
            line = line.Trim();

            // accept only pipes, dashes, colons, spaces/tabs
            foreach (char ch in line)
            {
                if (ch != '|' && ch != '-' && ch != ':' && ch != ' ' && ch != '\t')
                    return false;
            }

            // must contain at least one dash
            return line.Contains("-");
        }

        private static List<string> SplitPipeRow(string line)
        {
            line = (line ?? "").Trim();

            // Remove leading/trailing pipe
            if (line.StartsWith("|")) line = line.Substring(1);
            if (line.EndsWith("|")) line = line.Substring(0, line.Length - 1);

            var parts = line.Split('|')
                            .Select(x => x.Trim())
                            .ToList();

            // remove empty edges
            while (parts.Count > 0 && parts.Count > 0 && parts[0] == "") parts.RemoveAt(0);
            while (parts.Count > 0 && parts[^1] == "") parts.RemoveAt(parts.Count - 1);

            return parts;
        }

        private static void AppendWordTable(Body body, List<List<string>> rows)
        {
            if (rows == null || rows.Count == 0) return;

            var table = new Table();

            // Borders
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

            // normalize columns
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

                    if (r == 0)
                    {
                        // header row bold
                        var run = new Run(new RunProperties(new Bold()));
                        p.Append(run);
                        AppendTextWithBoldRuns(p, rows[r][c]);
                    }
                    else
                    {
                        AppendMixedContent(p, rows[r][c]);
                    }

                    tc.Append(p);
                    tc.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                    tr.Append(tc);
                }

                table.Append(tr);
            }

            body.Append(table);
            body.Append(new Paragraph(new Run(new Text("")))); // spacing after table
        }

        // -------------------- OMML builder (simple LaTeX) --------------------

        private static M.OfficeMath BuildOfficeMath(string latex)
        {
            latex = (latex ?? string.Empty)
                .Replace(@"\times", "×")
                .Trim();

            var tokens = new TokenStream(latex);
            var math = new M.OfficeMath();

            foreach (var el in ParseSequence(tokens))
                math.Append(el);

            return math;
        }

        private static List<OpenXmlElement> ParseSequence(TokenStream ts, char? stopAt = null)
        {
            var list = new List<OpenXmlElement>();

            while (!ts.End)
            {
                if (stopAt.HasValue && ts.Peek() == stopAt.Value)
                    break;

                if (ts.StartsWith(@"\frac"))
                {
                    ts.Consume(@"\frac");
                    var num = ParseGroup(ts);
                    var den = ParseGroup(ts);
                    list.Add(BuildFrac(num, den));
                    continue;
                }

                if (ts.StartsWith(@"\sqrt"))
                {
                    ts.Consume(@"\sqrt");
                    var rad = ParseGroup(ts);
                    list.Add(BuildSqrt(rad));
                    continue;
                }

                var atom = ParseAtom(ts);

                while (!ts.End && (ts.Peek() == '^' || ts.Peek() == '_'))
                {
                    var op = ts.Read();
                    var arg = ParseScript(ts);
                    atom = (op == '^') ? BuildSup(atom, arg) : BuildSub(atom, arg);
                }

                list.Add(atom);
            }

            return list;
        }

        private static OpenXmlElement ParseAtom(TokenStream ts)
        {
            char c = ts.Peek();

            if (c == '{')
            {
                var grp = ParseGroup(ts);
                return BuildRun(string.Concat(grpText(grp)));
            }

            ts.Read();
            return BuildRun(c.ToString());

            static IEnumerable<string> grpText(List<OpenXmlElement> els)
            {
                foreach (var el in els)
                {
                    if (el is M.Run r)
                    {
                        foreach (var t in r.Elements<M.Text>())
                            yield return t.Text;
                    }
                    else
                    {
                        yield return "";
                    }
                }
            }
        }

        private static List<OpenXmlElement> ParseGroup(TokenStream ts)
        {
            if (ts.End || ts.Peek() != '{')
                return new List<OpenXmlElement> { BuildRun("") };

            ts.Read();
            var inner = ParseSequence(ts, stopAt: '}');
            if (!ts.End && ts.Peek() == '}') ts.Read();
            return inner;
        }

        private static List<OpenXmlElement> ParseScript(TokenStream ts)
        {
            if (!ts.End && ts.Peek() == '{')
                return ParseGroup(ts);

            if (ts.End) return new List<OpenXmlElement> { BuildRun("") };
            var c = ts.Read();
            return new List<OpenXmlElement> { BuildRun(c.ToString()) };
        }

        private static M.Run BuildRun(string text) => new M.Run(new M.Text(text ?? string.Empty));

        private static M.Fraction BuildFrac(List<OpenXmlElement> num, List<OpenXmlElement> den)
        {
            var f = new M.Fraction();

            var n = new M.Numerator();
            foreach (var el in num) n.Append(el.CloneNode(true));

            var d = new M.Denominator();
            foreach (var el in den) d.Append(el.CloneNode(true));

            f.Append(n);
            f.Append(d);
            return f;
        }

        private static M.Radical BuildSqrt(List<OpenXmlElement> radicand)
        {
            var r = new M.Radical();
            r.Append(new M.Degree(new M.Run(new M.Text(""))));

            var b = new M.Base();
            foreach (var el in radicand) b.Append(el.CloneNode(true));
            r.Append(b);

            return r;
        }

        private static M.Superscript BuildSup(OpenXmlElement @base, List<OpenXmlElement> supArg)
        {
            var s = new M.Superscript();

            var b = new M.Base();
            b.Append(@base.CloneNode(true));

            var a = new M.SuperArgument();
            foreach (var el in supArg) a.Append(el.CloneNode(true));

            s.Append(b);
            s.Append(a);
            return s;
        }

        private static M.Subscript BuildSub(OpenXmlElement @base, List<OpenXmlElement> subArg)
        {
            var s = new M.Subscript();

            var b = new M.Base();
            b.Append(@base.CloneNode(true));

            var a = new M.SubArgument();
            foreach (var el in subArg) a.Append(el.CloneNode(true));

            s.Append(b);
            s.Append(a);
            return s;
        }

        private class TokenStream
        {
            private readonly string _s;
            private int _i;

            public TokenStream(string s)
            {
                _s = s ?? "";
                _i = 0;
            }

            public bool End => _i >= _s.Length;
            public char Peek() => End ? '\0' : _s[_i];
            public char Read() => End ? '\0' : _s[_i++];

            public bool StartsWith(string t)
            {
                if (t == null) return false;
                if (_i + t.Length > _s.Length) return false;
                return string.Compare(_s, _i, t, 0, t.Length, StringComparison.Ordinal) == 0;
            }

            public void Consume(string t)
            {
                if (StartsWith(t)) _i += t.Length;
            }
        }
    }
}
