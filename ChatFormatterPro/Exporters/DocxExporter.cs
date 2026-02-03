using System;
using System.Collections.Generic;
using System.Linq; // ✅ REQUIRED for .Any()
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using M = DocumentFormat.OpenXml.Math;
using Markdig;

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

            foreach (var rawLine in lines)
            {
                var line = (rawLine ?? string.Empty);

                // Blank line
                if (string.IsNullOrWhiteSpace(line))
                {
                    body.Append(new Paragraph(new Run(new Text(""))));
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

                // Heading 2: ## Title
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

                    // 🔥 IMPORTANT: reuse bold parser
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

            mainPart.Document.Save();
            OpenFileAfterExport(filePath);
        }

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

        // ✅ FIXED: OpenXML 3.x uses NumberingInstance (not Num)
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

        private static void OpenFileAfterExport(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath)) return;
                if (!System.IO.File.Exists(filePath)) return;

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            catch
            {
                // optional: ignore
            }
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
