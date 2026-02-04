using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ChatFormatterPro.Exporters
{
    public static class HtmlExporter
    {
        // ✅ convert **bold** to <strong>
        private static readonly Regex BoldRegex = new Regex(@"\*\*(.+?)\*\*", RegexOptions.Singleline);

        public static void Export(string content, string filePath, string title)
        {
            content ??= "";

            var sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset='UTF-8'>");
            sb.AppendLine($"<title>{HtmlEncode(title ?? "")}</title>");

            // ✅ MathJax for LaTeX rendering
            sb.AppendLine("<script src='https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js'></script>");

            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: Arial; font-size: 16px; padding: 30px; }");
            sb.AppendLine("p { margin: 8px 0; }");
            sb.AppendLine("table.cfp { border-collapse: collapse; width: 100%; margin: 12px 0; }");
            sb.AppendLine("table.cfp th, table.cfp td { border: 1px solid #333; padding: 8px; text-align: left; vertical-align: top; }");
            sb.AppendLine("table.cfp th { font-weight: bold; }");
            sb.AppendLine("</style>");

            sb.AppendLine("</head>");
            sb.AppendLine("<body>");

            var lines = content.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i] ?? "";

                // Blank line
                if (string.IsNullOrWhiteSpace(line))
                {
                    sb.AppendLine("<p></p>");
                    continue;
                }

                // ✅ TABLE detection (pipe-table)
                if (IsPipeTableLine(line))
                {
                    var pipeLines = new List<string>();

                    // collect contiguous table lines
                    while (i < lines.Length && IsPipeTableLine(lines[i]))
                    {
                        pipeLines.Add(lines[i]);
                        i++;
                    }
                    i--; // step back

                    // Convert to rows (skip separator if present at index 1)
                    var rows = new List<List<string>>();
                    for (int k = 0; k < pipeLines.Count; k++)
                    {
                        if (k == 1 && IsMarkdownSeparatorLine(pipeLines[k]))
                            continue;

                        rows.Add(SplitPipeRow(pipeLines[k]));
                    }

                    AppendHtmlTable(sb, rows);
                    continue;
                }

                // Normal paragraph (supports **bold**)
                sb.AppendLine($"<p>{ConvertBoldToHtml(HtmlEncode(line))}</p>");
            }

            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            // ✅ AUTO OPEN AFTER EXPORT (you already have FileOpener in your project)
            ChatFormatterPro.FileOpener.Open(filePath);
        }

        // -------------------- TABLE SUPPORT --------------------

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
            {
                if (ch != '|' && ch != '-' && ch != ':' && ch != ' ' && ch != '\t')
                    return false;
            }
            return line.Contains("-");
        }

        private static List<string> SplitPipeRow(string line)
        {
            line = (line ?? "").Trim();
            if (line.StartsWith("|")) line = line.Substring(1);
            if (line.EndsWith("|")) line = line.Substring(0, line.Length - 1);

            return line.Split('|')
                       .Select(x => x.Trim())
                       .ToList();
        }

        private static void AppendHtmlTable(StringBuilder sb, List<List<string>> rows)
        {
            if (rows == null || rows.Count == 0) return;

            int colCount = rows.Max(r => r.Count);
            foreach (var r in rows)
                while (r.Count < colCount) r.Add("");

            sb.AppendLine("<table class='cfp'>");

            // Header row
            sb.AppendLine("<tr>");
            for (int c = 0; c < colCount; c++)
            {
                var cell = ConvertBoldToHtml(HtmlEncode(rows[0][c]));
                sb.AppendLine($"<th>{cell}</th>");
            }
            sb.AppendLine("</tr>");

            // Body rows
            for (int r = 1; r < rows.Count; r++)
            {
                sb.AppendLine("<tr>");
                for (int c = 0; c < colCount; c++)
                {
                    var cell = ConvertBoldToHtml(HtmlEncode(rows[r][c]));
                    sb.AppendLine($"<td>{cell}</td>");
                }
                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</table>");
        }

        // -------------------- HELPERS --------------------

        private static string HtmlEncode(string s)
            => System.Net.WebUtility.HtmlEncode(s ?? "");

        private static string ConvertBoldToHtml(string encodedText)
        {
            // encodedText is already HTML-encoded, so bold markers are safe to replace
            return BoldRegex.Replace(encodedText, "<strong>$1</strong>");
        }
    }
}
