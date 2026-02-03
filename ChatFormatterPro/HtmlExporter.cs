using System.IO;
using System.Text;

namespace ChatFormatterPro.Exporters
{
    public static class HtmlExporter
    {
        public static void Export(string content, string filePath, string title)
        {
            content ??= "";

            var sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset='UTF-8'>");
            sb.AppendLine($"<title>{title}</title>");

            // ✅ MathJax for LaTeX rendering
            sb.AppendLine("<script src='https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js'></script>");

            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: Arial; font-size: 16px; padding: 30px; }");
            sb.AppendLine("</style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");

            foreach (var line in content.Replace("\r\n", "\n").Split('\n'))
            {
                sb.AppendLine($"<p>{System.Net.WebUtility.HtmlEncode(line)}</p>");
            }

            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }
    }
}
