using ChatFormatterPro;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace ChatFormatterPro.Exporters
{
    public static class PdfExporter
    {
        public static void Export(string content, string filePath, string title)
        {
            content ??= "";

            QuestPDF.Settings.License = LicenseType.Community;

            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(40);
                    page.DefaultTextStyle(x => x.FontSize(12));

                    page.Content().Column(col =>
                    {
                        // Title
                        if (!string.IsNullOrWhiteSpace(title))
                        {
                            col.Item().Text(title)
                                .FontSize(20)
                                .Bold()
                                .AlignCenter();

                            col.Item().PaddingBottom(15);
                        }

                        // Content
                        foreach (var line in content.Replace("\r\n", "\n").Split('\n'))
                        {
                            col.Item().Text(line);
                        }
                    });
                });
            })
            .GeneratePdf(filePath);

            // ✅ AUTO OPEN AFTER EXPORT
            FileOpener.Open(filePath);
        }
    }
}
