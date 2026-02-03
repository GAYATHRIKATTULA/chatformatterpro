using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Microsoft.Win32;

using ChatFormatterPro.Exporters;

using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;

namespace ChatFormatterPro
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        #region Clipboard

        private void PasteReplace_Click(object sender, RoutedEventArgs e)
        {
            if (!Clipboard.ContainsText()) return;
            InputTextBox.Text = Clipboard.GetText();
            ApplyTextCleaning();
        }

        private void PasteInsert_Click(object sender, RoutedEventArgs e)
        {
            if (!Clipboard.ContainsText()) return;
            InputTextBox.Text += Clipboard.GetText();
            ApplyTextCleaning();
        }

        private void Cut_Click(object sender, RoutedEventArgs e) => InputTextBox.Cut();
        private void Copy_Click(object sender, RoutedEventArgs e) => InputTextBox.Copy();

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            InputTextBox.Focus();
            InputTextBox.SelectAll();
        }

        private void Undo_Click(object sender, RoutedEventArgs e)
        {
            if (InputTextBox.CanUndo) InputTextBox.Undo();
        }

        private void Clear_Click(object sender, RoutedEventArgs e) => InputTextBox.Clear();

        #endregion

        #region Text Cleaning

        private string ProcessText(string input)
        {
            string text = input ?? string.Empty;

            if (RemoveLinesCheckBox.IsChecked == true)
                text = Regex.Replace(text, @"^[-=]{3,}$", "", RegexOptions.Multiline);

            if (RemoveHeadingEmojiCheckBox.IsChecked == true)
            {
                text = Regex.Replace(
                    text,
                    @"^[\s]*(👉|🔥|⭐|❗|⚠️|➡️|✔|✖|💡|📌)+",
                    "",
                    RegexOptions.Multiline);
            }

            if (NormalBulletsRadio.IsChecked == true)
            {
                text = Regex.Replace(text, @"^[\s]*[•●▪️▶️➤➔➜➡️]", "•", RegexOptions.Multiline);
            }
            else if (EmojiBulletsRadio.IsChecked == true)
            {
                text = Regex.Replace(text, @"^[\s]*[-*•]", "👉", RegexOptions.Multiline);
            }
            else if (EmojiOnlyRadio.IsChecked == true)
            {
                text = Regex.Replace(text, @"^[\s]*[-*•▶️➤➜➔➡️]", "⭐", RegexOptions.Multiline);
            }

            if (ForceMathCheckBox.IsChecked == true)
            {
                text = Regex.Replace(text, @"\\\((.*?)\\\)", m => m.Groups[1].Value);
                text = Regex.Replace(text, @"\\\[(.*?)\\\]", m => m.Groups[1].Value);
            }

            return text;
        }

        private void ApplyTextCleaning()
        {
            InputTextBox.Text = ProcessText(InputTextBox.Text);
        }

        #endregion

        #region ChatGPT helpers

        private void RenderLatex_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("LaTeX rendering will be added soon.");
        }

        private void RewriteLatex_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("LaTeX rewriting will be added soon.");
        }

        #endregion

        #region Export menu

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            ExportButton.ContextMenu = ExportMenu;
            ExportMenu.PlacementTarget = ExportButton;
            ExportMenu.IsOpen = true;
        }

        #endregion

        #region Export DOCX

        private void ExportToDocx_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new SaveFileDialog
                {
                    Filter = "Word Document (*.docx)|*.docx",
                    DefaultExt = ".docx",
                    FileName = "ChatFormatterPro_Export.docx"
                };

                if (dlg.ShowDialog() != true) return;

                DocxExporter.Export(InputTextBox.Text, dlg.FileName, "Exported Content");

                MessageBox.Show("DOCX saved successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving DOCX: " + ex.Message);
            }
        }

        #endregion

        #region Export HTML  ✅ (MathJax will render LaTeX in browser)

        private void ExportHtml_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new SaveFileDialog
                {
                    Filter = "HTML File (*.html)|*.html",
                    DefaultExt = ".html",
                    FileName = "ChatFormatterPro_Export.html"
                };

                if (dlg.ShowDialog() != true) return;

                HtmlExporter.Export(InputTextBox.Text, dlg.FileName, "Exported Content");

                MessageBox.Show("HTML saved successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving HTML: " + ex.Message);
            }
        }

        #endregion

        #region Export CSV (Excel)

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new SaveFileDialog
                {
                    Filter = "CSV File (*.csv)|*.csv",
                    DefaultExt = ".csv"
                };

                if (dlg.ShowDialog() != true) return;

                string[] lines = (InputTextBox.Text ?? string.Empty)
                    .Replace("\r\n", "\n")
                    .Split('\n');

                var sb = new StringBuilder();

                foreach (var line in lines)
                {
                    string cell = (line ?? string.Empty).Replace("\"", "\"\"");
                    sb.AppendLine($"\"{cell}\"");
                }

                File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show("CSV saved successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving CSV: " + ex.Message);
            }
        }

        #endregion

        #region Export TXT

        private void ExportPlainText_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new SaveFileDialog
                {
                    Filter = "Text File (*.txt)|*.txt",
                    DefaultExt = ".txt"
                };

                if (dlg.ShowDialog() != true) return;

                File.WriteAllText(dlg.FileName, InputTextBox.Text ?? string.Empty, Encoding.UTF8);
                MessageBox.Show("TXT saved successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving TXT: " + ex.Message);
            }
        }

        #endregion

        #region Export PDF (MigraDoc) ⚠️ LaTeX will NOT render here

        private void SaveToPdf_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new SaveFileDialog
                {
                    Filter = "PDF File (*.pdf)|*.pdf",
                    DefaultExt = ".pdf"
                };

                if (dlg.ShowDialog() != true) return;

                var pdfDoc = new Document
                {
                    Info = { Title = "ChatFormatter Export" }
                };

                Section section = pdfDoc.AddSection();
                section.PageSetup.LeftMargin = Unit.FromCentimeter(2);
                section.PageSetup.RightMargin = Unit.FromCentimeter(2);
                section.PageSetup.TopMargin = Unit.FromCentimeter(2);
                section.PageSetup.BottomMargin = Unit.FromCentimeter(2);

                string[] lines = (InputTextBox.Text ?? string.Empty)
                    .Replace("\r\n", "\n")
                    .Split('\n');

                foreach (string line in lines)
                {
                    Paragraph p = section.AddParagraph();
                    p.Format.Font.Name = "Calibri";
                    p.Format.Font.Size = 11;
                    p.AddText(line ?? string.Empty);
                }

                var renderer = new PdfDocumentRenderer(unicode: true)
                {
                    Document = pdfDoc
                };
                renderer.RenderDocument();
                renderer.Save(dlg.FileName);

                MessageBox.Show("PDF saved successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("PDF Error: " + ex.Message);
            }
        }

        #endregion
    }
}
