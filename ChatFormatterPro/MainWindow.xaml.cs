using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Microsoft.Win32;

using ChatFormatterPro.Exporters;

using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;

// ✅ Needed for .Count() / .Select() in PDF table code
using System.Linq;
using System.Collections.Generic;

namespace ChatFormatterPro
{
    public partial class MainWindow : Window
    {
        // ✅ Convert emoji tick/box to PDF-safe symbols that fonts can render
        private static string PdfSafeSymbols(string s)
        {
            return (s ?? "")
                .Replace("✅", "✔")
                .Replace("✔️", "✔")
                .Replace("✔", "✔")
                .Replace("☑", "✔")
                .Replace("☐", "");   // or use "□" if you want empty box visible
        }

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
            MessageBox.Show("TEST: I am running the latest code");

            string input = InputTextBox.Text;
            InputTextBox.Text = MathLatexHelper.ConvertPowersToLatex(input);
        }

        private void RewriteLatex_Click(object sender, RoutedEventArgs e)
        {
            string input = InputTextBox.Text;
            InputTextBox.Text = MathLatexHelper.ConvertPowersToLatex(input);
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

        #region Export HTML ✅

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
                    DefaultExt = ".csv",
                    FileName = "ChatFormatterPro_Export.csv"
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

                // ✅ AUTO OPEN CSV (opens in Excel / default app)
                FileOpener.Open(dlg.FileName);

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
                    DefaultExt = ".txt",
                    FileName = "ChatFormatterPro_Export.txt"
                };

                if (dlg.ShowDialog() != true) return;

                File.WriteAllText(dlg.FileName, InputTextBox.Text ?? string.Empty, Encoding.UTF8);

                // ✅ AUTO OPEN TXT (opens in Notepad / default app)
                FileOpener.Open(dlg.FileName);

                MessageBox.Show("TXT saved successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving TXT: " + ex.Message);
            }
        }

        #endregion

        #region Export PDF (MigraDoc) ✅ NOW SUPPORTS TABLES + TICK SYMBOL

        private void SaveToPdf_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new SaveFileDialog
                {
                    Filter = "PDF File (*.pdf)|*.pdf",
                    DefaultExt = ".pdf",
                    FileName = "ChatFormatterPro_Export.pdf"
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

                // ✅ Local helpers for pipe-table parsing
                bool IsPipeRow(string s)
                {
                    if (string.IsNullOrWhiteSpace(s)) return false;
                    s = s.Trim();
                    return s.Contains("|") && s.Count(ch => ch == '|') >= 2;
                }

                bool IsSeparator(string s)
                {
                    if (string.IsNullOrWhiteSpace(s)) return false;
                    s = s.Trim();
                    foreach (char ch in s)
                    {
                        if (ch != '|' && ch != '-' && ch != ':' && ch != ' ' && ch != '\t')
                            return false;
                    }
                    return s.Contains("-");
                }

                string[] SplitCells(string s)
                {
                    s = (s ?? "").Trim();
                    if (s.StartsWith("|")) s = s.Substring(1);
                    if (s.EndsWith("|")) s = s.Substring(0, s.Length - 1);
                    return s.Split('|').Select(x => x.Trim()).ToArray();
                }

                int idx = 0;
                while (idx < lines.Length)
                {
                    string line = lines[idx] ?? "";

                    // blank line
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        section.AddParagraph();
                        idx++;
                        continue;
                    }

                    // ✅ TABLE block
                    if (IsPipeRow(line))
                    {
                        var pipeLines = new List<string>();

                        while (idx < lines.Length && IsPipeRow(lines[idx]))
                        {
                            pipeLines.Add(lines[idx]);
                            idx++;
                        }

                        // determine number of columns
                        int colCount = 0;
                        for (int k = 0; k < pipeLines.Count; k++)
                        {
                            if (k == 1 && IsSeparator(pipeLines[k])) continue;
                            colCount = Math.Max(colCount, SplitCells(pipeLines[k]).Length);
                        }
                        if (colCount == 0) colCount = 1;

                        // create MigraDoc table
                        var table = section.AddTable();
                        table.Borders.Width = 0.75;

                        // ✅ IMPORTANT: set a font that can render ✔
                        table.Format.Font.Name = "Segoe UI Symbol";
                        table.Format.Font.Size = 11;

                        // fit page width roughly (16 cm usable area with 2 cm margins)
                        for (int c = 0; c < colCount; c++)
                            table.AddColumn(Unit.FromCentimeter(16.0 / colCount));

                        bool headerDone = false;

                        for (int k = 0; k < pipeLines.Count; k++)
                        {
                            if (k == 1 && IsSeparator(pipeLines[k]))
                                continue;

                            var cells = SplitCells(pipeLines[k]);

                            var row = table.AddRow();
                            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;

                            if (!headerDone)
                            {
                                row.Format.Font.Bold = true;
                                headerDone = true;
                            }

                            for (int c = 0; c < colCount; c++)
                            {
                                string cellText = (c < cells.Length) ? cells[c] : "";
                                cellText = PdfSafeSymbols(cellText);

                                var para = row.Cells[c].AddParagraph(cellText);
                                para.Format.Font.Name = "Segoe UI Symbol"; // ✅ ensure per-cell too
                                para.Format.Font.Size = 11;
                            }
                        }

                        section.AddParagraph(); // spacing after table
                        continue;
                    }

                    // normal text line
                    Paragraph p = section.AddParagraph();
                    p.Format.Font.Name = "Segoe UI Symbol";  // ✅ supports ✔
                    p.Format.Font.Size = 11;
                    p.AddText(PdfSafeSymbols(line));

                    idx++;
                }

                var renderer = new PdfDocumentRenderer(unicode: true)
                {
                    Document = pdfDoc
                };
                renderer.RenderDocument();
                renderer.Save(dlg.FileName);

                // ✅ AUTO OPEN PDF (important: after Save)
                FileOpener.Open(dlg.FileName);

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
