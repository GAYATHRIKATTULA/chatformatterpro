#nullable enable

using ChatFormatterPro.Exporters;
using Microsoft.Win32;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace ChatFormatterPro
{
    public partial class MainWindow : Window
    {
        // ---------------- IMAGE INSERT (A + B) ----------------
        private const string ImgTokenPrefix = "{{IMG:";
        private const string ImgTokenSuffix = "}}";

        private bool _mathPasteHooked;

        private static string EnsureImagesFolder()
        {
            string folder = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "ChatFormatterPro_Images"
            );
            Directory.CreateDirectory(folder);
            return folder;
        }

        // ✅ Detect token line: {{IMG:C:\path\file.png}}
        private static bool TryGetImgPathFromLine(string line, out string? path)
        {
            path = null;

            if (string.IsNullOrWhiteSpace(line)) return false;

            string t = line.Trim();
            if (!t.StartsWith(ImgTokenPrefix) || !t.EndsWith(ImgTokenSuffix))
                return false;

            string inner = t.Substring(
                ImgTokenPrefix.Length,
                t.Length - ImgTokenPrefix.Length - ImgTokenSuffix.Length
            ).Trim();

            inner = inner.Trim('"');

            if (!File.Exists(inner)) return false;
            path = inner;
            return true;
        }

        private void PasteImage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!Clipboard.ContainsImage())
                {
                    MessageBox.Show("Clipboard has no image.");
                    return;
                }

                var img = Clipboard.GetImage();
                if (img == null)
                {
                    MessageBox.Show("Clipboard image not found.");
                    return;
                }

                string folder = EnsureImagesFolder();
                string file = System.IO.Path.Combine(
                    folder,
                    $"img_{DateTime.Now:yyyyMMdd_HHmmss}.png"
                );

                using (var fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(img));
                    encoder.Save(fs);
                }

                InsertImageTokenIntoTextbox(file);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Paste Image failed: " + ex.Message);
            }
        }

        private void AddImageFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Images|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|All files|*.*",
                    Multiselect = false
                };

                if (dlg.ShowDialog() != true) return;

                string src = dlg.FileName;
                string folder = EnsureImagesFolder();

                string ext = System.IO.Path.GetExtension(src);
                string dest = System.IO.Path.Combine(
                    folder,
                    $"img_{DateTime.Now:yyyyMMdd_HHmmss}{ext}"
                );

                File.Copy(src, dest, overwrite: true);

                InsertImageTokenIntoTextbox(dest);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Add Image failed: " + ex.Message);
            }
        }

        private void InsertImageTokenIntoTextbox(string imagePath)
        {
            string tokenLine = $"{ImgTokenPrefix}{imagePath}{ImgTokenSuffix}";

            int caret = InputTextBox.CaretIndex;
            string text = InputTextBox.Text ?? "";

            string insert =
                (caret > 0 && caret < text.Length && text[caret - 1] != '\n')
                ? "\r\n"
                : "";

            insert += tokenLine + "\r\n";

            InputTextBox.Text = text.Insert(caret, insert);
            InputTextBox.CaretIndex = caret + insert.Length;
            InputTextBox.Focus();
        }

        public MainWindow()
        {
            InitializeComponent();
            HookMathPasteHandlerOnce();
        }

        private void HookMathPasteHandlerOnce()
        {
            if (_mathPasteHooked) return;

            DataObject.AddPastingHandler(InputTextBox, OnEditorPaste);
            _mathPasteHooked = true;
        }

        private void OnEditorPaste(object sender, DataObjectPastingEventArgs e)
        {
            if (!e.SourceDataObject.GetDataPresent(DataFormats.UnicodeText, true))
                return;

            string? pastedText = e.SourceDataObject.GetData(DataFormats.UnicodeText) as string;
            if (string.IsNullOrWhiteSpace(pastedText))
                return;

            pastedText = pastedText.Replace("\r\n", "\n").Replace("\r", "\n");

            string normalized = MathLatexHelper.NormalizeInformalMath(pastedText);

            e.CancelCommand();
            InsertTextAtCaret(InputTextBox, normalized);
        }

        private void NormalizeEditorContent()
        {
            string current = InputTextBox.Text ?? "";
            if (string.IsNullOrWhiteSpace(current))
                return;

            string normalized = MathLatexHelper.NormalizeInformalMath(current);
            if (normalized == current)
                return;

            int caretPos = InputTextBox.CaretIndex;
            InputTextBox.Text = normalized;
            InputTextBox.CaretIndex = Math.Min(caretPos, normalized.Length);
        }

        private static void InsertTextAtCaret(TextBox textBox, string text)
        {
            if (textBox == null || text == null)
                return;

            int start = textBox.SelectionStart;
            int length = textBox.SelectionLength;

            string before = textBox.Text.Substring(0, start);
            string after = textBox.Text.Substring(start + length);

            textBox.Text = before + text + after;
            textBox.CaretIndex = start + text.Length;
        }

        #region PDF helpers

        private static string PdfSafeSymbols(string s)
        {
            return (s ?? "")
                .Replace("✅", "✔")
                .Replace("✔️", "✔")
                .Replace("✔", "✔")
                .Replace("☑", "✔")
                .Replace("☐", "");
        }

        #endregion

        #region Superscript helpers

        private static string NormalizeUnicodeSuperscripts(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            var map = new Dictionary<char, char>
            {
                ['⁰'] = '0',
                ['¹'] = '1',
                ['²'] = '2',
                ['³'] = '3',
                ['⁴'] = '4',
                ['⁵'] = '5',
                ['⁶'] = '6',
                ['⁷'] = '7',
                ['⁸'] = '8',
                ['⁹'] = '9',
                ['⁻'] = '-',
                ['⁺'] = '+'
            };

            var sb = new StringBuilder();

            for (int i = 0; i < text.Length; i++)
            {
                if (map.ContainsKey(text[i]))
                {
                    sb.Append("^");
                    while (i < text.Length && map.ContainsKey(text[i]))
                    {
                        sb.Append(map[text[i]]);
                        i++;
                    }
                    i--;
                }
                else
                {
                    sb.Append(text[i]);
                }
            }

            return sb.ToString();
        }

        private static string ConvertCaretPowersToUnicodeSuperscripts(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            string ToSuper(string s)
            {
                var sb = new StringBuilder();
                foreach (char c in s)
                {
                    sb.Append(c switch
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
                        _ => c
                    });
                }
                return sb.ToString();
            }

            return Regex.Replace(
                text,
                @"\^\{?(?<exp>[+-]?\d+)\}?",
                m => ToSuper(m.Groups["exp"].Value)
            );
        }

        #endregion

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

        #region Cleaning pipeline (UI display)

        private string CleanTextForDisplay(string input)
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
                    RegexOptions.Multiline
                );
            }

            if (NormalBulletsRadio.IsChecked == true)
            {
                text = Regex.Replace(
                    text,
                    @"^[\s]*[•●▪️▶️➤➔➜➡️]",
                    "•",
                    RegexOptions.Multiline
                );
            }
            else if (EmojiBulletsRadio.IsChecked == true)
            {
                text = Regex.Replace(text, @"^[\s]*[-*•]", "👉", RegexOptions.Multiline);
            }
            else if (EmojiOnlyRadio.IsChecked == true)
            {
                text = Regex.Replace(
                    text,
                    @"^[\s]*[-*•▶️➤➜➔➡️]",
                    "⭐",
                    RegexOptions.Multiline
                );
            }

            if (ForceMathCheckBox.IsChecked == true)
            {
                text = Regex.Replace(text, @"\\\((.*?)\\\)", m => m.Groups[1].Value);
                text = Regex.Replace(text, @"\\\[(.*?)\\\]", m => m.Groups[1].Value);
            }

            return text;
        }

        private string ProcessTextForDisplay(string input)
        {
            string text = CleanTextForDisplay(input);

            text = NormalizeUnicodeSuperscripts(text);
            text = ConvertCaretPowersToUnicodeSuperscripts(text);

            return text;
        }

        private void ApplyTextCleaning()
        {
            InputTextBox.Text = ProcessTextForDisplay(InputTextBox.Text);
        }

        #endregion

        #region Export pipeline (Word Equation mode)

        private static bool LooksLikeMath(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();

            if (s.Contains(@"\(") || s.Contains(@"\)") || s.Contains(@"\[") || s.Contains(@"\]"))
                return true;

            if (Regex.IsMatch(
                s,
                @"\\(frac|sqrt|times|div|cdot|pi|theta|alpha|beta|gamma|left|right|sum|int)\b"
            ))
                return true;

            if (s.Contains("^") || s.Contains("_") || s.Contains("{") || s.Contains("}"))
                return true;

            bool hasDigit = s.Any(char.IsDigit);
            bool hasOps = s.IndexOfAny(new[] { '=', '+', '-', '×', '*', '÷', '/', '(', ')', '[', ']' }) >= 0;
            int letterCount = s.Count(char.IsLetter);

            if (hasDigit && hasOps && letterCount <= 6) return true;

            return false;
        }

        private static string WrapStandaloneMathLines(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            var lines = text.Replace("\r\n", "\n").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string raw = lines[i] ?? "";
                string t = raw.Trim();

                if (string.IsNullOrWhiteSpace(t)) continue;

                if (t.StartsWith(@"\(") && t.EndsWith(@"\)")) continue;
                if (t.StartsWith(@"\[") && t.EndsWith(@"\]")) continue;

                if (LooksLikeMath(t))
                {
                    lines[i] = $@"\({t}\)";
                }
            }

            return string.Join("\n", lines).Replace("\n", "\r\n");
        }

        private static string AutoWrapInlineMathSegments(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            var rx = new Regex(
                @"(?<!\\)\b\\(frac|sqrt|times|div|cdot|pi|theta|alpha|beta|gamma|left|right)\b[^\s]*" +
                @"|" +
                @"(?<!\\)\([^\(\)\r\n]{0,60}[0-9a-zA-Z\}\{_\^\+\-\*/×÷=][^\(\)\r\n]{0,60}\)" +
                @"|" +
                @"(?<!\\)\b[a-zA-Z0-9]+\s*(\^|_)\s*\{?[+\-]?\d+[a-zA-Z0-9]*\}?" +
                @"|" +
                @"(?<!\\)\b\d+\s*(/|×|÷|\*)\s*\d+\b",
                RegexOptions.Compiled
            );

            var lines = text.Replace("\r\n", "\n").Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (string.IsNullOrWhiteSpace(line)) continue;

                string t = line.Trim();

                if ((t.StartsWith(@"\(") && t.EndsWith(@"\)")) || (t.StartsWith(@"\[") && t.EndsWith(@"\]")))
                    continue;

                lines[i] = rx.Replace(line, m =>
                {
                    string s = m.Value;

                    if (s.Contains(@"\(") || s.Contains(@"\)") || s.Contains(@"\[") || s.Contains(@"\]"))
                        return s;

                    if (s.Contains("http") || s.Contains(@":\") || s.Contains(@"\\"))
                        return s;

                    return $@"\({s}\)";
                });
            }

            return string.Join("\n", lines).Replace("\n", "\r\n");
        }

        private string PrepareForEquationExport(string input)
        {
            bool forceMath = ForceMathCheckBox.IsChecked == true;

            string text = CleanTextForDisplay(input);

            text = NormalizeUnicodeSuperscripts(text);

            text = Regex.Replace(text, @"(?<![\w\)])-(\d+|[A-Za-z]+)\^", @"(-$1)^");
            text = Regex.Replace(text, @"([A-Za-z0-9\)\]](?:\s*[+\-*/]\s*[A-Za-z0-9\(\)\[\]]+)+)\^", @"($1)^");
            text = Regex.Replace(text, @"(\([^\)]+\)|[A-Za-z0-9]+)\^(-?\d+)\^(-?\d+)", @"($1^$2)^$3");
            text = Regex.Replace(text, @"\^-(\d+)", @"^{- $1}".Replace(" ", ""));
            text = MathLatexHelper.ConvertPowersToLatex(text);

            text = WrapStandaloneMathLines(text);

            if (forceMath)
            {
                text = AutoWrapInlineMathSegments(text);
            }

            return text;
        }

        #endregion

        private static bool LooksLikeStandaloneMath(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;

            if (s.Contains("=") && s.Contains("^"))
                return true;

            if (s.Contains("^"))
            {
                if (!Regex.IsMatch(s, @"[A-Za-z]{4,}"))
                    return true;
            }

            if (s.Contains(@"\frac") || s.Contains(@"\sqrt"))
                return true;

            return false;
        }

        #region ChatGPT helpers

        private void RenderLatex_Click(object sender, RoutedEventArgs e)
        {
            NormalizeEditorContent();
        }

        private void RewriteLatex_Click(object sender, RoutedEventArgs e)
        {
            string input = InputTextBox.Text ?? "";
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
                    FileName = $"ChatFormatterPro_Export_{DateTime.Now:yyyyMMdd_HHmmss}.docx"
                };

                if (dlg.ShowDialog() != true) return;

                string rawInput = InputTextBox.Text ?? "";

                if (ExportAsWordEquationsCheckBox.IsChecked != true)
                {
                    string content = ProcessTextForDisplay(rawInput);
                    DocxExporter.Export(content, dlg.FileName, "Exported Content");
                    MessageBox.Show("DOCX saved successfully! (Superscript mode)");
                    return;
                }

                string eqContent = PrepareForEquationExport(rawInput);

                DocxExporter.Export(eqContent, dlg.FileName, "Exported Content");
                MessageBox.Show("DOCX saved successfully! (Word Equation mode)");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving DOCX: " + ex.Message);
            }
        }

        #endregion

        #region Export HTML

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

                string text = ProcessTextForDisplay(InputTextBox.Text ?? "");
                var lines = text.Replace("\r\n", "\n").Split('\n');

                for (int i = 0; i < lines.Length; i++)
                {
                    if (TryGetImgPathFromLine(lines[i], out var imgPath))
                    {
                        if (string.IsNullOrWhiteSpace(imgPath) || !File.Exists(imgPath))
                            continue;

                        string uri = new Uri(imgPath).AbsoluteUri;
                        lines[i] = $"<img src=\"{uri}\" style=\"max-width:100%; height:auto;\" />";
                    }
                }

                string htmlReady = string.Join("\n", lines).Replace("\n", "\r\n");

                HtmlExporter.Export(htmlReady, dlg.FileName, "Exported Content");
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

                string text = ProcessTextForDisplay(InputTextBox.Text ?? "");
                string[] lines = text.Replace("\r\n", "\n").Split('\n');

                var sb = new StringBuilder();
                foreach (var line in lines)
                {
                    string cell = (line ?? string.Empty).Replace("\"", "\"\"");
                    sb.AppendLine($"\"{cell}\"");
                }

                File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);

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

                File.WriteAllText(
                    dlg.FileName,
                    ProcessTextForDisplay(InputTextBox.Text ?? ""),
                    Encoding.UTF8
                );

                FileOpener.Open(dlg.FileName);
                MessageBox.Show("TXT saved successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving TXT: " + ex.Message);
            }
        }

        #endregion

        #region Export PDF (MigraDoc)

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

                string[] lines = ProcessTextForDisplay(InputTextBox.Text ?? "")
                    .Replace("\r\n", "\n")
                    .Split('\n');

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

                    if (string.IsNullOrWhiteSpace(line))
                    {
                        section.AddParagraph();
                        idx++;
                        continue;
                    }

                    if (TryGetImgPathFromLine(line, out var imgPath))
                    {
                        if (!string.IsNullOrWhiteSpace(imgPath) && File.Exists(imgPath))
                        {
                            var img = section.AddImage(imgPath);
                            img.LockAspectRatio = true;
                            img.Width = Unit.FromCentimeter(14);
                            section.AddParagraph();
                        }
                        idx++;
                        continue;
                    }

                    if (IsPipeRow(line))
                    {
                        var pipeLines = new List<string>();

                        while (idx < lines.Length && IsPipeRow(lines[idx]))
                        {
                            pipeLines.Add(lines[idx]);
                            idx++;
                        }

                        int colCount = 0;
                        for (int k = 0; k < pipeLines.Count; k++)
                        {
                            if (k == 1 && IsSeparator(pipeLines[k])) continue;
                            colCount = Math.Max(colCount, SplitCells(pipeLines[k]).Length);
                        }
                        if (colCount == 0) colCount = 1;

                        var table = section.AddTable();
                        table.Borders.Width = 0.75;
                        table.Format.Font.Name = "Segoe UI Symbol";
                        table.Format.Font.Size = 11;

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
                                para.Format.Font.Name = "Segoe UI Symbol";
                                para.Format.Font.Size = 11;
                            }
                        }

                        section.AddParagraph();
                        continue;
                    }

                    Paragraph p = section.AddParagraph();
                    p.Format.Font.Name = "Segoe UI Symbol";
                    p.Format.Font.Size = 11;
                    p.AddText(PdfSafeSymbols(line));

                    idx++;
                }

                var renderer = new PdfDocumentRenderer() { Document = pdfDoc };

                renderer.RenderDocument();
                renderer.Save(dlg.FileName);

                FileOpener.Open(dlg.FileName);
                MessageBox.Show("PDF saved successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("PDF Error: " + ex.Message);
            }
        }

        #endregion

        private void ExportAsWordEquationsCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            // no-op
        }
    }
}