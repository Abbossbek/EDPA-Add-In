using EDPA_Add_In.Properties;

using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace EDPA_Add_In
{
    public partial class MainRibbon
    {
        private static string[] keywords = new string[] { "objection", "relevance", "no further questions" };
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Settings.Default.TemplateFilePath) || !File.Exists(Settings.Default.TemplateFilePath))
            {
                MessageBox.Show("You need to select template file first!");
                return;
            }
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Document|*.docx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var inputDoc = Globals.ThisAddIn.Application.ActiveDocument;
                var text = inputDoc.Range(0, inputDoc.Characters.Count).Text;
                var variablesText = inputDoc.Range(0, text.IndexOf("PROCEEDINGS"));
                var variables = new Dictionary<string, string>();
                Index indexPlaintiff = null, indexDefendant = null;
                var exhibits = new List<Exhibit>();
                foreach (Paragraph parag in variablesText.Paragraphs)
                {
                    if (!parag.Range.Text.Equals("\r"))
                    {
                        var values = parag.Range.Text.Contains('(') ? parag.Range.Text.Remove(parag.Range.Text.IndexOf('(')).Split(':') : parag.Range.Text.Split(':');
                        variables.Add($"{values[0].Trim()}:", $"{values[1].Trim()}:");
                        if (parag.Range.Text.Contains("Plaintiff") && !parag.Range.Text.Contains("Examiner"))
                            indexPlaintiff = new Index { Witness = values[1].Trim() };
                        if (parag.Range.Text.Contains("Defendant") && !parag.Range.Text.Contains("Examiner"))
                            indexDefendant = new Index { Witness = values[1].Trim() };
                    }
                }
                File.Copy(Settings.Default.TemplateFilePath, dialog.FileName, true);
                var newDoc = Globals.ThisAddIn.Application.Documents.Open(dialog.FileName);
                var startIndex = text.IndexOf("\r\r", text.IndexOf("PROCEEDINGS"));
                //text = $"{text.Substring(startIndex).Replace("\r\r", "\r").Trim()}\r{Resources.footerText}";
                string A = null, Q = null, lastSpeaked = null;
                bool defaultSetting = false;
                List<Paragraph> lines = inputDoc.Range(startIndex).Paragraphs.Cast<Paragraph>().Where(x => !string.IsNullOrWhiteSpace(x.Range.Text.Trim())).ToList();

                for (int i = 1; i < lines.Count - 1; i++)
                {
                    if (lines[i].Range.Text.Contains("EXAMINATION"))
                    {
                        newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Substring(0, lines[i].Range.Text.IndexOf("EXAMINATION") + 11)}\r";
                        newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        var pageNumber = newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Range.Information[WdInformation.wdActiveEndPageNumber];
                        var q = lines[i].Range.Text.Substring(lines[i].Range.Text.IndexOf(" BY "));
                        var a = lines[i].Range.Text.Remove(lines[i].Range.Text.IndexOf(" BY ")).Substring(lines[i].Range.Text.IndexOf(" OF ")) + ":";
                        newDoc.Paragraphs.Last.Range.Text = $"{q.TrimEnd()}\r";
                        Q = variables.FirstOrDefault(x => x.Value == q.Substring(4).Trim()).Key;
                        A = variables.FirstOrDefault(x => x.Value == a.Substring(4).Trim()).Key;
                        if (lines[i].Range.Text.Contains("DIRECT"))
                        {
                            if (a.Contains(indexDefendant.Witness))
                                indexDefendant.DirectPage = pageNumber;
                            if (a.Contains(indexPlaintiff.Witness))
                                indexPlaintiff.DirectPage = pageNumber;
                        }
                        if (lines[i].Range.Text.Contains("CROSS"))
                        {
                            if (a.Contains(indexDefendant.Witness))
                                indexDefendant.CrossPage = pageNumber;
                            if (a.Contains(indexPlaintiff.Witness))
                                indexPlaintiff.CrossPage = pageNumber;
                        }
                        defaultSetting = true;
                        lastSpeaked = null;
                    }
                    else
                    {
                        if (variables.Keys.Any(x => lines[i].Range.Text.Contains(x)))
                            foreach (var keyValue in variables)
                            {
                                if (lines[i].Range.Text.Contains(keyValue.Key))
                                {
                                    if (Q != null && lines[i].Range.Text.StartsWith(Q))
                                    {
                                        if (keywords.Any(x => lines[i].Range.Text.ToLower().Contains(x)))
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Replace(Q, $"\t\t{variables[Q]}").TrimEnd()}\r";
                                            lastSpeaked = Q;
                                            continue;
                                        }
                                        if (lines[i].Range.HighlightColorIndex == WdColorIndex.wdYellow)
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Replace(Q, $"\t\t{variables[Q]}").TrimEnd()}\r";
                                            lastSpeaked = Q;
                                            continue;
                                        }
                                        if (lastSpeaked == A || lastSpeaked == Q || lastSpeaked == null || lastSpeaked != A && lastSpeaked != Q && lines[i + 1].Range.Text.StartsWith(A))
                                        {
                                            if (!defaultSetting)
                                            {
                                                newDoc.Paragraphs.Last.Range.Text = $"BY {variables[Q].TrimEnd()}\r";
                                                defaultSetting = true;
                                            }
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Replace(Q, $"{Regex.Replace(Q, @"[\d-]", string.Empty)}").Replace(':', '.').TrimEnd()}\r";
                                        }
                                        else if (!defaultSetting)
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Replace(Q, $"\t\t{variables[Q]}").TrimEnd()}\r";
                                        }
                                        lastSpeaked = Q;
                                    }
                                    else if (A != null && lines[i].Range.Text.StartsWith(A))
                                    {
                                        if (lines[i].Range.HighlightColorIndex == WdColorIndex.wdYellow)
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Replace(A, $"\t\t{variables[A]}").TrimEnd()}\r";
                                            lastSpeaked = A;
                                            continue;
                                        }
                                        if (defaultSetting || lines[i + 1].Range.Text.StartsWith(Q))
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Replace(A, $"{Regex.Replace(A, @"[\d-]", string.Empty)}").Replace(':','.').TrimEnd()}\r";
                                        }
                                        else
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Replace(A, $"\t\t{variables[A]}").TrimEnd()}\r";
                                        }
                                        lastSpeaked = A;
                                    }
                                    else
                                    {
                                        newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Replace(keyValue.Key, $"\t\t{keyValue.Value}").TrimEnd()}\r";
                                        defaultSetting = false;
                                        lastSpeaked = keyValue.Key;
                                    }
                                }

                            }
                        else
                        {
                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.TrimEnd()}\r";
                            if (lines[i].Range.Text.Contains("Exhibit") && lines[i].Range.Text.Contains("marked"))
                            {
                                newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                var exhibit = new Exhibit { Page = newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Range.Information[WdInformation.wdActiveEndPageNumber] };
                                if (lines[i].Range.Text.Contains("Defendant"))
                                    exhibit.Name = $"D-{lines[i].Range.Text.Substring(lines[i].Range.Text.IndexOf("-") + 1, 2)}";
                                if (lines[i].Range.Text.Contains("Plaintiff"))
                                    exhibit.Name = $"P-{lines[i].Range.Text.Substring(lines[i].Range.Text.IndexOf("-") + 1, 2)}";
                                exhibits.Add(exhibit);
                            }
                        }
                    }
                }
                if (lines.Last().Range.Text.Contains("court was adjourned"))
                {
                    newDoc.Paragraphs.Last.Range.Text = "(Court adjourned)\r";
                }
                else
                    newDoc.Paragraphs.Last.Range.Text = $"{lines.Last().Range.Text}\r";
                var page = newDoc.Paragraphs.Last.Range.Information[WdInformation.wdActiveEndPageNumber];
                while (page == newDoc.Paragraphs.Last.Range.Information[WdInformation.wdActiveEndPageNumber])
                {
                    newDoc.Paragraphs.Last.Range.InsertParagraphAfter();
                }
                //newDoc.Paragraphs.Last.Range.InsertBreak(WdBreakType.wdPageBreak);
                newDoc.Paragraphs.Last.Range.Text = $"CERTIFICATION\r";
                newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                newDoc.Paragraphs.Last.Range.Text = $"{Resources.footerText}\r";

                try
                {
                    var plaintiffRow = newDoc.Tables[3].Rows[4];
                    plaintiffRow.Cells[1].Range.Text = $"{indexPlaintiff.Witness}";
                    plaintiffRow.Cells[2].Range.Text = indexPlaintiff.DirectPage.ToString();
                    plaintiffRow.Cells[3].Range.Text = indexPlaintiff.CrossPage.ToString();
                    var defendantRow = newDoc.Tables[3].Rows[9];
                    defendantRow.Cells[1].Range.Text = $"{indexDefendant.Witness}";
                    defendantRow.Cells[2].Range.Text = indexDefendant.DirectPage.ToString();
                    defendantRow.Cells[3].Range.Text = indexDefendant.CrossPage.ToString();
                    foreach (var exhibit in exhibits)
                    {
                        var exhibitRow = newDoc.Tables[4].Rows.Add();
                        exhibitRow.Cells[1].Range.Text = $"{exhibit.Name}";
                        exhibitRow.Cells[2].Range.Text = $"Letter to Judge";
                        exhibitRow.Cells[3].Range.Text = exhibit.Page.ToString();
                    }
                }
                catch
                {

                }
                //var directExamIndex = text.IndexOf("DIRECT EXAMINATION");
                //var firstPart = text.Substring(startIndex, directExamIndex - startIndex);
                //foreach (var item in variables)
                //{
                //    firstPart = firstPart.Replace($"{item.Key}:", $"\t\t{item.Value}:").Replace("\r\r", "\r").Trim();
                //}
                //newDoc.Paragraphs.Last.Range.Text = firstPart + "\r";

                //newDoc.Paragraphs.Last.Range.Text = "DIRECT EXAMINATION\r";
                //newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                //var crossExamIndex = text.IndexOf("CROSS EXAMINATION");
                //var secondPart = text.Substring(directExamIndex, crossExamIndex - directExamIndex);
                //secondPart = secondPart.Substring(secondPart.IndexOf(" BY "));
                //newDoc.Paragraphs.Last.Range.Text = secondPart + "\r";

                //newDoc.Paragraphs.Last.Range.Text = "CROSS EXAMINATION\r";
                //newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                //var thirdPart = text.Substring(crossExamIndex, crossExamIndex - directExamIndex);
                //secondPart = secondPart.Substring(secondPart.IndexOf(" BY "));
                //newDoc.Paragraphs.Last.Range.Text = secondPart + "\r";
                //var page = newDoc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, 3, 3);
                //page.Bookmarks["\\Page"].Range.Delete();

            }
        }
        private void btnSelectTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select template";
            dialog.Filter = "Document|*.docx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Settings.Default.TemplateFilePath = dialog.FileName;
                Settings.Default.Save();
            }
        }
    }
}
