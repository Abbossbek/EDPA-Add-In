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
        private static string[] keywords = new string[] {"objection", "relevance", "no further questions"};
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
                var text = Globals.ThisAddIn.Application.ActiveDocument.Range(0, Globals.ThisAddIn.Application.ActiveDocument.Characters.Count).Text;
                var variablesText = text.Substring(0, text.IndexOf("PROCEEDINGS")).Trim();
                var variables = new Dictionary<string, string>();
                foreach (var row in variablesText.Split('\r'))
                {
                    var values = row.Contains('(') ? row.Remove(row.IndexOf('(')).Split(':') : row.Split(':');
                    variables.Add($"{values[0].Trim()}:", $"{values[1].Trim()}:");
                }
                File.Copy(Settings.Default.TemplateFilePath, dialog.FileName, true);
                var newDoc = Globals.ThisAddIn.Application.Documents.Open(dialog.FileName);
                var startIndex = text.IndexOf("\r\r", text.IndexOf("PROCEEDINGS"));
                text = $"{text.Substring(startIndex).Replace("\r\r", "\r").Trim()}\r{Resources.footerText}";
                string A = null, Q = null, lastSpeaked = null;
                bool defaultSetting = false;
                var lines = text.Split('\r');
                for (int i = 0; i < lines.Length - 1; i++)
                {
                    if (lines[i].Contains("EXAMINATION"))
                    {
                        newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Substring(0, lines[i].IndexOf("EXAMINATION") + 11)}\r";
                        newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        var q = lines[i].Substring(lines[i].IndexOf(" BY "));
                        var a = lines[i].Remove(lines[i].IndexOf(" BY ")).Substring(lines[i].IndexOf(" OF ")) + ":";
                        newDoc.Paragraphs.Last.Range.Text = $"{q}\r";
                        Q = variables.FirstOrDefault(x => x.Value == q.Substring(4).Trim()).Key;
                        A = variables.FirstOrDefault(x => x.Value == a.Substring(4).Trim()).Key;
                        defaultSetting = true;
                        lastSpeaked = null;
                    }
                    else if (lines[i].Contains("CERTIFICATION"))
                    {
                        newDoc.Paragraphs.Last.Range.Text = $"\r\r\r\r\r\r{lines[i]}\r";
                        newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else if (lines[i].Contains("the court was adjourned"))
                    {
                        continue;
                    }
                    else
                    {
                        if (variables.Keys.Any(x => lines[i].Contains(x)))
                            foreach (var keyValue in variables)
                            {
                                if (lines[i].Contains(keyValue.Key))
                                {
                                    if (Q != null && lines[i].StartsWith(Q))
                                    {
                                        if (keywords.Any(x => lines[i].ToLower().Contains(x)))
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Replace(Q, $"\t\t{variables[Q]}")}\r";
                                            continue;
                                        }
                                        if (lastSpeaked == A || lastSpeaked == Q || lastSpeaked == null)
                                        {
                                            if (!defaultSetting)
                                            {
                                                newDoc.Paragraphs.Last.Range.Text = $"BY {variables[Q]}\r";
                                                defaultSetting = true;
                                            }
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Replace(Q, $"{Regex.Replace(Q, @"[\d-]", string.Empty)}")}\r";
                                        }
                                        else if (!defaultSetting)
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Replace(Q, $"\t\t{variables[Q]}")}\r";
                                        }
                                        lastSpeaked = Q;
                                    }
                                    else if (A != null && lines[i].StartsWith(A))
                                    {
                                        if (defaultSetting)
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Replace(A, $"{Regex.Replace(A, @"[\d-]", string.Empty)}")}\r";
                                        }
                                        else if (lines[i + 1].StartsWith(Q))
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Replace(A, $"{Regex.Replace(A, @"[\d-]", string.Empty)}")}\r";
                                        }
                                        else
                                        {
                                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Replace(A, $"\t\t{variables[A]}")}\r";
                                        }
                                        lastSpeaked = A;
                                    }
                                    else
                                    {
                                        newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Replace(keyValue.Key, $"\t\t{keyValue.Value}")}\r";
                                        defaultSetting = false;
                                        lastSpeaked = keyValue.Key;
                                    }
                                }

                            }
                        else
                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i]}\r";
                    }
                }
                newDoc.Paragraphs.Last.Range.Text = $"{lines.Last()}\r";

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
