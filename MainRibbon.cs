﻿using EDPA_Add_In.Properties;

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
                foreach (Paragraph parag in variablesText.Paragraphs)
                {
                    if (!parag.Range.Text.Equals("\r"))
                    {
                        var values = parag.Range.Text.Contains('(') ? parag.Range.Text.Remove(parag.Range.Text.IndexOf('(')).Split(':') : parag.Range.Text.Split(':');
                        variables.Add($"{values[0].Trim()}:", $"{values[1].Trim()}:");
                    }
                }
                File.Copy(Settings.Default.TemplateFilePath, dialog.FileName, true);
                var newDoc = Globals.ThisAddIn.Application.Documents.Open(dialog.FileName);
                var startIndex = text.IndexOf("\r\r", text.IndexOf("PROCEEDINGS"));
                //text = $"{text.Substring(startIndex).Replace("\r\r", "\r").Trim()}\r{Resources.footerText}";
                string A = null, Q = null, lastSpeaked = null;
                bool defaultSetting = false;
                Paragraphs lines = inputDoc.Range(startIndex).Paragraphs;

                for (int i = 1; i < lines.Count - 1; i++)
                {
                    if (lines[i].Range.Text.Contains("the court was adjourned") || string.IsNullOrWhiteSpace(lines[i].Range.Text))
                    {
                        continue;
                    }
                    else if (lines[i].Range.Text.Contains("EXAMINATION"))
                    {
                        newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Substring(0, lines[i].Range.Text.IndexOf("EXAMINATION") + 11)}\r";
                        newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        var q = lines[i].Range.Text.Substring(lines[i].Range.Text.IndexOf(" BY "));
                        var a = lines[i].Range.Text.Remove(lines[i].Range.Text.IndexOf(" BY ")).Substring(lines[i].Range.Text.IndexOf(" OF ")) + ":";
                        newDoc.Paragraphs.Last.Range.Text = $"{q}\r";
                        Q = variables.FirstOrDefault(x => x.Value == q.Substring(4).Trim()).Key;
                        A = variables.FirstOrDefault(x => x.Value == a.Substring(4).Trim()).Key;
                        defaultSetting = true;
                        lastSpeaked = null;
                    }
                    else if (lines[i].Range.Text.Contains("CERTIFICATION"))
                    {
                        newDoc.Paragraphs.Last.Range.Text = $"\r\r\r\r\r\r{lines[i]}\r";
                        newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
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
                                            continue;
                                        }
                                        if (lastSpeaked == A || lastSpeaked == Q || lastSpeaked == null)
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
                                        if (defaultSetting || lines[i + 1].Range.Text.StartsWith(Q) || lines[i].Range.HighlightColorIndex == WdColorIndex.wdYellow)
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
                            newDoc.Paragraphs.Last.Range.Text = $"{lines[i].Range.Text.Trim()}\r";
                    }
                }
                newDoc.Paragraphs.Last.Range.Text = $"{lines.Last.Range.Text}\r";
                newDoc.Paragraphs.Last.Range.Text = $"{Resources.footerText}\r";

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
