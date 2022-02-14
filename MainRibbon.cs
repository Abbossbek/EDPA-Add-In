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
                var startIndex = text.IndexOf("(10:03 o'clock a.m.)") + 20;
                text = text.Substring(startIndex).Replace("\r\r", "\r").Trim();
                bool byAnyOne = false;
                foreach (var item in text.Split('\r'))
                {
                    if (item.Contains("EXAMINATION"))
                    {
                        byAnyOne = true;
                        newDoc.Paragraphs.Last.Range.Text = $"{item.Substring(0, item.IndexOf("EXAMINATION") + 11)}\r";
                        newDoc.Paragraphs[newDoc.Paragraphs.Count - 1].Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        newDoc.Paragraphs.Last.Range.Text = item.Substring(item.IndexOf(" BY ")) + "\r";
                    }
                    else
                    {
                        foreach (var keyValue in variables)
                        {
                            if (item.Contains(keyValue.Key))
                            {
                                if (byAnyOne)
                                {
                                    newDoc.Paragraphs.Last.Range.Text = $"{item.Replace(keyValue.Key, $"{Regex.Replace(keyValue.Key, @"[\d-]", string.Empty)}")}\r";
                                }
                                else
                                {
                                    newDoc.Paragraphs.Last.Range.Text = $"{item.Replace(keyValue.Key, $"\t\t{keyValue.Value}")}\r";
                                }
                            }
                        }
                    }
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
