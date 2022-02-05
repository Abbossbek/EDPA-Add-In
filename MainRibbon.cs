using EDPA_Add_In.Properties;

using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
                    variables.Add(values[0].Trim(), values[1].Trim());
                }
                File.Copy(Settings.Default.TemplateFilePath, dialog.FileName, true);
                var newDoc = Globals.ThisAddIn.Application.Documents.Open(dialog.FileName);
                text = text.Substring(text.IndexOf("(10:03 o'clock a.m.)") + 20);
                foreach (var item in variables)
                {
                    text = text.Replace($"{item.Key}:", $"{item.Value}:").Replace("\r\r", "\r").Trim();
                }
                newDoc.Paragraphs.Last.Range.Text = text;
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
