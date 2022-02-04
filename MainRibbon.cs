using EDPA_Add_In.Properties;

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
            while(string.IsNullOrWhiteSpace(Settings.Default.TemplateFilePath) || !File.Exists(Settings.Default.TemplateFilePath))
            {
                btnSelectTemplate_Click(null, null);
            }
            var text = Globals.ThisAddIn.Application.ActiveDocument.Range(0, Globals.ThisAddIn.Application.ActiveDocument.Characters.Count).Text;
            var variables = text.Substring(0, text.IndexOf("PROCEEDINGS"));
        }

        private void btnSelectTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
                dialog.Title = "Select template";
            dialog.Filter = "Document|*.docx";
            if(dialog.ShowDialog() == DialogResult.OK)
            {
                Settings.Default.TemplateFilePath = dialog.FileName;
                Settings.Default.Save();
            }
        }
    }
}
