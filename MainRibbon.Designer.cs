
namespace EDPA_Add_In
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupEdpa = this.Factory.CreateRibbonGroup();
            this.btnGenerate = this.Factory.CreateRibbonButton();
            this.btnSelectTemplate = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupEdpa.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.groupEdpa);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // groupEdpa
            // 
            this.groupEdpa.Items.Add(this.btnGenerate);
            this.groupEdpa.Items.Add(this.btnSelectTemplate);
            this.groupEdpa.Label = "EDPA";
            this.groupEdpa.Name = "groupEdpa";
            // 
            // btnGenerate
            // 
            this.btnGenerate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGenerate.Image = global::EDPA_Add_In.Properties.Resources.img_397790;
            this.btnGenerate.Label = "Generate";
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.ShowImage = true;
            this.btnGenerate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGenerate_Click);
            // 
            // btnSelectTemplate
            // 
            this.btnSelectTemplate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSelectTemplate.Image = global::EDPA_Add_In.Properties.Resources.img_282126;
            this.btnSelectTemplate.Label = "Select template";
            this.btnSelectTemplate.Name = "btnSelectTemplate";
            this.btnSelectTemplate.ShowImage = true;
            this.btnSelectTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectTemplate_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupEdpa.ResumeLayout(false);
            this.groupEdpa.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupEdpa;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGenerate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectTemplate;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
