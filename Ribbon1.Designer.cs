
namespace Latin_Cyrillic_TextConvert
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.convert = this.Factory.CreateRibbonGroup();
            this.choose = this.Factory.CreateRibbonMenu();
            this.LatinToCyrillic = this.Factory.CreateRibbonButton();
            this.CyrillicToLatin = this.Factory.CreateRibbonButton();
            this.About = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.convert.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.convert);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // convert
            // 
            this.convert.Items.Add(this.choose);
            this.convert.Label = "Text Convert";
            this.convert.Name = "convert";
            // 
            // choose
            // 
            this.choose.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.choose.Image = global::Latin_Cyrillic_TextConvert.Properties.Resources.AddInIcon;
            this.choose.Items.Add(this.LatinToCyrillic);
            this.choose.Items.Add(this.CyrillicToLatin);
            this.choose.Items.Add(this.About);
            this.choose.Label = "Choose";
            this.choose.Name = "choose";
            this.choose.ShowImage = true;
            // 
            // LatinToCyrillic
            // 
            this.LatinToCyrillic.Label = "Latin to Cyrillic text convert";
            this.LatinToCyrillic.Name = "LatinToCyrillic";
            this.LatinToCyrillic.ShowImage = true;
            this.LatinToCyrillic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LatinToCyrillic_Click);
            // 
            // CyrillicToLatin
            // 
            this.CyrillicToLatin.Label = "Cyrillic to Latin text convert";
            this.CyrillicToLatin.Name = "CyrillicToLatin";
            this.CyrillicToLatin.ShowImage = true;
            this.CyrillicToLatin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CyrillicToLatin_Click);
            // 
            // About
            // 
            this.About.Image = global::Latin_Cyrillic_TextConvert.Properties.Resources.infoicon;
            this.About.Label = "About";
            this.About.Name = "About";
            this.About.ShowImage = true;
            this.About.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.About_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.convert.ResumeLayout(false);
            this.convert.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup convert;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu choose;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LatinToCyrillic;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CyrillicToLatin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton About;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
