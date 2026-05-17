namespace WordAddIn1
{
    partial class AltTextRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AltTextRibbon()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.toggleButtonHideShowAltTextPane = this.Factory.CreateRibbonToggleButton();
            this.checkBoxDisplayOnStartup = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.toggleButtonHideShowAltTextPane);
            this.group1.Items.Add(this.checkBoxDisplayOnStartup);
            this.group1.Label = "Alt Text Add-In";
            this.group1.Name = "group1";
            // 
            // toggleButtonHideShowAltTextPane
            // 
            this.toggleButtonHideShowAltTextPane.Image = global::WordAddIn1.Properties.Resources.DockPanelElement_10705_24;
            this.toggleButtonHideShowAltTextPane.Label = "Display Alt Text Pane";
            this.toggleButtonHideShowAltTextPane.Name = "toggleButtonHideShowAltTextPane";
            this.toggleButtonHideShowAltTextPane.ShowImage = true;
            this.toggleButtonHideShowAltTextPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonHideShowAltTextPane_Click);
            // 
            // checkBoxDisplayOnStartup
            // 
            this.checkBoxDisplayOnStartup.Checked = true;
            this.checkBoxDisplayOnStartup.Label = "Display on startup";
            this.checkBoxDisplayOnStartup.Name = "checkBoxDisplayOnStartup";
            this.checkBoxDisplayOnStartup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBoxDisplayOnStartup_Click);
            // 
            // AltTextRibbon
            // 
            this.Name = "AltTextRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AltTextRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonHideShowAltTextPane;
        public Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxDisplayOnStartup;
    }

    partial class ThisRibbonCollection
    {
        internal AltTextRibbon AltTextRibbon
        {
            get { return this.GetRibbon<AltTextRibbon>(); }
        }
    }
}
