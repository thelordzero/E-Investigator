namespace E_Investigator
{
    partial class DefaultRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DefaultRibbon()
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.bSpam = this.Factory.CreateRibbonButton();
            this.bPossMal = this.Factory.CreateRibbonButton();
            this.bVerMal = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.bPoss = this.Factory.CreateRibbonButton();
            this.bVer = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.bFOPESearch = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.bInspect = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "E-Investigator";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.bSpam);
            this.group1.Items.Add(this.bPossMal);
            this.group1.Items.Add(this.bVerMal);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.bPoss);
            this.group1.Items.Add(this.bVer);
            this.group1.Label = "Phishing";
            this.group1.Name = "group1";
            // 
            // bSpam
            // 
            this.bSpam.Label = "Spam";
            this.bSpam.Name = "bSpam";
            this.bSpam.ShowImage = true;
            this.bSpam.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bSpam_Click);
            // 
            // bPossMal
            // 
            this.bPossMal.Label = "Possible Malicious";
            this.bPossMal.Name = "bPossMal";
            this.bPossMal.ShowImage = true;
            this.bPossMal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bPossMal_Click);
            // 
            // bVerMal
            // 
            this.bVerMal.Label = "Verified Malicious";
            this.bVerMal.Name = "bVerMal";
            this.bVerMal.ShowImage = true;
            this.bVerMal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bVerMal_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // bPoss
            // 
            this.bPoss.Label = "Possible";
            this.bPoss.Name = "bPoss";
            this.bPoss.ShowImage = true;
            this.bPoss.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bPoss_Click);
            // 
            // bVer
            // 
            this.bVer.Label = "Verified";
            this.bVer.Name = "bVer";
            this.bVer.ShowImage = true;
            this.bVer.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bVer_Click);
            // 
            // group2
            // 
            this.group2.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group2.Items.Add(this.bFOPESearch);
            this.group2.Label = "FOPE";
            this.group2.Name = "group2";
            // 
            // bFOPESearch
            // 
            this.bFOPESearch.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.bFOPESearch.Image = global::E_Investigator.Properties.Resources.Forefront_Symbol;
            this.bFOPESearch.Label = "Search";
            this.bFOPESearch.Name = "bFOPESearch";
            this.bFOPESearch.ShowImage = true;
            this.bFOPESearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bFOPESearch_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.bInspect);
            this.group3.Label = "Detail(s)";
            this.group3.Name = "group3";
            this.group3.Visible = false;
            // 
            // bInspect
            // 
            this.bInspect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.bInspect.Image = global::E_Investigator.Properties.Resources.MagnifyGlass;
            this.bInspect.Label = "Inspect";
            this.bInspect.Name = "bInspect";
            this.bInspect.ShowImage = true;
            this.bInspect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bInspect_Click);
            // 
            // DefaultRibbon
            // 
            this.Name = "DefaultRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DefaultRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bSpam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bPossMal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bVerMal;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bPoss;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bVer;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bFOPESearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bInspect;
    }

    partial class ThisRibbonCollection
    {
        internal DefaultRibbon DefaultRibbon
        {
            get { return this.GetRibbon<DefaultRibbon>(); }
        }
    }
}
