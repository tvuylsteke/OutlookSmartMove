namespace OutlookSmartMove
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
            this.d = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.detectButton = this.Factory.CreateRibbonButton();
            this.moveButton = this.Factory.CreateRibbonButton();
            this.moveOptions = this.Factory.CreateRibbonGallery();
            this.folderBox = this.Factory.CreateRibbonEditBox();
            this.createButton = this.Factory.CreateRibbonButton();
            this.learnButton = this.Factory.CreateRibbonButton();
            this.searchButton = this.Factory.CreateRibbonButton();
            this.focusButton = this.Factory.CreateRibbonButton();
            this.homeButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.d.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.d);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // d
            // 
            this.d.Items.Add(this.box1);
            this.d.Items.Add(this.moveOptions);
            this.d.Items.Add(this.folderBox);
            this.d.Items.Add(this.createButton);
            this.d.Items.Add(this.learnButton);
            this.d.Items.Add(this.searchButton);
            this.d.Items.Add(this.focusButton);
            this.d.Items.Add(this.homeButton);
            this.d.Label = "Smart Move";
            this.d.Name = "d";
            // 
            // box1
            // 
            this.box1.Items.Add(this.detectButton);
            this.box1.Items.Add(this.moveButton);
            this.box1.Name = "box1";
            // 
            // detectButton
            // 
            this.detectButton.Label = "Detect";
            this.detectButton.Name = "detectButton";
            this.detectButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.detectButton_Click);
            // 
            // moveButton
            // 
            this.moveButton.Label = "Move";
            this.moveButton.Name = "moveButton";
            this.moveButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.moveButton_Click);
            // 
            // moveOptions
            // 
            this.moveOptions.Label = "Moves";
            this.moveOptions.Name = "moveOptions";
            this.moveOptions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallery1_Click);
            // 
            // folderBox
            // 
            this.folderBox.Label = "Folder";
            this.folderBox.Name = "folderBox";
            this.folderBox.ShowLabel = false;
            this.folderBox.SizeString = "Long example customer";
            this.folderBox.Text = null;
            this.folderBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.folderBox_TextChanged);
            // 
            // createButton
            // 
            this.createButton.Label = "Create";
            this.createButton.Name = "createButton";
            this.createButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createButton_Click);
            // 
            // learnButton
            // 
            this.learnButton.Label = "Learn";
            this.learnButton.Name = "learnButton";
            this.learnButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.learnButton_Click);
            // 
            // searchButton
            // 
            this.searchButton.Label = "Search";
            this.searchButton.Name = "searchButton";
            this.searchButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.searchButton_Click);
            // 
            // focusButton
            // 
            this.focusButton.Label = "Focus";
            this.focusButton.Name = "focusButton";
            this.focusButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.focusButton_Click);
            // 
            // homeButton
            // 
            this.homeButton.Label = "Home";
            this.homeButton.Name = "homeButton";
            this.homeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.homeButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.d.ResumeLayout(false);
            this.d.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup d;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton moveButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton detectButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery moveOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton learnButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox folderBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton searchButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton focusButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton homeButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
