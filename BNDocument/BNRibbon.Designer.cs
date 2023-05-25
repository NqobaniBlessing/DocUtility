namespace BNDocument
{
    partial class BNRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public BNRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BNRibbon));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.testTab = this.Factory.CreateRibbonTab();
            this.btnGroup = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.button4 = this.Factory.CreateRibbonButton();
            this.testTab.SuspendLayout();
            this.btnGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // testTab
            // 
            this.testTab.Groups.Add(this.btnGroup);
            this.testTab.Label = "Test";
            this.testTab.Name = "testTab";
            // 
            // btnGroup
            // 
            this.btnGroup.Items.Add(this.button1);
            this.btnGroup.Items.Add(this.button2);
            this.btnGroup.Items.Add(this.comboBox1);
            this.btnGroup.Items.Add(this.button4);
            this.btnGroup.Label = "icTools";
            this.btnGroup.Name = "btnGroup";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "First";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Second";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.Image = ((System.Drawing.Image)(resources.GetObject("comboBox1.Image")));
            ribbonDropDownItemImpl1.Label = "Paragraph";
            ribbonDropDownItemImpl2.Label = "Document";
            this.comboBox1.Items.Add(ribbonDropDownItemImpl1);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl2);
            this.comboBox1.Label = "Third";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.ShowImage = true;
            this.comboBox1.Text = null;
            this.comboBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBox1_TextChanged);
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.Label = "Fourth";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // BNRibbon
            // 
            this.Name = "BNRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.testTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.BNRibbon_Load);
            this.testTab.ResumeLayout(false);
            this.testTab.PerformLayout();
            this.btnGroup.ResumeLayout(false);
            this.btnGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab testTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup btnGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox1;
    }

    partial class ThisRibbonCollection
    {
        internal BNRibbon BNRibbon
        {
            get { return this.GetRibbon<BNRibbon>(); }
        }
    }
}
