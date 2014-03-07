namespace ResponsesExportExcel
{
    partial class SurveyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SurveyRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SurveyRibbon));
            this.tabSurvey = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnLogin = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.cbSurvey = this.Factory.CreateRibbonComboBox();
            this.btnSurvey = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.cbQuestion = this.Factory.CreateRibbonComboBox();
            this.btnChart = this.Factory.CreateRibbonButton();
            this.tabSurvey.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            // 
            // tabSurvey
            // 
            this.tabSurvey.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabSurvey.Groups.Add(this.group1);
            this.tabSurvey.Groups.Add(this.group2);
            this.tabSurvey.Groups.Add(this.group3);
            this.tabSurvey.Label = "Survey Tools";
            this.tabSurvey.Name = "tabSurvey";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnLogin);
            this.group1.Label = "Log in";
            this.group1.Name = "group1";
            // 
            // btnLogin
            // 
            this.btnLogin.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLogin.Image = ((System.Drawing.Image)(resources.GetObject("btnLogin.Image")));
            this.btnLogin.Label = "Log In";
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.ShowImage = true;
            this.btnLogin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLogin_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.cbSurvey);
            this.group2.Items.Add(this.btnSurvey);
            this.group2.Label = "Survey";
            this.group2.Name = "group2";
            // 
            // cbSurvey
            // 
            this.cbSurvey.Label = " ";
            this.cbSurvey.Name = "cbSurvey";
            this.cbSurvey.Text = null;
            // 
            // btnSurvey
            // 
            this.btnSurvey.Image = ((System.Drawing.Image)(resources.GetObject("btnSurvey.Image")));
            this.btnSurvey.Label = "Select Survey";
            this.btnSurvey.Name = "btnSurvey";
            this.btnSurvey.ShowImage = true;
            this.btnSurvey.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSurvey_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.cbQuestion);
            this.group3.Items.Add(this.btnChart);
            this.group3.Label = "Chart";
            this.group3.Name = "group3";
            // 
            // cbQuestion
            // 
            this.cbQuestion.Label = " ";
            this.cbQuestion.Name = "cbQuestion";
            this.cbQuestion.Text = null;
            // 
            // btnChart
            // 
            this.btnChart.Image = ((System.Drawing.Image)(resources.GetObject("btnChart.Image")));
            this.btnChart.Label = "Chart";
            this.btnChart.Name = "btnChart";
            this.btnChart.ShowImage = true;
            this.btnChart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChart_Click);
            // 
            // SurveyRibbon
            // 
            this.Name = "SurveyRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabSurvey);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SampleRibbon_Load);
            this.tabSurvey.ResumeLayout(false);
            this.tabSurvey.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabSurvey;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLogin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSurvey;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cbSurvey;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChart;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cbQuestion;

    }

    partial class ThisRibbonCollection
    {
        internal SurveyRibbon SampleRibbon
        {
            get { return this.GetRibbon<SurveyRibbon>(); }
        }
    }
}
