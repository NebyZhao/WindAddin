namespace WindAddin
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnInsertCoordinate = this.Factory.CreateRibbonToggleButton();
            this.btnExtractData = this.Factory.CreateRibbonToggleButton();
            this.btnWindHistory = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "风洞风振";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.btnInsertCoordinate);
            this.group1.Items.Add(this.btnExtractData);
            this.group1.Label = "风洞模拟";
            this.group1.Name = "group1";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnWindHistory);
            this.group2.Label = "风振验算";
            this.group2.Name = "group2";
            // 
            // btnInsertCoordinate
            // 
            this.btnInsertCoordinate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertCoordinate.Image = ((System.Drawing.Image)(resources.GetObject("btnInsertCoordinate.Image")));
            this.btnInsertCoordinate.Label = "插入坐标";
            this.btnInsertCoordinate.Name = "btnInsertCoordinate";
            this.btnInsertCoordinate.ShowImage = true;
            this.btnInsertCoordinate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertCoordinate_Click);
            // 
            // btnExtractData
            // 
            this.btnExtractData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExtractData.Image = global::WindAddin.Properties.Resources.插入坐标;
            this.btnExtractData.Label = "提取数据";
            this.btnExtractData.Name = "btnExtractData";
            this.btnExtractData.ShowImage = true;
            this.btnExtractData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExtractData_Click);
            // 
            // btnWindHistory
            // 
            this.btnWindHistory.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWindHistory.Image = global::WindAddin.Properties.Resources.风速时程;
            this.btnWindHistory.Label = "计算风速时程";
            this.btnWindHistory.Name = "btnWindHistory";
            this.btnWindHistory.ShowImage = true;
            this.btnWindHistory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWindHistory_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnInsertCoordinate;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnExtractData;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWindHistory;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
