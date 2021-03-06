﻿namespace WindAddin
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
            this.btnGambit = this.Factory.CreateRibbonButton();
            this.btnFluent = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnInsertCoordinate = this.Factory.CreateRibbonToggleButton();
            this.btnExtractData = this.Factory.CreateRibbonToggleButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnWindHistory = this.Factory.CreateRibbonButton();
            this.btnE2k = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnExtractEtabs = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.btnGambit);
            this.group1.Items.Add(this.btnFluent);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.btnInsertCoordinate);
            this.group1.Items.Add(this.btnExtractData);
            this.group1.Label = "风洞模拟";
            this.group1.Name = "group1";
            // 
            // btnGambit
            // 
            this.btnGambit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGambit.Label = "Gambit";
            this.btnGambit.Name = "btnGambit";
            this.btnGambit.ShowImage = true;
            // 
            // btnFluent
            // 
            this.btnFluent.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFluent.Label = "Fluent";
            this.btnFluent.Name = "btnFluent";
            this.btnFluent.ShowImage = true;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
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
            // group2
            // 
            this.group2.Items.Add(this.btnWindHistory);
            this.group2.Items.Add(this.btnE2k);
            this.group2.Items.Add(this.separator2);
            this.group2.Items.Add(this.btnExtractEtabs);
            this.group2.Label = "风振验算";
            this.group2.Name = "group2";
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
            // btnE2k
            // 
            this.btnE2k.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnE2k.Image = global::WindAddin.Properties.Resources.e2k;
            this.btnE2k.Label = "生成e2k";
            this.btnE2k.Name = "btnE2k";
            this.btnE2k.ShowImage = true;
            this.btnE2k.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnE2k_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnExtractEtabs
            // 
            this.btnExtractEtabs.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExtractEtabs.Label = "提取Etabs结果";
            this.btnExtractEtabs.Name = "btnExtractEtabs";
            this.btnExtractEtabs.ShowImage = true;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGambit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFluent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnE2k;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtractEtabs;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
