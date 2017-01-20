namespace WindAddin
{
    partial class UserControl_ExtractData
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnBrowse = new DevExpress.XtraEditors.SimpleButton();
            this.label13 = new System.Windows.Forms.Label();
            this.btnGetResult = new DevExpress.XtraEditors.SimpleButton();
            this.cboCoordinate = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkEnvelope = new DevExpress.XtraEditors.CheckEdit();
            ((System.ComponentModel.ISupportInitialize)(this.chkEnvelope.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "|*.dat";
            this.openFileDialog1.Multiselect = true;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(77, 123);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.ShowFocusRectangle = DevExpress.Utils.DefaultBoolean.False;
            this.btnBrowse.Size = new System.Drawing.Size(95, 28);
            this.btnBrowse.TabIndex = 17;
            this.btnBrowse.Text = "选择Dat文件";
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label13.Location = new System.Drawing.Point(20, 75);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(92, 17);
            this.label13.TabIndex = 19;
            this.label13.Text = "选择坐标工作表";
            // 
            // btnGetResult
            // 
            this.btnGetResult.Location = new System.Drawing.Point(77, 223);
            this.btnGetResult.Name = "btnGetResult";
            this.btnGetResult.ShowFocusRectangle = DevExpress.Utils.DefaultBoolean.False;
            this.btnGetResult.Size = new System.Drawing.Size(95, 28);
            this.btnGetResult.TabIndex = 17;
            this.btnGetResult.Text = "提取结果";
            this.btnGetResult.Click += new System.EventHandler(this.btnGetResult_Click);
            // 
            // cboCoordinate
            // 
            this.cboCoordinate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboCoordinate.FormattingEnabled = true;
            this.cboCoordinate.Location = new System.Drawing.Point(118, 75);
            this.cboCoordinate.Name = "cboCoordinate";
            this.cboCoordinate.Size = new System.Drawing.Size(100, 20);
            this.cboCoordinate.TabIndex = 20;
            this.cboCoordinate.DropDown += new System.EventHandler(this.cboCoordinate_DropDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(36, 280);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(179, 12);
            this.label1.TabIndex = 21;
            this.label1.Text = "提取过程中，请不要操作Excel！";
            // 
            // chkEnvelope
            // 
            this.chkEnvelope.EditValue = true;
            this.chkEnvelope.Location = new System.Drawing.Point(60, 177);
            this.chkEnvelope.Name = "chkEnvelope";
            this.chkEnvelope.Properties.Caption = "同时生成正负包络";
            this.chkEnvelope.Size = new System.Drawing.Size(120, 19);
            this.chkEnvelope.TabIndex = 22;
            // 
            // UserControl_ExtractData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.chkEnvelope);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboCoordinate);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.btnGetResult);
            this.Controls.Add(this.btnBrowse);
            this.Name = "UserControl_ExtractData";
            this.Size = new System.Drawing.Size(245, 547);
            ((System.ComponentModel.ISupportInitialize)(this.chkEnvelope.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private DevExpress.XtraEditors.SimpleButton btnBrowse;
        private System.Windows.Forms.Label label13;
        private DevExpress.XtraEditors.SimpleButton btnGetResult;
        private System.Windows.Forms.ComboBox cboCoordinate;
        private System.Windows.Forms.Label label1;
        private DevExpress.XtraEditors.CheckEdit chkEnvelope;
    }
}
