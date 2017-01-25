using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Drawing;
using DevExpress.XtraEditors;

namespace WindAddin
{
    public class E2KSheet
    {
        private Workbook xlWb
        {
            get { return Globals.ThisAddIn.Application.ActiveWorkbook; }
        }
        private Microsoft.Office.Tools.Excel.Worksheet xlWs;
        public void CreateSheet()
        {
            xlWs = Globals.Factory.GetVstoObject(xlWb.Worksheets.Add(Before: xlWb.Sheets[1]));
            xlWs.Name = "生成e2k文件";
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            ((Range)xlWs.Rows).RowHeight = 23.25;
            ((Range)xlWs.Rows).HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ((Range)xlWs.Rows[1]).RowHeight = 29.25;
            ((Range)xlWs.Columns["B:F"]).ColumnWidth = 12;
            xlWs.Range["H7:M7"].Merge();
            xlWs.Range["I9:J9"].Merge();
            xlWs.Range["K9:L9"].Merge();
            xlWs.Range["J11:M11"].Merge();
            xlWs.Range["J13:M13"].Merge();
            xlWs.Range["I15:K15"].Merge();
            ((Range)xlWs.Columns).Font.Name = "微软雅黑";
            ((Range)xlWs.Columns).Font.Color = Color.FromArgb(64, 64, 64);
            ((Range)xlWs.Range["B:F,I9,I15:M15"]).Font.Size = 14;
            ((Range)xlWs.Range["H7,J11,J13"]).Font.Size = 12;
            xlWs.Range["B:F,O:XFD,J11,J13,L15"].Interior.Color = Color.FromArgb(242, 242, 242);
            xlWs.Range["B1:F1"].Interior.Color = Color.FromArgb(86, 89, 108);
            xlWs.Range["B1:F1"].Font.Color = Color.White;
            xlWs.Range["B1:F1,I9,I15,M15"].Font.Bold = true;
            xlWs.Columns["O:XFD"].Hidden = true;
            Globals.ThisAddIn.Application.ActiveWindow.DisplayGridlines = false;
            xlWs.Range["B1"].Value = "层号";
            xlWs.Range["C1"].Value = "面积(m\u00b2)";
            xlWs.Range["D1"].Value = "体型系数";
            xlWs.Range["E1"].Value = "点号";
            xlWs.Range["F1"].Value = "角度";
            xlWs.Range["H7"].Value = "填写左侧表格选择文件路径后，点击“生成e2k文件”";
            xlWs.Range["I9"].Value = "风速时程表单：";
            xlWs.Range["I15"].Value = "建筑一层对应模型中的";
            xlWs.Range["L15"].Value = 1;
            xlWs.Range["M15"].Value = "层";
            xlWs.Range["I15"].HorizontalAlignment = XlHAlign.xlHAlignRight;
            xlWs.Range["M15"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            PictureBox pb1 = new PictureBox()
            {
                SizeMode = PictureBoxSizeMode.Zoom,
                Image = Properties.Resources.层号点号
            };
            PictureBox pb2 = new PictureBox()
            {
                SizeMode = PictureBoxSizeMode.Zoom,
                Image = Properties.Resources.角度
            };
            xlWs.Controls.AddControl(pb1, xlWs.Range["H2:J5"], "pb1");
            xlWs.Controls.AddControl(pb2, xlWs.Range["K2:M5"], "pb2");
            System.Windows.Forms.ComboBox cboHistSheet = new System.Windows.Forms.ComboBox()
            {
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboHistSheet.Font = new System.Drawing.Font(cboHistSheet.Font.FontFamily, 15);
            cboHistSheet.DropDown += CboHistSheet_DropDown;
            cboHistSheet.SelectedIndexChanged += CboHistSheet_SelectedIndexChanged;
            xlWs.Controls.AddControl(cboHistSheet, xlWs.Range["K9:L9"], "cboHistSheet");
            SimpleButton btnResultFolder = new SimpleButton()
            {
                Text = "结果保存位置",
                ShowFocusRectangle = DevExpress.Utils.DefaultBoolean.False
            };
            btnResultFolder.Font = new System.Drawing.Font(btnResultFolder.Font.FontFamily, 12);
            btnResultFolder.Click += BtnResultFolder_Click;
            xlWs.Controls.AddControl(btnResultFolder, xlWs.Range["H11:I11"], "btnResultFolder");  
            SimpleButton btnSelectE2K = new SimpleButton()
            {
                Text = "选择e2k文件",
                ShowFocusRectangle = DevExpress.Utils.DefaultBoolean.False
            };
            btnSelectE2K.Font = new System.Drawing.Font(btnSelectE2K.Font.FontFamily, 12);
            btnSelectE2K.Click += BtnSelectE2K_Click; ;
            xlWs.Controls.AddControl(btnSelectE2K, xlWs.Range["H13:I13"], "btnSelectE2K");
            SimpleButton btnProduceE2K = new SimpleButton()
            {
                Text = "生成e2k文件",
                ShowFocusRectangle = DevExpress.Utils.DefaultBoolean.False
            };
            btnProduceE2K.Font = new System.Drawing.Font(btnProduceE2K.Font.FontFamily, 12);
            btnProduceE2K.Click += BtnProduceE2K_Click; ;
            xlWs.Controls.AddControl(btnProduceE2K, xlWs.Range["I18:L18"], "btnProduceE2K");
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void CboHistSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Range xlRng = xlWs.Columns["O:XFD"];
            xlRng.Hidden = false;
            xlRng.Font.Color = Color.FromArgb(64, 64, 64);
            xlRng.Font.Size = 11;
            xlRng.ColumnWidth = 9;
            xlRng.ClearContents();
            Worksheet ws = xlWb.Sheets[((System.Windows.Forms.ComboBox)sender).SelectedItem.ToString()];
            Range history = ws.Range["H2"].CurrentRegion;
            history.Copy();
            xlWs.Range["O1"].Select();
            xlWs.Paste(Link: true);
            history.Rows[1].copy(xlWs.Range["O1"]);
            for (int i = 0; i < history.Columns.Count; i++)
            {
                if (i == 0)
                {
                    xlWs.Cells[2, 15 + i].Formula = "";
                }
                else
                {
                    xlWs.Cells[2, 15 + i].Formula += "-L15+1";
                }
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void BtnSelectE2K_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "e2k文件|*.e2k";
            ofd.FileName = "";
            ofd.Title = "选择e2k文件";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                xlWs.Range["J13:M13"].Value = ofd.FileName;
            }
        }

        private void BtnProduceE2K_Click(object sender, EventArgs e)
        {
  
        }

        private void CboHistSheet_DropDown(object sender, EventArgs e)
        {
            ((System.Windows.Forms.ComboBox)sender).Items.Clear();
            foreach (Worksheet xlSh in xlWb.Worksheets)
            {
                ((System.Windows.Forms.ComboBox)sender).Items.Add(xlSh.Name);
            }
        }

        private void BtnResultFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if(fbd.ShowDialog()== DialogResult.OK)
            {
                xlWs.Range["J11:M11"].Value = fbd.SelectedPath;
            }
        }
    }
}
