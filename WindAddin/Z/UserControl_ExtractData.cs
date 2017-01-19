using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using DevExpress.XtraEditors;
using System.Threading;

namespace WindAddin
{
    public partial class UserControl_ExtractData : UserControl
    {
        private Microsoft.Office.Interop.Excel.Application xlApp;
        private Workbook xlWb;
        private string[] filePath;
        private Worksheet progressSheet;
        public UserControl_ExtractData()
        {
            InitializeComponent();
            this.xlApp = Globals.ThisAddIn.Application;
            this.xlWb = xlApp.ActiveWorkbook;
        }
        private void cboCoordinate_DropDown(object sender, EventArgs e)
        {
            cboCoordinate.Items.Clear();
            foreach (Worksheet xlSh in xlWb.Worksheets)
            {
                cboCoordinate.Items.Add(xlSh.Name);
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileNames;
                CreateProgressSheet(filePath);
            }
        }

        private void CreateProgressSheet(string[] filePath)
        {
            xlApp.ScreenUpdating = false;
            progressSheet = xlWb.Worksheets.Add(Before: xlWb.Sheets[1]);
            progressSheet.Name = "完成情况";
            progressSheet.Cells[1, 2].Value = "文件路径";
            progressSheet.Cells[1, 3].Value = "角度";
            progressSheet.Cells[1, 4].Value = "完成情况";
            object[,] str = new object[filePath.Length, 3];
            for (int i = 0; i < filePath.Length; i++)
            {
                str[i, 0] = filePath[i];
                MatchCollection mc = Regex.Matches(filePath[i], "\\d+");
                str[i, 1] = decimal.Parse(mc[mc.Count - 1].ToString());
                str[i, 2] = "未开始";
            }
            progressSheet.Range[progressSheet.Cells[2, 2], progressSheet.Cells[1 + filePath.Length, 4]].Value = str;
            ((Range)progressSheet.UsedRange).RowHeight = 23.25;
            ((Range)progressSheet.Rows[1]).RowHeight = 29.25;
            ((Range)progressSheet.Columns[4]).ColumnWidth = 9.5;
            ((Range)progressSheet.Columns[5]).ColumnWidth = 12.4;
            for (int i = 2; i <= progressSheet.UsedRange.Rows.Count; i++)
            {
                if (i % 2 == 0)
                {
                    ((Range)progressSheet.Rows[i]).Interior.Color = Color.FromArgb(242, 242, 242);
                }
            }
            ((Range)progressSheet.Rows[1]).Interior.Color = Color.FromArgb(86, 89, 108);
            ((Range)progressSheet.Rows[1]).HorizontalAlignment = XlVAlign.xlVAlignCenter;
            progressSheet.Range[progressSheet.Cells[2, 3], progressSheet.Cells[1 + filePath.Length, 4]].HorizontalAlignment = XlVAlign.xlVAlignCenter;
            progressSheet.Range["B:B"].Columns.AutoFit();
            ((Range)progressSheet.UsedRange).Font.Name = "微软雅黑";
            ((Range)progressSheet.UsedRange).Font.Size = 11;
            ((Range)progressSheet.UsedRange).Font.Color = Color.FromArgb(64, 64, 64);
            ((Range)progressSheet.Rows[1]).Font.Size = 14;
            ((Range)progressSheet.Rows[1]).Font.Color = Color.White;
            ((Range)progressSheet.Rows[1]).Font.Bold = true;
            ((Range)progressSheet.Range[progressSheet.Cells[2, 4], progressSheet.Cells[1 + filePath.Length, 4]]).Font.Color = Color.Red;
            ((Range)progressSheet.Range[progressSheet.Cells[2, 4], progressSheet.Cells[1 + filePath.Length, 4]]).Font.Bold = true;
            xlApp.ActiveWindow.DisplayGridlines = false;
            //xlApp.ActiveWindow.DisplayHeadings = false;
            progressSheet.Columns["F:XFD"].Hidden = true;
            xlApp.ScreenUpdating = true;
        }

        private void btnGetResult_Click(object sender, EventArgs e)
        {
            progressSheet.Activate();
            Worksheet xlWs;
            try
            {
                xlWs = xlWb.Worksheets[cboCoordinate.Text];
            }
            catch
            {
                XtraMessageBox.Show("请选择坐标工作表！", "提示");
                return;
            }
            for (int i = 0; i < filePath.Length; i++)
            {
                progressSheet.Cells[i + 2, 4].Value = "正在提取";
                progressSheet.Cells[i + 2, 4].Font.Color = Color.Yellow;
                xlApp.ScreenUpdating = false;
                DataFile DaF = new DataFile(filePath[i]);
                FuckExcel FE = new FuckExcel(DaF, xlWb, xlWs);
                FE.ManageExcel();
                xlApp.ScreenUpdating = true;
                progressSheet.Activate();
                progressSheet.Cells[i + 2, 4].Value = "已完成";
                progressSheet.Cells[i + 2, 4].Font.Color = Color.Green;
            }
        }
    }
}
