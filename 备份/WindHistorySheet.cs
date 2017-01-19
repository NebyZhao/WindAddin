using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using DevExpress.XtraEditors;
using System.Drawing;

namespace WindAddin
{
   public static class WindHistorySheet
    {
        private static Workbook xlWb = Globals.ThisAddIn.Application.ActiveWorkbook;
        private static Microsoft.Office.Tools.Excel.Worksheet xlWs;
        public static void CreateSheet()
        {
            xlWs = Globals.Factory.GetVstoObject(xlWb.Worksheets.Add(Before: xlWb.Sheets[1]));
            xlWs.Name = "风速时程计算";
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            ((Range)xlWs.Rows).RowHeight = 23.25;
            ((Range)xlWs.Rows).HorizontalAlignment = XlVAlign.xlVAlignCenter;
            ((Range)xlWs.Rows[1]).RowHeight = 29.25;
            ((Range)xlWs.Columns["B:C"]).ColumnWidth = 12;
            ((Range)xlWs.Columns["E:E"]).ColumnWidth = 23.75;
            ((Range)xlWs.Columns["F:F"]).ColumnWidth = 14;
            xlWs.Range["E2:F3"].Merge();
            xlWs.Range["E2:F3"].WrapText=true;
            ((Range)xlWs.Columns).Font.Name = "微软雅黑";
            ((Range)xlWs.Columns).Font.Color = Color.FromArgb(64, 64, 64);
            ((Range)xlWs.Range["B:C,F:F"]).Font.Size = 14;
            ((Range)xlWs.Columns["E:E"]).Font.Size = 12;
            xlWs.Range["B:C,H:XFD,F4,F6,F8,F10"].Interior.Color = Color.FromArgb(242, 242, 242);
            xlWs.Range["B1:C1,E4,E6,E8,E10"].Interior.Color= Color.FromArgb(86, 89, 108);
            xlWs.Range["B1:C1,E4,E6,E8,E10"].Font.Color = Color.White;
            xlWs.Range["B1:C1,E4,E6,E8,E10"].Font.Bold = true;
            xlWs.Columns["H:XFD"].Hidden = true;
            Globals.ThisAddIn.Application.ActiveWindow.DisplayGridlines = false;
            Range rng = xlWs.Range["E13:F13"];
            xlWs.Range["B1"].Value = "层号";
            xlWs.Range["C1"].Value = "层高(m)";
            xlWs.Range["E2"].Value = "填写完左侧表格和下面参数后，点击“计算风速时程”";
            xlWs.Range["E4"].Value = "地貌参数";
            xlWs.Range["E6"].Value = "风谱频率上限(Hz)";
            xlWs.Range["E8"].Value = "10米高度处风速(m/s)";
            xlWs.Range["E10"].Value = "地面粗糙度系数";
            xlWs.Range["F6"].Value = 4;
            SimpleButton button = new SimpleButton()
            {
                Text = "计算风速时程",
                ShowFocusRectangle = DevExpress.Utils.DefaultBoolean.False
            };
            button.Font = new System.Drawing.Font(button.Font.FontFamily, 12);
            button.Click += Button_Click;
            xlWs.Controls.AddControl(button, rng, "btnCalculateWindHistory");        
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private static void Button_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            xlWs.Columns["H:XFD"].Hidden = false;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
    }
}
