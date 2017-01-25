using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using DevExpress.XtraEditors;
using System.Drawing;

namespace WindAddin
{
    public class WindHistorySheet
    {
        private Workbook xlWb
        {
            get { return Globals.ThisAddIn.Application.ActiveWorkbook; }
        }
        private Microsoft.Office.Tools.Excel.Worksheet xlWs;
        public void CreateSheet()
        {
            xlWs = Globals.Factory.GetVstoObject(xlWb.Worksheets.Add(Before: xlWb.Sheets[1]));
            xlWs.Name = "风速时程计算";
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            ((Range)xlWs.Rows).RowHeight = 23.25;
            ((Range)xlWs.Rows).HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ((Range)xlWs.Rows[1]).RowHeight = 29.25;
            ((Range)xlWs.Columns["B:C"]).ColumnWidth = 12;
            ((Range)xlWs.Columns["E:E"]).ColumnWidth = 23.75;
            ((Range)xlWs.Columns["F:F"]).ColumnWidth = 14;
            xlWs.Range["E2:F3"].Merge();
            xlWs.Range["E2:F3"].WrapText = true;
            ((Range)xlWs.Columns).Font.Name = "微软雅黑";
            ((Range)xlWs.Columns).Font.Color = Color.FromArgb(64, 64, 64);
            ((Range)xlWs.Range["B:C,F:F"]).Font.Size = 14;
            ((Range)xlWs.Columns["E:E"]).Font.Size = 12;
            xlWs.Range["B:C,H:XFD,F4,F6,F8,F10"].Interior.Color = Color.FromArgb(242, 242, 242);
            xlWs.Range["B1:C1,E4,E6,E8,E10"].Interior.Color = Color.FromArgb(86, 89, 108);
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
            button.Click += btnCalculateWindHistory_Click;
            xlWs.Controls.AddControl(button, rng, "btnCalculateWindHistory");
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void btnCalculateWindHistory_Click(object sender, EventArgs e)
        {


            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Range xlRng = xlWs.Columns["H:XFD"];
            xlRng.Hidden = false;
            xlRng.Font.Color = Color.FromArgb(64, 64, 64);
            xlRng.Font.Size = 11;
            xlRng.ColumnWidth = 9;

            //计算风速时程
            Range history = xlWb.Sheets["Sheet1"].UsedRange;    //测试
            history.Copy();    //测试
            ((Range)xlRng.Cells[2, 2]).PasteSpecial(XlPasteType.xlPasteValues);    //测试
            for (int i = 1; i < history.Rows.Count; i++)
            {
                xlRng.Cells[i + 2, 1].Value = i * 0.1;
            }
            Range title = xlWs.Range[xlRng.Cells[1, history.Columns.Count / 2 + 1], xlRng.Cells[1, history.Columns.Count / 2 + 2]];
            title.Merge();
            title.Value = "风 速 时 程";
            title.Font.Size = 14;
            title.Font.Bold = true;

            //保存风速时程button
            System.Windows.Forms.Button btnSaveHistoryPictures = new System.Windows.Forms.Button()
            {
                Text = "",
                FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                BackgroundImage = Properties.Resources.save1,
                BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
            };
            btnSaveHistoryPictures.FlatAppearance.BorderSize = 0;
            btnSaveHistoryPictures.Click += BtnSaveHistoryPictures_Click;

            //生成报告附录button
            System.Windows.Forms.Button btnProduceAppendix = new System.Windows.Forms.Button()
            {
                Text = "",
                FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                BackgroundImage = Properties.Resources.Word,
                BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
            };
            btnProduceAppendix.FlatAppearance.BorderSize = 0;
            btnProduceAppendix.Click += BtnProduceAppendix_Click;

            //tooltip控件
            System.Windows.Forms.ToolTip tip = new System.Windows.Forms.ToolTip();
            tip.SetToolTip(btnSaveHistoryPictures, "保存时程图片");
            tip.SetToolTip(btnProduceAppendix, "生成报告附录");

            //包含俩控件的Panel
            System.Windows.Forms.Panel panel = new System.Windows.Forms.Panel()
            {
                BackColor = Color.FromArgb(242, 242, 242)
            };
            panel.Controls.Add(btnSaveHistoryPictures);
            panel.Controls.Add(btnProduceAppendix);
            xlWs.Controls.AddControl(panel, xlRng.Cells[1, 1], "panel1");

            //控件尺寸位置
            btnSaveHistoryPictures.Height = panel.Height;
            btnSaveHistoryPictures.Width = btnSaveHistoryPictures.Height;
            btnProduceAppendix.Height = panel.Height;
            btnProduceAppendix.Width = btnProduceAppendix.Height;
            btnProduceAppendix.Left = btnSaveHistoryPictures.Width;


            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void BtnSaveHistoryPictures_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string folder = fbd.SelectedPath;
                Range xlRng = xlWs.Columns["H:XFD"];
                xlRng = xlRng.Cells[2, 1].CurrentRegion;
                HistoryPictures hp = new HistoryPictures(folder, xlWs, xlRng);
                if (hp.ProducePictures())
                {
                    XtraMessageBox.Show("生成完毕！", "提示");
                }
            }
        }

        private void BtnProduceAppendix_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.Filter = "所有图片|*.emf;*.wmf;*.jpg;*.jpeg;*.jfif;*.jpe;*.png;*.bmp;*.dib;*.rle;*.emz;*.wmz;*.pcz;*.tif;*.tiff;*.svg;*.eps;*.pct;*.pict;*.wpg";
            ofd.Multiselect = true;
            ofd.FileName = "";
            ofd.Title = "选择时程图片";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ReportAppendix ra = new ReportAppendix(ofd.FileNames);
                if (ra.ProduceAppendix())
                {
                    XtraMessageBox.Show("生成完毕！", "提示");
                }
            }
        }
    }
}
