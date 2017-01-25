using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace WindAddin
{
    public class HistoryPictures
    {
        private string folderPath;
        private Range xlRng;
        private Microsoft.Office.Tools.Excel.Worksheet xlWs;
        public HistoryPictures(string folder, Microsoft.Office.Tools.Excel.Worksheet xlWs, Range xlRng)
        {
            this.folderPath = folder;
            this.xlWs = xlWs;
            this.xlRng = xlRng;
        }
        public bool ProducePictures()
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Chart chart = xlWs.Shapes.AddChart2(240, XlChartType.xlXYScatterSmoothNoMarkers).Chart;
            chart.ChartArea.Width = 448;
            chart.ChartArea.Height = 224;
            xlWs.Range[xlRng.Cells[3, 1], xlRng.Cells[xlRng.Rows.Count, 1]].Name = "X";
            xlWs.Range[xlRng.Cells[3, 2], xlRng.Cells[xlRng.Rows.Count, 2]].Name = "Y";
            chart.SetSourceData(xlWs.Range["X,Y"]);
            chart.ChartTitle.Text =string.Format("第{0}层风速时程",NumberToChinese(Convert.ToInt32(xlRng.Cells[2,2].Value)));
            chart.ChartTitle.Font.Bold = true;
            chart.ChartTitle.Font.Size = 14;
            chart.HasLegend = false;
            chart.Axes(XlAxisType.xlValue).HasTitle = true;
            chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "速度(m/s)";
            chart.Axes(XlAxisType.xlValue).AxisTitle.Font.Bold = true;
            chart.Axes(XlAxisType.xlValue).AxisTitle.Font.Size = 11;
            chart.Axes(XlAxisType.xlCategory).HasTitle = true;
            chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "时间(s)";
            chart.Axes(XlAxisType.xlCategory).AxisTitle.Font.Bold = true;
            chart.Axes(XlAxisType.xlCategory).AxisTitle.Font.Size = 11;
            chart.Axes(XlAxisType.xlCategory).HasMajorGridlines = false;
            chart.Axes(XlAxisType.xlValue).MajorUnit = 10;
            chart.Axes(XlAxisType.xlCategory).MaximumScale = 400;
            chart.Axes(XlAxisType.xlCategory).MajorUnit = 100;
            chart.Axes(XlAxisType.xlValue).TickLabels.Font.Bold = true;
            chart.Axes(XlAxisType.xlCategory).TickLabels.Font.Bold = true;
            chart.PlotArea.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            chart.PlotArea.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent1;
            chart.Axes(XlAxisType.xlValue).Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent1;
            chart.Axes(XlAxisType.xlCategory).Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent1;
            chart.Axes(XlAxisType.xlValue).MajorGridlines.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent1;
            chart.ChartArea.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            chart.ChartArea.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorText1;
            chart.SeriesCollection(1).Format.Line.Weight = 2;
            chart.Export(string.Format("{0}\\{1}{2}.png", folderPath, xlRng.Cells[2, 2].Text, chart.ChartTitle.Text));
            if (xlRng.Columns.Count > 2)
            {
                for(int i = 3; i <= xlRng.Columns.Count; i++)
                {
                    //xlWs.Range["Y"].Name = "";
                    xlWs.Range[xlRng.Cells[3, i], xlRng.Cells[xlRng.Rows.Count, i]].Name = "Y";
                    chart.SetSourceData(xlWs.Range["X,Y"]);
                    chart.ChartTitle.Text = string.Format("第{0}层风速时程", NumberToChinese(Convert.ToInt32(xlRng.Cells[2, i].Value)));
                    chart.Export(string.Format("{0}\\{1}{2}.png", folderPath, xlRng.Cells[2, i].Text, chart.ChartTitle.Text));
                }
            }
            chart.Parent.Delete();
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            return true;
        }

        private string Num2CHN(int num)
        {
            switch (num)
            {
                case 1:
                    return "一";
                case 2:
                    return "二";
                case 3:
                    return "三";
                case 4:
                    return "四";
                case 5:
                    return "五";
                case 6:
                    return "六";
                case 7:
                    return "七";
                case 8:
                    return "八";
                case 9:
                    return "九";
                default:
                    return "";
            }
        }
        private string NumberToChinese(int num)
        {
            string str = "";
            int h, t, o;
            h = num / 100;
            if (h > 0)
            {
                str = Num2CHN(h) + "百";
                t = num % 100/10;
                o = num % 10;
                if (t > 0)
                {
                    str += Num2CHN(t) + "十" + Num2CHN(o);
                }
                else
                {
                    if (o > 0)
                    {
                        str += "零" + Num2CHN(o);
                    }
                    else
                    {
                        str += Num2CHN(o);
                    }
                }
            }
            else
            {
                t = num / 10;
                o = num % 10;
                if (t > 1)
                {
                    str += Num2CHN(t) + "十" + Num2CHN(o);
                }
                else
                {
                    if (t > 0)
                    {
                        str += "十" + Num2CHN(o);
                    }
                    else
                    {
                        str = Num2CHN(o);
                    }
                }
            }
            return str;
        }
    }
}
