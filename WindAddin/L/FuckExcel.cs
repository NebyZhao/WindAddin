using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindAddin.L
{
    class FuckExcel
    {
        private Workbook xlWb;
        private decimal angle;
        private int count;
        private Point[] points;
        private decimal precision;
        private Worksheet xlWs;



        public FuckExcel(DataFile daF, Microsoft.Office.Interop.Excel.Workbook xlWb,Microsoft.Office.Interop.Excel.Worksheet xlWs)
        {
            this.xlWb = xlWb;
            this.xlWs = xlWs;
            this.angle = daF.angle;
            this.count = daF.count;
            this.points = daF.points;
        }

        public void ManageExcel()
        {

            Worksheet xlWs1 = xlWs;
            xlWs1.Copy(After: xlWb.Sheets[xlWb.Sheets.Count]);
            xlWb.Sheets[xlWb.Sheets.Count].Name = angle.ToString();


            int row = xlWs1.UsedRange.Rows.Count;
            int col = xlWs1.UsedRange.Columns.Count;

            Range xlRn = xlWs1.Range[xlWs1.Cells[1, 1], xlWs1.Cells[row, col]];
            object[,] xlRng = xlRn.Value;


            for (int i1 = 0; i1 < row; i1++)                            //循环excel行
            {
                for (int i2 = 0; i2 < col; i2++)                        //循环excel列
                {
                    string info;
                    if (xlRng[i1 + 1, i2 + 1] == null)
                    {
                        continue;
                    }

                    info = xlRng[i1 + 1, i2 + 1].ToString();                 //object数组在这里与excel range一样是从1，1开始的，不是0，0开始的
                    FuckBorderDefine fbd = new FuckBorderDefine(info);


                    if (fbd.IsMatch)                                    //判断excel对应cell是否需要计算p
                    {
                        decimal totalP = 0;
                        int totalCount = 0;
                        for (int i3 = 0; i3 < count; i3++)              //从所有数据中找出满足单元格范围的P
                        {
                            if (fbd.IsIncluded(points[i3], fbd.precision))
                            {
                                totalP += points[i3].p;
                                totalCount++;
                            }
                        }

                        if (totalCount != 0)
                        {
                            decimal aveP = decimal.Round((totalP / totalCount), 2);          //四舍五入到整数
                            xlWb.Sheets[xlWb.Sheets.Count].Cells[i1 + 1, i2 + 1].Value = aveP.ToString();
                        }
                        else
                        {
                            xlWb.Sheets[xlWb.Sheets.Count].Cells[i1 + 1, i2 + 1].Value = "null";
                        }

                        //decimal totalP = 0;
                        //decimal averageP=0;
                        //for (int i4 = 0; i4 < arrayP[i1,i2].Count; i4++)        //单元格范围的P求和
                        //{
                        //    totalP += arrayP[i1, i2][i4];
                        //}
                        //averageP = totalP / arrayP[i1, i2].Count;               //单元格范围的P求均值

                        //xlWb.Sheets[xlWb.Sheets.Count].Cells[i1 + 1, i2 + 1].Value = averageP.ToString();//cells还是cell？？？？？？？？？？？？？
                        ////////////////////////////////////////////////////////////////////////////////////////////////////
                    }

                    xlWb.Sheets[xlWb.Sheets.Count].Columns[i2 + 1].Autofit();
                }

            }
        }
    }
}
