using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace WindAddin
{
    /// <summary>
    /// 生成包络值
    /// </summary>
    public class Envelope
    {
        private Workbook xlWb;
        private Worksheet xlWs;
        private string[] angles;

        public Envelope(Workbook xlWb, Worksheet xlWs, string[] angles)
        {
            this.xlWb = xlWb;
            this.xlWs = xlWs;
            this.angles = angles;
        }

        public void ManageExcel()
        {
            xlWs.Copy(After: xlWb.Sheets[xlWb.Sheets.Count]);
            xlWb.Sheets[xlWb.Sheets.Count].Name = "正包络";
            xlWs.Copy(After: xlWb.Sheets[xlWb.Sheets.Count]);
            xlWb.Sheets[xlWb.Sheets.Count].Name = "负包络";

            int row = xlWs.UsedRange.Rows.Count;
            int col = xlWs.UsedRange.Columns.Count;

            Range xlRn = xlWs.Range[xlWs.Cells[1, 1], xlWs.Cells[row, col]];
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
                        decimal[] dat = new decimal[angles.Length];
                        for (int i3 = 0; i3 < angles.Length; i3++)
                        {
                            try
                            {
                                dat[i3] = decimal.Parse(xlWb.Sheets[angles[i3]].Cells[i1 + 1, i2 + 1].Value.ToString());
                            }
                            catch { }
                        }
                        xlWb.Sheets["正包络"].Cells[i1 + 1, i2 + 1].Value = PlusMax(dat);
                        xlWb.Sheets["负包络"].Cells[i1 + 1, i2 + 1].Value = MinusMax(dat);
                        xlWb.Sheets["正包络"].Columns[i2 + 1].Autofit();
                        xlWb.Sheets["负包络"].Columns[i2 + 1].Autofit();
                    }                  
                }
            }
        }

        private decimal PlusMax(decimal[] arr)
        {
            decimal max = 0;
            foreach(decimal d in arr)
            {
                if (d > max)
                {
                    max = d;
                }
            }
            return max;
        }

        private decimal MinusMax(decimal[] arr)
        {
            decimal min = 0;
            foreach (decimal d in arr)
            {
                if (d < min)
                {
                    min = d;
                }
            }
            return min;
        }
    }

   
}
