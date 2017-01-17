using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindAddin.L
{
    class ParameterProduce
    {
        public  ParameterProduce()
        {

        }

        public void OneGenerateRec(decimal[] para)
        {
            if (para.Length == 6)       //应该是动态数组改为count
            {
                decimal z1 = para[0];
                decimal z2 = para[1];
                decimal x1 = para[2];
                decimal y1 = para[3];
                decimal x2 = para[4];
                decimal y2 = para[5];
                string cellPara = "Z(" + z1 + "," + z2 + ")A(" + x1 + "," + y1 + ")B(" + x2 + "," + y2 + ")";
                //在click中写入excel表格
            }

            if (para.Length == 7)      //水平或者垂直区域，动态数组改为count
            {
                decimal z1 = para[0];
                decimal z2 = para[1];
                decimal x1 = para[2];
                decimal y1 = para[3];
                decimal x2 = para[4];
                decimal y2 = para[5];
                decimal precision = para[6];
                string cellPara = "Z(" + z1 + "," + z2 + ")A(" + x1 + "," + y1 + ")B(" + x2 + "," + y2 + ")" + precision;

            }
            if (para.Length == 12)
            {
                decimal x0 = para[0];
                decimal y0 = para[1];
                decimal z0 = para[2];
                decimal x1 = para[3];
                decimal y1 = para[4];
                decimal z1 = para[5];
                decimal x2 = para[6];
                decimal y2 = para[7];
                decimal z2 = para[8];
                decimal x3 = para[9];
                decimal y3 = para[10];
                decimal z3 = para[11];
                string cellPara = "O(" + x0 + "," + y0 + "," + z0 + ")A(" + x1 + "," + y1 + "," + z1 + ")B(" + x2 + "," + y2 + "," + z2 + ")C(" + x3 + "," + y3 + "," + z3 + ")";

            }
        }

    }
}
