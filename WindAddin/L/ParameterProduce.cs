using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindAddin
{
    class ParameterProduce
    {
        public  ParameterProduce()
        {

        }

        public string OneGenerateRec(params string[] para)
        {
            string cellPara = "";
            if (para.Length == 6)       //应该是动态数组改为count
            {
                string z1 = para[0];
                string z2 = para[1];
                string x1 = para[2];
                string y1 = para[3];
                string x2 = para[4];
                string y2 = para[5];
                cellPara = "Z(" + z1 + "," + z2 + ")A(" + x1 + "," + y1 + ")B(" + x2 + "," + y2 + ")";
                //在click中写入excel表格
            }

            if (para.Length == 7)      //水平或者垂直区域，动态数组改为count
            {
                string z1 = para[0];
                string z2 = para[1];
                string x1 = para[2];
                string y1 = para[3];
                string x2 = para[4];
                string y2 = para[5];
                string precision = para[6];
                cellPara = "Z(" + z1 + "," + z2 + ")A(" + x1 + "," + y1 + ")B(" + x2 + "," + y2 + ")" + precision;

            }
            if (para.Length == 12)
            {
                string x0 = para[0];
                string y0 = para[1];
                string z0 = para[2];
                string x1 = para[3];
                string y1 = para[4];
                string z1 = para[5];
                string x2 = para[6];
                string y2 = para[7];
                string z2 = para[8];
                string x3 = para[9];
                string y3 = para[10];
                string z3 = para[11];
                cellPara = "O(" + x0 + "," + y0 + "," + z0 + ")A(" + x1 + "," + y1 + "," + z1 + ")B(" + x2 + "," + y2 + "," + z2 + ")C(" + x3 + "," + y3 + "," + z3 + ")";

            }
            return cellPara;
        }

    }
}
