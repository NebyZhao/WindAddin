using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindAddin
{
    class ParameterProduce
    {
        public string OneWrite(string A, string B, string z1, string z2, string precision = "")
        {

            return OneGenerateRec(A, B, z1, z2, precision);

        }

        public string[,] ContinuousWrite(string[] P, string[] Z, string precision = "")
        {
            string[,] para = new string[Z.Length - 1, P.Length - 1];
            for (int i1 = 0; i1 < P.Length - 1; i1++)
            {
                for (int i2 = 0; i2 < Z.Length - 1; i2++)
                {
                    para[i2, i1] = OneGenerateRec(P[i1], P[i1 + 1], Z[i2], Z[i2 + 1], precision);
                }
            }
            return para;
        }



        private string OneGenerateRec(string A, string B, string z1, string z2, string precision ="")
        {
            string cellPara = "";
            string splitA = null, splitB = null;
            if (A.Contains(","))
            {
                splitA = ",";
            }
            else if (A.Contains("，"))
            {
                splitA = "，";
            }
            if (B.Contains(","))
            {
                splitB = ",";
            }
            else if (B.Contains("，"))
            {
                splitB = "，";
            }
            string x1, y1, x2, y2;
            if (splitA != null && splitB != null)
            {
                x1 = A.Split(splitA.ToCharArray())[0];
                y1 = A.Split(splitA.ToCharArray())[1];
                x2 = B.Split(splitB.ToCharArray())[0];
                y2 = B.Split(splitB.ToCharArray())[1];
                cellPara = "Z(" + z1 + "," + z2 + ")A(" + x1 + "," + y1 + ")B(" + x2 + "," + y2 + ")";
            }
            else
            {
                throw new Exception("分隔符格式不正确");
            }
            if (precision!="")
            {
                cellPara += precision;
            }
            return cellPara;

            //if (para.Length == 12)
            //{
            //    string x0 = para[0];
            //    string y0 = para[1];
            //    string z0 = para[2];
            //    string x1 = para[3];
            //    string y1 = para[4];
            //    string z1 = para[5];
            //    string x2 = para[6];
            //    string y2 = para[7];
            //    string z2 = para[8];
            //    string x3 = para[9];
            //    string y3 = para[10];
            //    string z3 = para[11];
            //    cellPara = "O(" + x0 + "," + y0 + "," + z0 + ")A(" + x1 + "," + y1 + "," + z1 + ")B(" + x2 + "," + y2 + "," + z2 + ")C(" + x3 + "," + y3 + "," + z3 + ")";

            //}
        }

    }
}
