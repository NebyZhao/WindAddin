using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WindAddin.L
{
    public struct Point
    {
        public decimal x;
        public decimal y;
        public decimal z;
        public decimal p;
    }
    public class DataFile
    {
        public decimal angle { private set; get; }
        public Point[] points { private set; get; }



        private string path;
        public int count { private set; get; }
        private List<decimal> data = new List<decimal>();



        public DataFile(string path)
        {
            this.path = path;
            GetData();
            GetResult();
            ProceedResult();
        }

        private void GetData()
        {
            //判断文件名称中的角度，mc是一个数组
            MatchCollection mc = Regex.Matches(path, "\\d+");
            angle = decimal.Parse(mc[mc.Count - 1].ToString());
            //读取文件内容
            string[] content = File.ReadAllLines(path);
            //找到点个数N=？ （N=?在content数组的第三个元素里）
            MatchCollection mc1 = Regex.Matches(content[2], "(?<=N=)\\d+");
            count = int.Parse(mc1[0].ToString());
            //将xyzp读取进data数组
            for (int i = 4; i < content.Length; i++)
            {
                MatchCollection mc2 = Regex.Matches(content[i], "\\S+");
                for (int n = 0; n < mc2.Count; n++)
                {
                    data.Add(Convert.ToDecimal(double.Parse(mc2[n].ToString())));
                }
            }
        }


        private void GetResult()
        {

            points = new Point[count];
            for (int i = 0; i < count; i++)
            {
                points[i].x = data[i];
                points[i].y = data[i + count];
                points[i].z = data[i + 2 * count];
                points[i].p = data[i + 3 * count];
            }
            //count
            //将data[]里的数据分配到x,y,z,p
        }


        private void ProceedResult()
        {
            //变换为笛卡尔坐标系,旋转建筑
            decimal xx, zz, yy;
            double theta = Convert.ToDouble(-angle);
            for (int i = 0; i < count; i++)
            {
                //xx = points[i].x;
                //yy = -points[i].z;
                //zz = points[i].y;
                xx = points[i].x;
                yy = points[i].y;
                zz = points[i].z;

                points[i].x = xx * Convert.ToDecimal(Math.Cos(theta / 180 * Math.PI)) - yy * Convert.ToDecimal(Math.Sin(theta / 180 * Math.PI));
                points[i].y = yy * Convert.ToDecimal(Math.Cos(theta / 180 * Math.PI)) + xx * Convert.ToDecimal(Math.Sin(theta / 180 * Math.PI));
                points[i].z = zz;
            }

        }
    }
}
