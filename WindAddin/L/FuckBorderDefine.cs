using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WindAddin.L
{
    public class FuckBorderDefine
    {
        public decimal x1 { private set; get; }
        public decimal x2 { private set; get; }
        public decimal y1 { private set; get; }
        public decimal y2 { private set; get; }
        public decimal z1 { private set; get; }
        public decimal z2 { private set; get; }
        public decimal precision { private set; get; }

        public decimal x0 { private set; get; }
        public decimal y0 { private set; get; }
        public decimal z0 { private set; get; }
        public decimal x3 { private set; get; }
        public decimal y3 { private set; get; }
        public decimal z3 { private set; get; }

        public bool IsMatch { private set; get; }
        //public bool IsMatch { private set; get; }//////////////////////////////
        /// <summary>
        /// 
        /// </summary>
        /// <param name="info">excel某个单元格里面的内容</param>
        public FuckBorderDefine(string info)
        {
            Match(info);
        }


        private void Match(string info)
        {
            Regex regex = new Regex(@"Z[(（]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[)）]A[(（]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[)）]B[(（]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[)）](\d+(.\d+)?)?", RegexOptions.IgnoreCase);
            Regex regex1 = new Regex(@"[zZ][(（]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[)）](\d+(.\d+)?)?[aA][(（]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[)）](\d+(.\d+)?)?[bB][(（]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[)）](\d+(.\d+)?)?[Cc][(（]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[,，]-?\d+(\.\d+)?[)）](\d+(.\d+)?)?", RegexOptions.IgnoreCase);
            if (regex.IsMatch(info.Trim()))
            {
                MatchCollection mc = Regex.Matches(info, @"-?\d+(\.\d+)?");
                if (mc.Count == 6)
                {
                    z1 = decimal.Parse(mc[0].ToString());
                    z2 = decimal.Parse(mc[1].ToString());
                    x1 = decimal.Parse(mc[2].ToString());
                    y1 = decimal.Parse(mc[3].ToString());
                    x2 = decimal.Parse(mc[4].ToString());
                    y2 = decimal.Parse(mc[5].ToString());
                    precision = 0.001M;
                    IsMatch = true;
                }
                else if (mc.Count == 7)
                {
                    z1 = decimal.Parse(mc[0].ToString());
                    z2 = decimal.Parse(mc[1].ToString());
                    x1 = decimal.Parse(mc[2].ToString());
                    y1 = decimal.Parse(mc[3].ToString());
                    x2 = decimal.Parse(mc[4].ToString());
                    y2 = decimal.Parse(mc[5].ToString());
                    precision = decimal.Parse(mc[6].ToString());
                    IsMatch = true;
                }
            }

            else if (regex1.IsMatch(info.Trim()))
            {
                MatchCollection mc = Regex.Matches(info, @"-?\d+(\.\d+)?");
                x0 = decimal.Parse(mc[0].ToString());
                y0 = decimal.Parse(mc[1].ToString());
                z0 = decimal.Parse(mc[2].ToString());
                x1 = decimal.Parse(mc[3].ToString());
                y1 = decimal.Parse(mc[4].ToString());
                z1 = decimal.Parse(mc[5].ToString());
                x2 = decimal.Parse(mc[6].ToString());
                y2 = decimal.Parse(mc[7].ToString());
                z2 = decimal.Parse(mc[8].ToString());
                x3 = decimal.Parse(mc[9].ToString());
                y3 = decimal.Parse(mc[10].ToString());
                z3 = decimal.Parse(mc[11].ToString());
                IsMatch = true;

            }
            else
            {
                IsMatch = false;
            }

            //excel格式匹配的话就给x1等赋值，并且ismatch=true
            //否则不赋值，且ismatch=false
        }

        internal void IsIncluded(object p)
        {
            throw new NotImplementedException();
        }

        public bool IsIncluded(Point p, decimal precision)        //precision是可选参数，缺省值为0.001
        {
            bool flag = false;
            if (IsMatch)
            {
                //if (z1 != z2)
                //{
                if ((p.z - z1) * (p.z - z2) <= 0)
                {
                    if ((p.x - x1) * (p.x - x2) <= 0 || (p.y - y1) * (p.y - y2) <= 0)
                    {
                        decimal a, b, c;
                        a = y2 - y1;
                        b = x1 - x2;
                        c = y1 * x2 - y2 * x1;
                        if ((a * p.x + b * p.y + c) * (a * p.x + b * p.y + c) <= (a * a + b * b) * precision * precision)
                        {
                            flag = true;
                        }
                    }
                }
                //}

                //else if (z1 == z2)
                //{

                //    if ((p.x - x1) * (p.x - x2) <= 0 || (p.y - y1) * (p.y - y2) <= 0)
                //    {

                //    }

                //}
            }
            return flag;
        }

    }
}
