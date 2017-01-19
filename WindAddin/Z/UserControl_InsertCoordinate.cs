using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using DevExpress.XtraEditors;
using System.Text.RegularExpressions;

namespace WindAddin
{
    public partial class UserControl_InsertCoordinate : UserControl
    {
        private Worksheet xlSh { get { return Globals.ThisAddIn.Application.ActiveSheet; } }            //当前工作表
        private Range Cell { get { return Globals.ThisAddIn.Application.Application.ActiveCell; } }     //当前选中的单元格
        public UserControl_InsertCoordinate()
        {
            InitializeComponent();
        }
        
        private void btnInsertCoordinate_One_Click(object sender, EventArgs e)
        {
            ParameterProduce pp = new ParameterProduce();
            try
            {
                Cell.Value = pp.OneWrite(txt_One_A.Text, txt_One_B.Text, txt_One_Z1.Text, txt_One_Z2.Text, txt_One_Precision.Text);
            }
            catch (Exception E)
            {
                XtraMessageBox.Show(E.Message, "错误");
            }
        }

        private void btnInsertCoordinate_ContinueX_Click(object sender, EventArgs e)
        {
            txt_Continue_P.Text = SelectConituousCoodinate();
        }

        private void btnInsertCoordinate_ContinueZ_Click(object sender, EventArgs e)
        {
            txt_Continue_Z.Text = SelectConituousCoodinate();
        }

        private string SelectConituousCoodinate()
        {
            string str = "";
            Range xlRng = Globals.ThisAddIn.Application.Application.Selection;
            if (xlRng.Rows.Count == 1)
            {
                str =xlSh.Name+"!"+ xlRng.Address;
            }
            else if(xlRng.Columns.Count == 1)
            {
                str = xlSh.Name + "!" + xlRng.Address;
            }
            else
            {
                XtraMessageBox.Show("选择区域必须为列或行","提示");
            }
            return str;
        }

        private void btnInsertCoordinate_Continue_Click(object sender, EventArgs e)
        {
            Range RngP, RngZ;
            RngP = Globals.ThisAddIn.Application.Range[txt_Continue_P.Text];
            RngZ = Globals.ThisAddIn.Application.Range[txt_Continue_Z.Text];
            List<string> p = new List<string>();
            List<string> z = new List<string>();
            foreach(Range cell in RngP)
            {
                p.Add(cell.Value.ToString());
            }
            foreach (Range cell in RngZ)
            {
                z.Add(cell.Value.ToString());
            }
            Range rng = xlSh.Range[Cell, Cell.Offset[z.Count - 2, p.Count - 2]];
            ParameterProduce pp = new ParameterProduce();
            try
            {
                rng.Value = pp.ContinuousWrite(p.ToArray(), z.ToArray(), txt_Continue_Precision.Text);
            }
            catch (Exception E)
            {
                XtraMessageBox.Show(E.Message, "错误");
            }            
        }

        private void btnInsertCoordinate_Divide_Click(object sender, EventArgs e)
        {
            int widthCount, heightCount;
            widthCount = int.Parse(txt_Divide_PCount.Text);
            heightCount= int.Parse(txt_Divide_ZCount.Text);
            MatchCollection mc1 = Regex.Matches(txt_Divide_A.Text, @"-?\d+(\.\d+)?");
            MatchCollection mc2 = Regex.Matches(txt_Divide_B.Text, @"-?\d+(\.\d+)?");
            decimal x1, y1, x2, y2,z1,z2;
            string[] p =new string[widthCount+1];
            string[] z = new string[heightCount+1];
            x1 = decimal.Parse(mc1[0].ToString());
            y1 = decimal.Parse(mc1[1].ToString());
            x2 = decimal.Parse(mc2[0].ToString());
            y2 = decimal.Parse(mc2[1].ToString());
            z1 = decimal.Parse(txt_Divide_Z1.Text);
            z2 = decimal.Parse(txt_Divide_Z2.Text);
            for (int i = 0; i <= widthCount; i++)
            {
                decimal stepX, stepY;
                stepX = (x2 - x1) / widthCount;
                stepY = (y2 - y1) / widthCount;
                p[i] = (x1 + i * stepX) + "," + (y1 + i * stepY);
            }
            for(int i = 0; i <= heightCount; i++)
            {
                decimal step = (z2 - z1) / heightCount;
                z[i] = (z1 + i * step).ToString();
            }
            Range rng = xlSh.Range[Cell, Cell.Offset[heightCount-1, widthCount-1]];
            ParameterProduce pp = new ParameterProduce();
            rng.Value = pp.ContinuousWrite(p, z, txt_Divide_Precision.Text);
        }
    }
}
