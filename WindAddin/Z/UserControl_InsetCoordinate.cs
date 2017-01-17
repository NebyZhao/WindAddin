using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace WindAddin
{
    public partial class UserControl_InsetCoordinate : UserControl
    {
        public UserControl_InsetCoordinate()
        {
            InitializeComponent();
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            Worksheet xlSh = Globals.ThisAddIn.Application.ActiveSheet;
            Range c = Globals.ThisAddIn.Application.ActiveCell;
            Range rng = xlSh.Range[c, c.Offset[0, 0]];
            ParameterProduce pp = new ParameterProduce();
            rng.Value = pp.OneGenerateRec(txtZ1.Text, txtZ2.Text, txtX1.Text, txtY1.Text, txtX2.Text, txtY2.Text);
        }
    }
}
