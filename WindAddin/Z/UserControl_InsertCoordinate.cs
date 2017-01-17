using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace WindAddin
{
    public partial class UserControl_InsertCoordinate : UserControl
    {
        public UserControl_InsertCoordinate()
        {
            InitializeComponent();
        }

        private void btnInsertCoordinate_Click(object sender, EventArgs e)
        {
            Worksheet xlSh = Globals.ThisAddIn.Application.Application.ActiveSheet;
            Range cell = Globals.ThisAddIn.Application.Application.ActiveCell;
            Range xlRng = xlSh.Range[cell, cell.Offset[0, 0]];
            ParameterProduce pp = new ParameterProduce();
            xlRng.Value = pp.OneGenerateRec(txtZ1.Text, txtZ2.Text, txtX1.Text, txtY1.Text, txtX2.Text, txtY2.Text);
        }
    }
}
