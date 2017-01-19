using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace WindAddin
{
    class AllInOne
    {
        public string[] angle;
        private string[] files;

        public AllInOne(Microsoft.Office.Interop.Excel.Workbook xlWb,Microsoft.Office.Interop.Excel.Worksheet xlWs,string[] files)
        {
            this.files = files;
            angle = new string[files.Length];

            for (int i1 = 0; i1 < files.Length; i1++)
            {
                DataFile DaF = new DataFile(files[i1]);
                FuckExcel FE = new FuckExcel(DaF, xlWb, xlWs);
                FE.ManageExcel();
                angle[i1] = DaF.angle.ToString();
            }
                        

        }
    }
}
