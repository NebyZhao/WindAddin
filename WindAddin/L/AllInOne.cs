using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace WindAddin.L
{
    class AllInOne
    {
        public string[] angle;
        private string[] files;

        public AllInOne(Microsoft.Office.Interop.Excel.Workbook xlWb,Microsoft.Office.Interop.Excel.Worksheet xlWs,string[] files)
        {
            this.files = files;


            for (int i = 0; i < files.Length; i++)
            {
                DataFile DaF = new DataFile(files[i]);
                FuckExcel FE = new FuckExcel(DaF, xlWb, xlWs);
                FE.ManageExcel();
            }


        }
    }
}
