using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace WindAddin
{
    public class ReportAppendix
    {
        private string[] pictures;
        public ReportAppendix(string[] pictures)
        {
            this.pictures = pictures;
        }
        public bool ProduceAppendix()
        {
            Microsoft.Office.Interop.Word.Application wdApp = new Microsoft.Office.Interop.Word.Application();
            wdApp.Visible = true;
            Document wdDoc = wdApp.Documents.Add();
            Table table = wdDoc.Tables.Add(wdApp.Selection.Range, Convert.ToInt32(Math.Ceiling(pictures.Length / 2m)) * 2, 2);
            table.Columns[1].Width = 200.0f;
            table.Columns[2].Width = 200.0f;
            table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            Array.Sort(pictures, new Compare());
            for (int i = 0; i < pictures.Length; i++)
            {
                wdDoc.InlineShapes.AddPicture(pictures[i], Range: table.Cell((i / 2) * 2 + 1, i % 2 + 1).Range);
                Regex regex = new Regex(@"(?<=\d+)\D+?(?=\.jpg)");
                string name = regex.Match(pictures[i]).ToString();
                table.Cell((i / 2) * 2 + 2, i % 2 + 1).Range.Text = name;
            }
            return true;
        }
        private class Compare : IComparer<string>
        {
            int IComparer<string>.Compare(string x, string y)
            {
                Regex regex = new Regex(@"\d+");
                int X = int.Parse(regex.Match(x).ToString());
                int Y = int.Parse(regex.Match(y).ToString());
                return X - Y;
            }
        }
    }
}
