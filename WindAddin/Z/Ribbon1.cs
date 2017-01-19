using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace WindAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.Application.WindowActivate += Globals.Ribbons.Ribbon1.Application_WindowActivate;
        }

        public void Application_WindowActivate(Workbook Wb, Window Wn)
        {
            foreach (RibbonGroup group in tab1.Groups)
            {
                foreach (RibbonControl control in group.Items)
                {
                    if (control is RibbonToggleButton)
                    {
                        if (ThePanes.GetPane(Wn.Hwnd, ((RibbonToggleButton)control).Label) != null)
                        {
                            ((RibbonToggleButton)control).Checked = ThePanes.GetPane(Wn.Hwnd, ((RibbonToggleButton)control).Label).button.Checked;
                        }
                        else
                        {
                            ((RibbonToggleButton)control).Checked = false;
                        }
                    }
                }
            }
        }


        #region 按钮的单击事件
        private void AddPane(UserControl control, object sender)
        {
            RibbonToggleButton button = (RibbonToggleButton)sender;
            long hwnd = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
            if (ThePanes.GetPane(hwnd, button.Label)==null)
            {
                Microsoft.Office.Tools.CustomTaskPane pane = Globals.ThisAddIn.CustomTaskPanes.Add(control, button.Label, Globals.ThisAddIn.Application.ActiveWindow);
                ThePanes.AddPane(pane, button);
                pane.Width = 250;
            }
            ThePanes.GetPane(hwnd, button.Label).pane.Visible = button.Checked;
        }

        private void btnInsertCoordinate_Click(object sender, RibbonControlEventArgs e)
        {
            AddPane(new UserControl_InsertCoordinate(), sender);
        }
        private void btnExtractData_Click(object sender, RibbonControlEventArgs e)
        {
            AddPane(new UserControl_ExtractData(), sender);
        }
        #endregion

        private void btnWindHistory_Click(object sender, RibbonControlEventArgs e)
        {
            WindHistorySheet.CreateSheet();
        }
    }
}
