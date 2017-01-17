using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace WindAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnInsertCoordinate_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane pane;
            pane = Globals.ThisAddIn.CustomTaskPanes.Add(new UserControl_InsertCoordinate(), "插入坐标", Globals.ThisAddIn.Application.ActiveWindow);
            pane.Width = 250;
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            pane.Visible = true;
        }
    }
}
