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
            UserControl_InsetCoordinate control = new UserControl_InsetCoordinate();
            pane = Globals.ThisAddIn.CustomTaskPanes.Add(control, "插入坐标", Globals.ThisAddIn.Application.ActiveWindow);
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            pane.Visible = true;
        }
    }
}
