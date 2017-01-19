using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindAddin
{
    public static class ThePanes
    {
        public static List<Pane> panes = new List<Pane>();
        public class Pane
        {
            public long hwnd;
            public Microsoft.Office.Tools.CustomTaskPane pane;
            public string title;
            public RibbonToggleButton button;
        }
        public static void AddPane(Microsoft.Office.Tools.CustomTaskPane pane, RibbonToggleButton button)
        {
            Pane p = new Pane();
            p.hwnd = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            pane.VisibleChanged += Pane_VisibleChanged;
            p.pane = pane;
            p.title = pane.Title;
            p.button = button;
            panes.Add(p);
        }
        private static void Pane_VisibleChanged(object sender, EventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane pane = (Microsoft.Office.Tools.CustomTaskPane)sender;
            long hwnd = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
            if (pane.Visible)
            {
                Pane[] ps = GetPanes(hwnd);
                if (ps != null)
                {
                    for (int i = 0; i < ps.Length; i++)
                    {
                        if (ps[i].title != pane.Title)
                        {
                            ps[i].pane.Visible = false;
                        }
                    }
                }
            }
            GetPane(hwnd, pane.Title).button.Checked = pane.Visible;
        }

        public static Pane[] GetPanes(long hwnd)
        {
            List<Pane> ps = new List<Pane>();
            foreach (Pane p in panes)
            {
                if (p.hwnd == hwnd)
                {
                    ps.Add(p);
                }
            }
            if (ps.Count == 0)
            {
                return null;
            }
            else
            {
                return ps.ToArray();
            }
        }

        public static Pane GetPane(long hwnd, string title)
        {
            Pane p=null;
            Pane[] ps = GetPanes(hwnd);
            if (ps != null)
            {
                for (int i = 0; i < ps.Length; i++)
                {
                    if (ps[i].title == title)
                    {
                        p = ps[i];
                        break;
                    }
                }
            }
            return p;
        }


        public static bool Contains(long hwnd)
        {
            bool flag = false;
            foreach (Pane p in panes)
            {
                if (p.hwnd == hwnd)
                {
                    flag = true;
                }
            }
            return flag;
        }
    }
}
