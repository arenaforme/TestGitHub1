using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.ComponentModel;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {

        internal Microsoft.Office.Tools.CustomTaskPane helpTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            helpTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new UserControl2(), "Excel Help");
            helpTaskPane.Visible = true;
            helpTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
        }
        
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
