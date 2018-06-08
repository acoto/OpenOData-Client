using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using System.Windows.Forms;

namespace OpenODATA
{
    public partial class ThisAddIn
    {
        private CustomTaskPane taskPane;
        UserControl1 control = new UserControl1();
        internal CustomTaskPane TaskPane
        {
            get
            {
                return this.taskPane;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.taskPane = this.CustomTaskPanes.Add(control, "OData Panel");
            this.taskPane.Visible = false;

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
