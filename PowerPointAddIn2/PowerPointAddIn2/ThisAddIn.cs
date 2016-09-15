using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows;
using Microsoft.Kinect;
using System.Drawing;
//using System.Windows.Controls;

namespace PowerPointAddIn2
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e) //, System.Windows.Forms.PaintEventArgs e2)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }



        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
