using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools.Word.Extensions;

namespace WordDocAnalyzerAddin
{


    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //form = new DocAnalyzerControl();
            //Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane =
                    //this.CustomTaskPanes.Add(form, "Font Analyzer");
            //myCustomTaskPane.Visible = true;

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        //private DocAnalyzerControl form;


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }


}
