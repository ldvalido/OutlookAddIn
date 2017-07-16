using System;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace petOutlook
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            const string barNameKey = "myBar";

            CommandBar cmd = null;
            CommandBar cmdTmp = null;
            for (int i = 1;i<= Application.ActiveExplorer().CommandBars.Count;i++)
            {
                cmdTmp = Application.ActiveExplorer().CommandBars[i];
                if (cmdTmp.Name.Equals(barNameKey))
                {
                    cmd = cmdTmp;
                    break;
                }
            }
            
            if (cmd == null)
            {
                CommandBar cmdtmp = Application.ActiveExplorer().CommandBars.Add(
                "myBar",
                MsoBarPosition.msoBarTop,
                missing,
                missing);
            }
            cmd.Visible = true;

            CommandBarButton ctr = (CommandBarButton)cmd.Controls.Add(
                MsoControlType.msoControlButton,
                1,
                "Name", 
                this.missing, 
                true
                );
            ctr.Caption = "Globant Button";
            ctr.Click += ctr_Click;
        }

        void ctr_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MessageBox.Show("I'm here");
            Ctrl.Click += ctr_Click;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

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
