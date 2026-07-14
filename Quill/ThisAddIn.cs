using System;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Quill.Office;
using Ribbon.Vsto;

namespace Quill
{
    public partial class ThisAddIn
    {
        private VstoHostRuntime _runtime;
        private RibbonSidebarControl _sidebarControl;
        private CustomTaskPane _sidebarPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var context = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();
            _runtime = new VstoHostRuntime(new QuillOfficeHost(this.Application, context), context);
            _sidebarControl = new RibbonSidebarControl(_runtime);
            _sidebarPane = this.CustomTaskPanes.Add(_sidebarControl, "Ribbon Agents");
            _sidebarPane.Width = 420;
            _sidebarPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_sidebarPane != null) { _sidebarPane.Visible = false; _sidebarPane = null; }
            if (_sidebarControl != null) { _sidebarControl.Dispose(); _sidebarControl = null; }
            if (_runtime != null) { _runtime.Dispose(); _runtime = null; }
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
