using System;
using System.Threading;
using Microsoft.Office.Tools;
using Post.Office;
using Ribbon.Vsto;
using OfficeCore = global::Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Post
{
    public partial class ThisAddIn
    {
        private VstoHostRuntime _runtime;
        private RibbonSidebarControl _sidebarControl;
        private CustomTaskPane _sidebarPane;
        private Outlook.Explorers _explorers;

        protected override OfficeCore.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new PostRibbon(this);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var synchronizationContext = Ribbon.Vsto.OfficeDispatcher.CaptureCurrentContext();
            _runtime = new VstoHostRuntime(new PostOfficeHost(this.Application, synchronizationContext), synchronizationContext);

            _sidebarControl = new RibbonSidebarControl(_runtime);
            try
            {
                _explorers = Application.Explorers;
                _explorers.NewExplorer += OnNewExplorer;
            }
            catch
            {
                // Without explorer events the pane attaches to whichever window is active now.
            }
            EnsureSidebarPane();
        }

        private void EnsureSidebarPane()
        {
            if (_sidebarPane != null || _sidebarControl == null) return;
            Outlook.Explorer explorer = null;
            try
            {
                explorer = Application.ActiveExplorer();
                if (explorer == null) return;
                _sidebarPane = this.CustomTaskPanes.Add(
                    _sidebarControl,
                    RibbonProductIdentity.GetTaskPaneTitle("Outlook"),
                    explorer);
                _sidebarPane.Width = 420;
                _sidebarPane.Visible = true;
            }
            catch
            {
                _sidebarPane = null;
            }
            finally
            {
                if (explorer != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(explorer);
            }
        }

        private void OnNewExplorer(Outlook.Explorer explorer)
        {
            // The task pane is bound to the first available explorer window; later
            // windows surface it through the Ribbon button once created.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(explorer);
            EnsureSidebarPane();
        }

        internal void ShowSidebar()
        {
            if (_sidebarPane != null) _sidebarPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_explorers != null)
            {
                try { _explorers.NewExplorer -= OnNewExplorer; } catch { }
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(_explorers); } catch { }
                _explorers = null;
            }

            if (_sidebarPane != null)
            {
                _sidebarPane.Visible = false;
                _sidebarPane = null;
            }

            if (_sidebarControl != null)
            {
                _sidebarControl.Dispose();
                _sidebarControl = null;
            }

            if (_runtime != null)
            {
                _runtime.Dispose();
                _runtime = null;
            }
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
