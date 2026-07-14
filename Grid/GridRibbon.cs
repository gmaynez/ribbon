using System.Runtime.InteropServices;
using OfficeCore = global::Microsoft.Office.Core;

namespace Grid
{
    [ComVisible(true)]
    public sealed class GridRibbon : OfficeCore.IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;

        internal GridRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""RibbonGridGroup"" label=""Ribbon"">
          <button id=""RibbonGridOpenPane"" label=""Open Grid"" onAction=""OpenPane""
                  screentip=""Open Ribbon Grid""
                  supertip=""Show the Ribbon Grid agent workspace for Excel."" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OpenPane(OfficeCore.IRibbonControl control)
        {
            _addIn.ShowSidebar();
        }
    }
}
