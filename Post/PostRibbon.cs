using System;
using System.Runtime.InteropServices;
using OfficeCore = global::Microsoft.Office.Core;

namespace Post
{
    [ComVisible(true)]
    public sealed class PostRibbon : OfficeCore.IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;

        internal PostRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonId)
        {
            if (!string.Equals(ribbonId, "Microsoft.Outlook.Explorer", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""RibbonPostGroup"" label=""Ribbon"">
          <button id=""RibbonPostOpenPane"" label=""Open Post"" onAction=""OpenPane""
                  screentip=""Open Ribbon Post""
                  supertip=""Show the Ribbon Post agent workspace for Outlook."" />
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
