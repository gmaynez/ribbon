using System.Runtime.InteropServices;
using OfficeCore = global::Microsoft.Office.Core;

namespace Deck
{
    [ComVisible(true)]
    public sealed class DeckRibbon : OfficeCore.IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;

        internal DeckRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""RibbonDeckGroup"" label=""Ribbon"">
          <button id=""RibbonDeckOpenPane"" label=""Open Deck"" onAction=""OpenPane""
                  screentip=""Open Ribbon Deck""
                  supertip=""Show the Ribbon Deck agent workspace for PowerPoint."" />
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
