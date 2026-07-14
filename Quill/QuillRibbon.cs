using System.Runtime.InteropServices;
using OfficeCore = global::Microsoft.Office.Core;

namespace Quill
{
    [ComVisible(true)]
    public sealed class QuillRibbon : OfficeCore.IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;

        internal QuillRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""RibbonQuillGroup"" label=""Ribbon"">
          <button id=""RibbonQuillOpenPane"" label=""Open Quill"" onAction=""OpenPane""
                  screentip=""Open Ribbon Quill""
                  supertip=""Show the Ribbon Quill agent workspace for Word."" />
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
