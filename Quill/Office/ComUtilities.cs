using System.Runtime.InteropServices;

namespace Quill.Office
{
    internal static class ComUtilities
    {
        public static void TryRelease(object value)
        {
            try
            {
                if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
            }
            catch
            {
            }
        }
    }
}
