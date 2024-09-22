using System.Runtime.InteropServices;

namespace ECMDocumentHelper.Helpers
{
    public class PrinterHelper
    {
        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string Name);

        public static void SetPDFCreatorAsDefault()
        {
            SetDefaultPrinter("PDFCreator");
        }
    }
}
