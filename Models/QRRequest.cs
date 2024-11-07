namespace ECMDocumentHelper.Models
{
    public class QRRequest
    {
        public string FilePath { get; set; }  // Path to the file on which the barcode should be imprinted
        public string QRText { get; set; }  // The text to be converted to a barcode and imprinted
    }
}
