namespace ECMDocumentHelper.Models
{
    public class BarcodeRequest
    {
        public string FilePath { get; set; }  // Path to the file on which the barcode should be imprinted
        public string BarcodeText { get; set; }  // The text to be converted to a barcode and imprinted
        public string RegNumber { get; set; }  

    }
}
