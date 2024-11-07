namespace ECMDocumentHelper.Models
{
    public class ImageRequest
    {
        public string FilePath { get; set; }  // Path to the file on which the barcode should be imprinted
        public string Tag { get; set; }
        public string ImagePath { get; set; }
    }
}