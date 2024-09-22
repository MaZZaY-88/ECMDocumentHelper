namespace ECMDocumentHelper.Models
{
    public class MergeFilesRequest
    {
        public List<string> FilePaths { get; set; }  // List of file paths to be merged into a single PDF
    }
}
