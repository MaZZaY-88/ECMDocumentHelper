using System.Collections.Generic;

namespace ECMDocumentHelper.Services
{
    public interface IPdfProcessingService
    {
        string ConvertFilesToMergedPdf(List<string> filePaths);
    }
}