using ECMDocumentHelper.Helpers;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ECMDocumentHelper.Services
{
    public class PdfProcessingService
    {
        private readonly OfficeInteropHelper _officeInteropHelper;
        private readonly BarcodeHelper _barcodeHelper;
        private readonly ImageHelper _imageHelper;
        private readonly ILogger<PdfProcessingService> _logger;

        public PdfProcessingService(OfficeInteropHelper officeInteropHelper, BarcodeHelper barcodeHelper, ImageHelper imageHelper, ILogger<PdfProcessingService> logger)
        {
            _officeInteropHelper = officeInteropHelper;
            _barcodeHelper = barcodeHelper;
            _imageHelper = imageHelper;
            _logger = logger;
        }

        public async Task<(int StatusCode, string Message, string OutputPath)> GeneratePDFAsync(List<string> filePaths)
        {
            string outputFilePath = string.Empty;
            try
            {
                foreach (var filePath in filePaths)
                {
                    string extension = Path.GetExtension(filePath)?.ToLowerInvariant();
                    outputFilePath = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                    switch (extension)
                    {
                        case ".docx":
                        case ".doc":
                            await _officeInteropHelper.ConvertWordToPdfAsync(filePath, outputFilePath);
                            break;

                        case ".pptx":
                        case ".ppt":
                            await _officeInteropHelper.ConvertPowerPointToPdfAsync(filePath, outputFilePath);
                            break;

                        case ".msg":
                            await _officeInteropHelper.ConvertOutlookMsgToPdfAsync(filePath, outputFilePath);
                            break;

                        case ".png":
                        case ".jpg":
                        case ".jpeg":
                            await _imageHelper.ConvertImageToPdfAsync(filePath, outputFilePath);
                            break;

                        default:
                            _logger.LogWarning("Unsupported file format: {FilePath}", filePath);
                            return (0, "Unsupported file format", null);
                    }
                }

                return (1, "Files processed successfully.", outputFilePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing file: {FilePath}", outputFilePath);
                return (0, "Error processing files", null);
            }
        }

        public async Task<(int StatusCode, string Message, string OutputPath)> ImprintBarcodeOnPdfAsync(string filePath, string barcodeText)
        {
            return await _barcodeHelper.ImprintBarcodeOnPdf(filePath, barcodeText);                
        }
    }
}
