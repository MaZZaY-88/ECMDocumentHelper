
using ECMDocumentHelper.Helpers;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;

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

        public (int StatusCode, string Message, string OutputPath) GeneratePDF(List<string> filePaths)
        {
            //foreach (var filePath in filePaths)
            //{
            //    try
            //    {
            //        string extension = Path.GetExtension(filePath)?.ToLowerInvariant();
            //        string outputFilePath = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + ".pdf");

            //        switch (extension)
            //        {
            //            case ".docx":
            //            case ".doc":
            //                _officeInteropHelper.ConvertWordToPdf(filePath, outputFilePath);
            //                break;

            //            case ".pptx":
            //            case ".ppt":
            //                _officeInteropHelper.ConvertPowerPointToPdf(filePath, outputFilePath);
            //                break;

            //            case ".msg":
            //                _officeInteropHelper.ConvertOutlookMsgToPdf(filePath, outputFilePath);
            //                break;

            //            case ".png":
            //            case ".jpg":
            //            case ".jpeg":
            //                _imageHelper.ConvertImageToPdf(filePath, outputFilePath);
            //                break;

            //            default:
            //                _logger.LogWarning("Unsupported file format: {FilePath}", filePath);
            //                return (0, "Unsupported file format.", null);
            //        }

            //        return (1, "Files processed successfully.", outputFilePath);
            //    }
            //    catch (Exception ex)
            //    {
            //        _logger.LogError(ex, "Error processing file: {FilePath}", filePath);
            //        return (0, $"Error processing file: {filePath}", null);
            //    }
            //}
            return (1, "Files processed successfully.", null);
        }

        public (int StatusCode, string Message, string OutputPath) ImprintBarcodeOnPdf(string filePath, string barcodeText, string regNumber)
        {
            try
            {
                var outputFilePath = _barcodeHelper.ImprintBarcodeOnPdf(filePath, barcodeText, regNumber);
                return (1, "Barcode imprinted successfully.", outputFilePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error imprinting barcode on PDF: {FilePath}", filePath);
                return (0, "Error imprinting barcode on PDF.", ex.Message);
            }
        }
    }
}
