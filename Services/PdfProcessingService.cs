using ECMDocumentHelper.Helpers;
using System;
using System.Collections.Generic;
using System.Security.Cryptography;

namespace ECMDocumentHelper.Services
{
    public class PdfProcessingService : IPdfProcessingService
    {
        private readonly OfficeInteropHelper _officeInteropHelper;

        public PdfProcessingService(OfficeInteropHelper officeInteropHelper)
        {
            _officeInteropHelper = officeInteropHelper;
        }

        // Method to convert a list of files to a single merged PDF
        public string ConvertFilesToMergedPdf(List<string> filePaths)
        {
            var pdfFiles = new List<string>();

            try
            {
                foreach (var filePath in filePaths)
                {
                    var extension = System.IO.Path.GetExtension(filePath).ToLower();
                    string pdfFile = null;

                    // Convert each file based on its extension
                    switch (extension)
                    {
                        case ".doc":
                        case ".docx":
                            pdfFile = _officeInteropHelper.ConvertWordToPdf(filePath);
                            break;
                        case ".xls":
                        case ".xlsx":
                            pdfFile = _officeInteropHelper.ConvertExcelToPdf(filePath);
                            break;
                        case ".ppt":
                        case ".pptx":
                            pdfFile = _officeInteropHelper.ConvertPowerPointToPdf(filePath);
                            break;
                        case ".msg":
                            pdfFile = _officeInteropHelper.ConvertOutlookMsgToPdf(filePath);
                            break;
                        case ".pdf":
                            pdfFile = filePath;
                            break;
                        default:
                            throw new NotSupportedException($"Unsupported file format: {filePath}");
                    }

                    pdfFiles.Add(pdfFile);
                }

                // Merge all converted PDFs into a single file and return full path
                var fullOutputPath = _officeInteropHelper.MergePdfFiles(pdfFiles, $"{Guid.NewGuid()}.pdf");
                return fullOutputPath;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Error during PDF processing", ex);
            }
        }

        // Method to imprint a barcode on a PDF
        public (int StatusCode, string Message, string OutputPath) ImprintBarcodeOnPdf(string filePath, string barcodeText, string regNumber)
        {
            try
            {

                var outputFilePath = _officeInteropHelper.ImprintBarcodeOnPdf(filePath, barcodeText, regNumber);
                return (1, "Barcode imprinted successfully.", outputFilePath);
            }
            catch (Exception ex)
            {

                return (0, "Error imprinting barcode on PDF.", ex.Message);
            }
        }
    }
}
