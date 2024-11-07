using ExcelApp = Microsoft.Office.Interop.Excel;
using OutlookApp = Microsoft.Office.Interop.Outlook;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint;
using WordApp = Microsoft.Office.Interop.Word;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using System.Collections.Generic;
using System.Security.Cryptography;

namespace ECMDocumentHelper.Helpers
{
    public class OfficeInteropHelper
    {
        private readonly string _outputDirectory;
        private readonly string _pdfSaveDirectory;

        public OfficeInteropHelper(IConfiguration configuration)
        {
            _outputDirectory = configuration.GetSection("PdfSettings")["outputDirectory"];
            _pdfSaveDirectory = configuration.GetSection("PdfSettings")["pdfSaveDirectory"];

            if (!Directory.Exists(_outputDirectory))
            {
                Directory.CreateDirectory(_outputDirectory);
            }

            if (!Directory.Exists(_pdfSaveDirectory))
            {
                Directory.CreateDirectory(_pdfSaveDirectory);
            }
        }

        // Method to merge multiple PDF files into a single PDF
        public string MergePdfFiles(List<string> pdfFilePaths, string outputPdfFileName)
        {
            string fullOutputPath = Path.Combine(_outputDirectory, outputPdfFileName);

            using (var outputDocument = new PdfSharp.Pdf.PdfDocument())
            {
                foreach (var pdfFile in pdfFilePaths)
                {
                    using (var inputDocument = PdfSharp.Pdf.IO.PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import))
                    {
                        foreach (var page in inputDocument.Pages)
                        {
                            outputDocument.AddPage(page);
                        }
                    }
                }

                outputDocument.Save(fullOutputPath);
            }

            return fullOutputPath;
        }

        // Method to imprint barcode on each page of a PDF
        public string ImprintBarcodeOnPdf(string inputPdfPath, string barcodeText, string regNumber)
        {
            try
            {
                using (PdfDocument document = PdfReader.Open(inputPdfPath, PdfDocumentOpenMode.Modify))
                {

                    double xPosition = 10;
                    double yPosition = 135;
                    int rotation = -90;

                    // Convert positions from millimeters to points
                    double xPositionPt = XUnit.FromMillimeter(xPosition);
                    double yPositionPt = XUnit.FromMillimeter(yPosition);

                    // Define fonts
                    XFont barcodeFont = new XFont("LibreBarcode128Text", 20);
                    XFont textFont = new XFont("LiberationSans", 8);

                    foreach (PdfPage page in document.Pages)
                    {
                        using (XGraphics gfx = XGraphics.FromPdfPage(page))
                        {
                            // Save the current state of the graphics context
                            gfx.Save();

                            // Apply transformations
                            gfx.TranslateTransform(xPositionPt, yPositionPt);
                            gfx.RotateTransform(rotation);

                            // Draw the barcode
                            gfx.DrawString(barcodeText, barcodeFont, XBrushes.Black, new XPoint(0, 0));

                            // Measure the size of the barcode text
                            XSize barcodeSize = gfx.MeasureString(barcodeText, barcodeFont);

                            // Define a gap between the barcode and the human-readable text
                            double gap = XUnit.FromPoint(10); // 2-point gap

                            // Draw the human-readable text below the barcode
                            gfx.DrawString(regNumber, textFont, XBrushes.Black, new XPoint(0, gap));

                            // Restore the graphics context to its previous state
                            gfx.Restore();
                        }
                    }

                    string guid = Guid.NewGuid().ToString();
                    string outputPdfPath = Path.Combine(Path.GetDirectoryName(inputPdfPath), $"{guid}_barcode.pdf");
                    document.Save(outputPdfPath);
                    return outputPdfPath;
                }
            }
            catch (Exception ex)
            {

                throw new ApplicationException("Error during barcode imprinting", ex);
            }
        }

        public string ConvertWordToPdf(string wordFilePath)
        {
            var wordApp = new WordApp.Application();
            WordApp.Document doc = null;

            var outputPdfPath = Path.Combine(_outputDirectory, $"{Path.GetFileNameWithoutExtension(wordFilePath)}_{Guid.NewGuid()}.pdf");

            try
            {
                doc = wordApp.Documents.Open(wordFilePath);
                doc.ExportAsFixedFormat(outputPdfPath, WordApp.WdExportFormat.wdExportFormatPDF);
            }
            finally
            {
                doc?.Close();
                wordApp.Quit();
            }

            return outputPdfPath;
        }

        public string ConvertExcelToPdf(string excelFilePath)
        {
            var excelApp = new ExcelApp.Application();
            ExcelApp.Workbook workbook = null;

            var outputPdfPath = Path.Combine(_outputDirectory, $"{Path.GetFileNameWithoutExtension(excelFilePath)}_{Guid.NewGuid()}.pdf");

            try
            {
                workbook = excelApp.Workbooks.Open(excelFilePath);
                workbook.ExportAsFixedFormat(ExcelApp.XlFixedFormatType.xlTypePDF, outputPdfPath);
            }
            finally
            {
                workbook?.Close(false);
                excelApp.Quit();
            }

            return outputPdfPath;
        }

        public string ConvertPowerPointToPdf(string pptFilePath)
        {
            var pptApp = new PowerPointApp.Application();
            PowerPointApp.Presentation presentation = null;

            var outputPdfPath = Path.Combine(_outputDirectory, $"{Path.GetFileNameWithoutExtension(pptFilePath)}_{Guid.NewGuid()}.pdf");

            try
            {
                presentation = pptApp.Presentations.Open(pptFilePath, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
                presentation.SaveAs(outputPdfPath, PowerPointApp.PpSaveAsFileType.ppSaveAsPDF);
            }
            finally
            {
                presentation?.Close();
                pptApp.Quit();
            }

            return outputPdfPath;
        }

        public string ConvertOutlookMsgToPdf(string msgFilePath)
        {
            var outlookApp = new OutlookApp.Application();
            OutlookApp.MailItem mailItem = null;

            try
            {
                if (!Directory.Exists(_pdfSaveDirectory))
                {
                    throw new DirectoryNotFoundException($"The directory {_pdfSaveDirectory} does not exist.");
                }

                var beforePrintFiles = Directory.GetFiles(_pdfSaveDirectory).ToList();

                SetPDFCreatorAsDefault();
                mailItem = (OutlookApp.MailItem)outlookApp.Session.OpenSharedItem(msgFilePath);
                mailItem.PrintOut();

                var printedFilePath = WaitForNewFile(beforePrintFiles, _pdfSaveDirectory);
                if (string.IsNullOrEmpty(printedFilePath))
                {
                    throw new FileNotFoundException("No new PDF file was found after printing.");
                }

                return printedFilePath;
            }
            finally
            {
                if (mailItem != null)
                {
                    Marshal.ReleaseComObject(mailItem);
                }
                Marshal.ReleaseComObject(outlookApp);
            }
        }

        private string WaitForNewFile(List<string> beforePrintFiles, string pdfSaveDirectory)
        {
            for (int attempt = 0; attempt < 10; attempt++)
            {
                var afterPrintFiles = Directory.GetFiles(pdfSaveDirectory).ToList();
                var printedFiles = afterPrintFiles.Except(beforePrintFiles).ToList();

                if (printedFiles.Count == 1)
                {
                    return printedFiles[0];
                }
                else if (printedFiles.Count > 1)
                {
                    return printedFiles.OrderByDescending(f => File.GetLastWriteTime(f)).First();
                }

                System.Threading.Thread.Sleep(500);
            }

            return null;
        }

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string Name);

        public static void SetPDFCreatorAsDefault()
        {
            SetDefaultPrinter("PDFCreator");
        }
    }
}
