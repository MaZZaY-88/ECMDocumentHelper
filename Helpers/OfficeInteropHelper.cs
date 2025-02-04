using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Configuration;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;

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

            using (var outputDocument = new PdfDocument())
            {
                foreach (var pdfFile in pdfFilePaths)
                {
                    using (var inputDocument = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import))
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
                throw new ApplicationException("Error during barcode imprinting:" + ex.Message, ex);
            }
        }

        public string ConvertWordToPdf(string wordFilePath)
        {
            string outputPdfPath = Path.Combine(_outputDirectory, $"{Path.GetFileNameWithoutExtension(wordFilePath)}_{Guid.NewGuid()}.pdf");

            try
            {
                // Use PDFCreator printer to print the document as a PDF
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = wordFilePath;
                startInfo.Verb = "Print";
                startInfo.Arguments = $"/D:PDFCreator";
                startInfo.UseShellExecute = true;
                startInfo.CreateNoWindow = true;

                using (Process process = Process.Start(startInfo))
                {
                    process.WaitForExit();
                    if (process.ExitCode != 0)
                    {
                        throw new ApplicationException("PDFCreator failed to convert Word to PDF.");
                    }
                }

                // Wait for the PDF to be generated
                string printedPdfPath = Path.Combine(_pdfSaveDirectory, $"{Path.GetFileNameWithoutExtension(wordFilePath)}.pdf");
                if (File.Exists(printedPdfPath))
                {
                    File.Move(printedPdfPath, outputPdfPath);
                }
                else
                {
                    throw new FileNotFoundException("PDF file was not created by PDFCreator.");
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Error during Word to PDF conversion", ex);
            }

            return outputPdfPath;
        }

        public string ConvertExcelToPdf(string excelFilePath)
        {
            string outputPdfPath = Path.Combine(_outputDirectory, $"{Path.GetFileNameWithoutExtension(excelFilePath)}_{Guid.NewGuid()}.pdf");

            try
            {
                // Use PDFCreator printer to print the document as a PDF
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = excelFilePath;
                startInfo.Verb = "Print";
                startInfo.Arguments = $"/D:PDFCreator";
                startInfo.UseShellExecute = true;
                startInfo.CreateNoWindow = true;

                using (Process process = Process.Start(startInfo))
                {
                    process.WaitForExit();
                    if (process.ExitCode != 0)
                    {
                        throw new ApplicationException("PDFCreator failed to convert Excel to PDF.");
                    }
                }

                // Wait for the PDF to be generated
                string printedPdfPath = Path.Combine(_pdfSaveDirectory, $"{Path.GetFileNameWithoutExtension(excelFilePath)}.pdf");
                if (File.Exists(printedPdfPath))
                {
                    File.Move(printedPdfPath, outputPdfPath);
                }
                else
                {
                    throw new FileNotFoundException("PDF file was not created by PDFCreator.");
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Error during Excel to PDF conversion", ex);
            }

            return outputPdfPath;
        }

        public string ConvertPowerPointToPdf(string pptFilePath)
        {
            string outputPdfPath = Path.Combine(_outputDirectory, $"{Path.GetFileNameWithoutExtension(pptFilePath)}_{Guid.NewGuid()}.pdf");

            try
            {
                // Use PDFCreator printer to print the document as a PDF
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = pptFilePath;
                startInfo.Verb = "Print";
                startInfo.Arguments = $"/D:PDFCreator";
                startInfo.UseShellExecute = true;
                startInfo.CreateNoWindow = true;

                using (Process process = Process.Start(startInfo))
                {
                    process.WaitForExit();
                    if (process.ExitCode != 0)
                    {
                        throw new ApplicationException("PDFCreator failed to convert PowerPoint to PDF.");
                    }
                }

                // Wait for the PDF to be generated
                string printedPdfPath = Path.Combine(_pdfSaveDirectory, $"{Path.GetFileNameWithoutExtension(pptFilePath)}.pdf");
                if (File.Exists(printedPdfPath))
                {
                    File.Move(printedPdfPath, outputPdfPath);
                }
                else
                {
                    throw new FileNotFoundException("PDF file was not created by PDFCreator.");
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Error during PowerPoint to PDF conversion", ex);
            }

            return outputPdfPath;
        }
        public string ConvertToPdfUsingLibreOffice(string inputFilePath)
        {
            string outputPdfPath = Path.Combine(_outputDirectory, Path.GetFileNameWithoutExtension(inputFilePath) + ".pdf");

            try
            {
                // Use LibreOffice to convert the document to PDF
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "C:\\Program Files\\LibreOffice\\program\\soffice", // Path to LibreOffice executable (usually soffice)
                    Arguments = "--headless --convert-to pdf \"" + inputFilePath + "\" --outdir \"" + _outputDirectory + "\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(startInfo))
                {
                    process.WaitForExit();
                    if (process.ExitCode != 0)
                    {
                        throw new ApplicationException("LibreOffice failed to convert document to PDF.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Error during document to PDF conversion using LibreOffice: " + ex.Message, ex);
            }

            return outputPdfPath;
        }
    }
}
