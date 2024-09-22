
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using System;
using System.IO;

namespace ECMDocumentHelper.Helpers
{
    public class BarcodeHelper
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<BarcodeHelper> _logger;

        public BarcodeHelper(IConfiguration configuration, ILogger<BarcodeHelper> logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public string ImprintBarcodeOnPdf(string inputPdfPath, string barcodeText, string regNumber)
        {
            try
            {
                using (PdfDocument document = PdfReader.Open(inputPdfPath, PdfDocumentOpenMode.Modify))
                {
                    var barcodeProfile = _configuration.GetSection("BarcodeProfile");
                    double xPosition = barcodeProfile.GetValue<double>("XPosition", 10);
                    double yPosition = barcodeProfile.GetValue<double>("YPosition", 90);
                    int rotation = barcodeProfile.GetValue<int>("Rotation", 90);

                    // Convert positions from millimeters to points
                    double xPositionPt = XUnit.FromMillimeter(xPosition);
                    double yPositionPt = XUnit.FromMillimeter(yPosition);

                    // Define fonts
                    XFont barcodeFont = new XFont("LibreBarcode128Text", 10);
                    XFont textFont = new XFont("LiberationSans", 4);

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
                            gfx.DrawString(regNumber, textFont, XBrushes.Black, new XPoint(0, 0));

                            // Measure the size of the barcode text
                            XSize barcodeSize = gfx.MeasureString(barcodeText, barcodeFont);

                            // Define a gap between the barcode and the human-readable text
                            double gap = XUnit.FromPoint(1); // 2-point gap

                            // Draw the human-readable text below the barcode
                            gfx.DrawString(barcodeText, barcodeFont, XBrushes.Black, new XPoint(0, barcodeSize.Height + gap));

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
                _logger.LogError(ex, "Error during barcode imprinting on PDF {FilePath}", inputPdfPath);
                throw new ApplicationException("Error during barcode imprinting", ex);
            }
        }
    }
}
