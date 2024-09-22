using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using System;
using System.IO;
using System.Threading.Tasks;

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

        public async Task<(int statusCode, string message, string outputPath)> ImprintBarcodeOnPdf(string inputPdfPath, string barcodeText)
        {
            try
            {
                using (PdfDocument document = PdfReader.Open(inputPdfPath, PdfDocumentOpenMode.Modify))
                {
                    var barcodeProfile = _configuration.GetSection("BarcodeProfile");
                    double xPosition = barcodeProfile.GetValue<double>("XPosition", 10);
                    double yPosition = barcodeProfile.GetValue<double>("YPosition", 0.5);
                    int rotation = barcodeProfile.GetValue<int>("Rotation", 90);

                    XFont barcodeFont = new XFont("LibreBarcode128Text", 10);

                    foreach (PdfPage page in document.Pages)
                    {
                        XGraphics gfx = XGraphics.FromPdfPage(page);
                        gfx.Save();
                        gfx.RotateTransform(rotation);
                        gfx.DrawString(barcodeText, barcodeFont, XBrushes.Black, new XPoint(xPosition, yPosition));
                        gfx.Restore();
                    }

                    string guid = Guid.NewGuid().ToString();
                    string outputPdfPath = Path.Combine(Path.GetDirectoryName(inputPdfPath), guid + "_barcode.pdf");

                    document.Save(outputPdfPath);
                    return (1, "Barcode imprinted successfully.", outputPdfPath);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during barcode imprinting on PDF {FilePath}", inputPdfPath);
                return (0, $"Error during barcode imprinting: {ex.Message}", null);
            }
        }
    }
}
