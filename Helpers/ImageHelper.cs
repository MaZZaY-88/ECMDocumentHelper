using Microsoft.Extensions.Logging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.IO;
using System.Threading.Tasks;

namespace ECMDocumentHelper.Helpers
{
    public class ImageHelper
    {
        private readonly string _outputDirectory;
        private readonly ILogger<ImageHelper> _logger;

        // Constructor with outputDirectory and logger
        public ImageHelper(string outputDirectory, ILogger<ImageHelper> logger)
        {
            _outputDirectory = outputDirectory ?? throw new ArgumentNullException(nameof(outputDirectory));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        // ConvertImageToPdfAsync method
        public async Task ConvertImageToPdfAsync(string inputImagePath, string outputFileName)
        {
            try
            {
                if (!File.Exists(inputImagePath))
                {
                    _logger.LogError("Image file not found: {FilePath}", inputImagePath);
                    throw new FileNotFoundException("Image file not found", inputImagePath);
                }

                // Set the output path using the output directory and file name
                string outputPdfPath = Path.Combine(_outputDirectory, outputFileName);

                using (PdfDocument document = new PdfDocument())
                {
                    PdfPage page = document.AddPage();

                    using (XGraphics gfx = XGraphics.FromPdfPage(page))
                    {
                        using (XImage image = XImage.FromFile(inputImagePath))
                        {
                            // Resize the PDF page to fit the image
                            page.Width = image.PointWidth;
                            page.Height = image.PointHeight;

                            // Draw the image on the PDF page
                            gfx.DrawImage(image, 0, 0, page.Width, page.Height);
                        }
                    }

                    // Save the PDF to the output path
                    document.Save(outputPdfPath);
                    _logger.LogInformation("Image converted to PDF successfully: {OutputPath}", outputPdfPath);
                }

                await Task.CompletedTask; // Dummy await for async method signature
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error converting image to PDF: {FilePath}", inputImagePath);
                throw new ApplicationException($"Error converting image to PDF: {inputImagePath}", ex);
            }
        }
    }
}
