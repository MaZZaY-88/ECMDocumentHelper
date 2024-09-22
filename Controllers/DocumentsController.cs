using Microsoft.AspNetCore.Mvc;
using ECMDocumentHelper.Services;
using ECMDocumentHelper.Models;
using System.Threading.Tasks;

namespace ECMDocumentHelper.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentsController : ControllerBase
    {
        private readonly PdfProcessingService _pdfProcessingService;

        public DocumentsController(PdfProcessingService pdfProcessingService)
        {
            _pdfProcessingService = pdfProcessingService;
        }

        // Action for generating PDFs from files
        [HttpPost("generatepdf")]
        public async Task<IActionResult> GeneratePDF([FromBody] FileRequest request)
        {
            if (request == null || request.FilePaths == null || request.FilePaths.Count == 0)
            {
                return BadRequest(new { Message = "filePaths field is required." });
            }

            try
            {
                var result = await _pdfProcessingService.GeneratePDFAsync(request.FilePaths);
                return Ok(new { statusCode = result.StatusCode, message = result.Message, outputPath = result.OutputPath });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { statusCode = 0, message = "An error occurred while processing files.", Error = ex.Message });
            }
        }

        // Action for imprinting a barcode on a PDF
        [HttpPost("imprintbarcode")]
        public async Task<IActionResult> ImprintBarcode([FromBody] BarcodeRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.FilePath) || string.IsNullOrWhiteSpace(request.BarcodeText))
            {
                return BadRequest(new { Message = "FilePath and BarcodeText fields are required." });
            }

            try
            {
                var result = await _pdfProcessingService.ImprintBarcodeOnPdfAsync(request.FilePath, request.BarcodeText);
                return Ok(new { statusCode = result.StatusCode, message = result.Message, outputPath = result.OutputPath });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { statusCode = 0, message = "An error occurred while imprinting the barcode.", Error = ex.Message });
            }
        }
    }
}
