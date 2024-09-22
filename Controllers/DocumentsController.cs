
using Microsoft.AspNetCore.Mvc;
using ECMDocumentHelper.Services;
using ECMDocumentHelper.Models;

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

        [HttpPost("generatepdf")]
        public IActionResult GeneratePDF([FromBody] FileRequest request)
        {
            if (request == null || request.FilePaths == null || request.FilePaths.Count == 0)
            {
                return BadRequest(new { Message = "filePaths field is required." });
            }

            try
            {
                // Call service to generate PDF
                var result = _pdfProcessingService.GeneratePDF(request.FilePaths);
                return Ok(new { StatusCode = result.StatusCode, Message = result.Message, OutputPath = result.OutputPath });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { StatusCode = 0, Message = "An error occurred while generating PDF.", Error = ex.Message });
            }
        }

        [HttpPost("imprintbarcode")]
        public IActionResult ImprintBarcode([FromBody] BarcodeRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.FilePath) || string.IsNullOrWhiteSpace(request.BarcodeText))
            {
                return BadRequest(new { Message = "FilePath and BarcodeText fields are required." });
            }

            try
            {
                var result = _pdfProcessingService.ImprintBarcodeOnPdf(request.FilePath, request.BarcodeText, request.RegNumber);
                return Ok(new { StatusCode = result.StatusCode, Message = result.Message, OutputPath = result.OutputPath });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { StatusCode = 0, Message = "An error occurred while imprinting the barcode.", Error = ex.Message });
            }
        }
    }
}
