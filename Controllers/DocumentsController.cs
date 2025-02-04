using ECMDocumentHelper.Services;
using ECMDocumentHelper.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ECMDocumentHelper.Helpers;
using DocumentFormat.OpenXml;
using System.Text;

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

        // Existing GeneratePDF method (as requested)
        [HttpPost("generatepdf")]
        public IActionResult GeneratePdf([FromBody] FilePathsRequest request)
        {
            if (request.FilePaths == null || request.FilePaths.Count == 0)
            {
                return BadRequest(new { StatusCode = 0, Message = "No file paths provided." });
            }

            try
            {
                // The returned value is now the full path
                var fullOutputPath = _pdfProcessingService.ConvertFilesToMergedPdf(request.FilePaths);
                return Ok(new { StatusCode = 1, Message = "PDFs merged successfully.", OutputPath = fullOutputPath });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { StatusCode = 0, Message = $"Internal server error: {ex.Message}", Details = ex.InnerException?.Message });
            }
        }

        // Existing GeneratePDF method (as requested)
        [HttpPost("generateQR")]
        public IActionResult GenerateQR([FromBody] QRRequest request)
        {
            if (request.FilePath == null)
            {
                return BadRequest(new { StatusCode = 0, Message = "No file path provided." });
            }

            try
            {
                WordHelper wordHelper = new WordHelper();
                // The returned value is now the full path
                var fullOutputPath = wordHelper.ReplaceQrTagInWordDocument(request.FilePath, request.QRText);
                return Ok(new { StatusCode = 1, Message = "QR imprinted successfully", OutputPath = fullOutputPath });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { StatusCode = 0, Message = $"Internal server error: {ex.Message}", Details = ex.InnerException?.Message });
            }
        }

        [HttpPost("imprintImage")]
        public IActionResult imprintImage([FromBody] ImageRequest request)
        {
            if (request.FilePath == null)
            {
                return BadRequest(new { StatusCode = 0, Message = "No file path provided." });
            }
            if (request.ImagePath == null)
            {
                return BadRequest(new { StatusCode = 0, Message = "No image path provided." });
            }

            try
            {
                WordHelper wordHelper = new WordHelper();
                // The returned value is now the full path
                var fullOutputPath = wordHelper.ReplaceTagWithImageInWordDocument(request.FilePath, request.Tag, request.ImagePath);
                return Ok(new { StatusCode = 1, Message = "Image imprinted successfully", OutputPath = fullOutputPath });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { StatusCode = 0, Message = $"Internal server error: {ex.Message}", Details = ex.InnerException?.Message });
            }
        }

        // Endpoint to imprint a barcode on the PDF
        [HttpPost("imprintbarcode")]
        public IActionResult ImprintBarcodeOnPdf([FromBody] BarcodeRequest request)
        {
            if (string.IsNullOrEmpty(request.FilePath) || string.IsNullOrEmpty(request.BarcodeText))
            {
                return BadRequest(new { StatusCode = 0, Message = "File path or barcode text is missing." });
            }

            try
            {
                var result = _pdfProcessingService.ImprintBarcodeOnPdf(request.FilePath, request.BarcodeText, request.RegNumber);
                return Ok(new { StatusCode = result.StatusCode, Message = result.Message, OutputPath = result.OutputPath });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { StatusCode = 0, Message = $"Internal server error: {ex.Message}", Details = ex.InnerException?.Message });
            }
        }

        [HttpPost("generateword")]
        public IActionResult GenerateWord([FromBody] WordTemplateRequest request)
        {
            int statusCode = 0; // Execution status code (1 - success, 0 - error)
            string message = string.Empty; // Result message
            string outputPath = string.Empty; // Path to the resulting file

            try
            {
                Logger.LogInformation("Starting GenerateWord method.");

                // Check if the template file exists
                if (!System.IO.File.Exists(request.Template))
                {
                    statusCode = 0;
                    message = "Template not found.";
                    Logger.LogWarning("Template not found: " + request.Template);
                    return NotFound(new { statusCode, message });
                }

                // Generate the path for saving the resulting document
                string templateDirectory = Path.GetDirectoryName(request.Template);
                string templateFileName = Path.GetFileNameWithoutExtension(request.Template);
                string templateExtension = Path.GetExtension(request.Template);
                outputPath = Path.Combine(templateDirectory, $"{Guid.NewGuid()}_Result{templateExtension}");

                Logger.LogInformation("Generating output path: " + outputPath);

                // Copy the template to a new file to keep the original template unchanged
                System.IO.File.Copy(request.Template, outputPath, true);

                // Open the Word document for editing
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
                {
                    Logger.LogInformation("Opened Word document for editing: " + outputPath);

                    // Get the main document body
                    var body = wordDoc.MainDocumentPart.Document.Body;

                    // Process the list of text elements to combine split tags
                    var textElements = body.Descendants<Text>().ToList();
                    StringBuilder combinedText = new StringBuilder();
                    bool insideTag = false;
                    List<Text> elementsToRemove = new List<Text>();

                    for (int i = 0; i < textElements.Count; i++)
                    {
                        if (textElements[i].Text.StartsWith("{") && !textElements[i].Text.EndsWith("}"))
                        {
                            insideTag = true;
                            combinedText.Append(textElements[i].Text);
                            elementsToRemove.Add(textElements[i]);
                        }
                        else if (insideTag)
                        {
                            combinedText.Append(textElements[i].Text);
                            elementsToRemove.Add(textElements[i]);
                            if (textElements[i].Text.EndsWith("}"))
                            {
                                insideTag = false;
                                textElements[i].Text = combinedText.ToString();
                                combinedText.Clear();
                                elementsToRemove.Remove(textElements[i]);
                            }
                        }
                    }

                    foreach (var element in elementsToRemove)
                    {
                        element.Remove();
                    }

                    // Iterate over all text elements in the document

                    foreach (var item in request.Data)
                    {
                        Logger.LogInformation($"Processing key: {item.Key}, value: {item.Value}");
                    }

                    foreach (var text in textElements)
                    {
                        // Iterate over all key-value pairs for replacement
                        foreach (var item in request.Data)
                        {

                            // Create the tag in curly braces
                            string tag = $"{{{item.Key}}}";

                            // Check if the text element contains the current tag
                            if (text.Text.Contains(tag))
                            {
                                string tagValue = item.Value;
                                if (item.Value == null)
                                    tagValue = "";
                                // Split the value by the \n character
                                string[] lines = tagValue.Split(new[] { "\\n" }, StringSplitOptions.None);

                                // Get the parent Run for the current Text
                                Run parentRun = text.Parent as Run;

                                if (parentRun != null)
                                {
                                    // Copy RunProperties from the original Run
                                    RunProperties runProperties = parentRun.RunProperties != null
                                        ? (RunProperties)parentRun.RunProperties.CloneNode(true)
                                        : new RunProperties();

                                    // Check and set the font color
                                    Color color = runProperties.Elements<Color>().FirstOrDefault();

                                    if (color != null)
                                    {
                                        // Remove the existing Color element
                                        color.Remove();
                                    }

                                    // Add a new Color element with black color
                                    runProperties.Append(new Color() { Val = "000000" });

                                    // List of new elements to insert
                                    List<OpenXmlElement> newElements = new List<OpenXmlElement>();

                                    for (int i = 0; i < lines.Length; i++)
                                    {
                                        // Create a new Run with the text and copied RunProperties
                                        Run newRun = new Run();
                                        newRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
                                        newRun.AppendChild(new Text(lines[i]));

                                        newElements.Add(newRun);

                                        // Add a line break if this is not the last element
                                        if (i < lines.Length - 1)
                                        {
                                            Run breakRun = new Run();
                                            breakRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
                                            breakRun.AppendChild(new Break());
                                            newElements.Add(breakRun);
                                        }
                                    }

                                    // Insert new elements before the current Run
                                    foreach (var newElement in newElements)
                                    {
                                        parentRun.Parent.InsertBefore(newElement, parentRun);
                                    }

                                    // Remove the original Run
                                    parentRun.Remove();
                                }

                                // Move to the next Text element since the current one was removed
                                break;
                            }
                        }
                    }

                    // Save changes to the document
                    wordDoc.MainDocumentPart.Document.Save();
                    Logger.LogInformation("Document changes saved successfully.");
                }

                statusCode = 1; // Successful execution
                message = "Document generated successfully.";
                Logger.LogInformation(message);
            }
            catch (Exception ex)
            {
                // Handle the exception and set the error message
                statusCode = 0;
                message = $"Error generating document: {ex.Message}";
                Logger.LogError(message, ex);
            }

            // Return the result
            return Ok(new { statusCode, message, outputPath });
        }


    }
}
