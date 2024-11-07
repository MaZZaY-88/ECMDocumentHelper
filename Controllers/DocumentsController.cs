using ECMDocumentHelper.Services;
using ECMDocumentHelper.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ECMDocumentHelper.Helpers;
using DocumentFormat.OpenXml;

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

        // Existing GeneratePDF method (as requested)
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
            int statusCode = 0; // Код статуса выполнения (1 - успех, 0 - ошибка)
            string message = string.Empty; // Сообщение о результате работы
            string outputPath = string.Empty; // Путь до результирующего файла

            try
            {
                // Проверка существования файла шаблона
                if (!System.IO.File.Exists(request.Template))
                {
                    statusCode = 0;
                    message = "Шаблон не найден.";
                    return NotFound(new { statusCode, message });
                }

                // Генерация пути для сохранения результирующего документа
                string templateDirectory = Path.GetDirectoryName(request.Template);
                string templateFileName = Path.GetFileNameWithoutExtension(request.Template);
                string templateExtension = Path.GetExtension(request.Template);
                outputPath = Path.Combine(templateDirectory, $"{Guid.NewGuid()}_Result{templateExtension}");

                // Копирование шаблона в новый файл для сохранения исходного шаблона неизменным
                System.IO.File.Copy(request.Template, outputPath, true);

                // Открытие Word-документа для редактирования
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
                {
                    // Получение тела основного документа
                    var body = wordDoc.MainDocumentPart.Document.Body;

                    // Проход по всем текстовым элементам в документе
                    foreach (var text in body.Descendants<Text>().ToList())
                    {
                        // Проход по всем парам ключ-значение для замены
                        foreach (var item in request.Data)
                        {
                            // Формирование тега в фигурных скобках
                            string tag = $"{{{item.Key}}}";

                            // Проверка, содержит ли текстовый элемент текущий тег
                            if (text.Text.Contains(tag))
                            {
                                string tagValue = item.Value;
                                if (item.Value == null)
                                    tagValue = "";
                                // Разбиваем значение по символу \n
                                string[] lines = tagValue.Split(new[] { "\\n" }, StringSplitOptions.None);

                                // Получаем родительский Run для текущего Text
                                Run parentRun = text.Parent as Run;

                                if (parentRun != null)
                                {
                                    // Копируем RunProperties из оригинального Run
                                    RunProperties runProperties = parentRun.RunProperties != null
                                        ? (RunProperties)parentRun.RunProperties.CloneNode(true)
                                        : new RunProperties();

                                    // Проверяем и устанавливаем цвет шрифта
                                    Color color = runProperties.Elements<Color>().FirstOrDefault();

                                    if (color != null)
                                    {
                                        // Удаляем существующий элемент Color
                                        color.Remove();
                                    }

                                    // Добавляем новый элемент Color с чёрным цветом
                                    runProperties.Append(new Color() { Val = "000000" });

                                    // Список новых элементов для вставки
                                    List<OpenXmlElement> newElements = new List<OpenXmlElement>();

                                    for (int i = 0; i < lines.Length; i++)
                                    {
                                        // Создаем новый Run с текстом и скопированными RunProperties
                                        Run newRun = new Run();
                                        newRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
                                        newRun.AppendChild(new Text(lines[i]));

                                        newElements.Add(newRun);

                                        // Добавляем разрыв строки, если это не последний элемент
                                        if (i < lines.Length - 1)
                                        {
                                            Run breakRun = new Run();
                                            breakRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
                                            breakRun.AppendChild(new Break());
                                            newElements.Add(breakRun);
                                        }
                                    }

                                    // Вставляем новые элементы перед текущим Run
                                    foreach (var newElement in newElements)
                                    {
                                        parentRun.Parent.InsertBefore(newElement, parentRun);
                                    }

                                    // Удаляем оригинальный Run
                                    parentRun.Remove();
                                }

                                // Переходим к следующему Text элементу, так как текущий был удален
                                break;
                            }
                        }
                    }

                    // Сохранение изменений в документе
                    wordDoc.MainDocumentPart.Document.Save();
                }

                statusCode = 1; // Успешное выполнение
                message = "Документ успешно сгенерирован.";
            }
            catch (Exception ex)
            {
                // Обработка исключения и установка сообщения об ошибке
                statusCode = 0;
                message = $"Ошибка при генерации документа: {ex.Message}";
            }

            // Возврат результата выполнения
            return Ok(new { statusCode, message, outputPath });
        }





    }

}
