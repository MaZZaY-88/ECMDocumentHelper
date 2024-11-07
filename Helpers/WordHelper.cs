using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ZXing;
using ZXing.Common;
using System.Drawing;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml;
using PdfSharp.Drawing;
using SkiaSharp;
using ZXing;
using ZXing.SkiaSharp;
using ZXing.SkiaSharp.Rendering;
//using Microsoft.Office.Interop.Word;
using static System.Net.Mime.MediaTypeNames;


namespace ECMDocumentHelper.Helpers
{
    public class WordHelper
    {
        public string ReplaceQrTagInWordDocument(string inputDocPath, string qrText)
        {
            try
            {
                string outputDocPath = Path.Combine(Path.GetDirectoryName(inputDocPath), $"{Guid.NewGuid()}_qr.docx");
                File.Copy(inputDocPath, outputDocPath, true);

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputDocPath, true))
                {
                    MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                    // Сгенерировать QR-код
                    byte[] qrImageBytes = GenerateQrCodeImage(qrText);

                    // Найти все текстовые элементы в документе
                    var paragraphs = mainPart.Document.Descendants<Paragraph>();

                    foreach (var paragraph in paragraphs)
                    {
                        var runs = paragraph.Descendants<Run>().ToList();
                        for (int i = 0; i < runs.Count; i++)
                        {
                            var texts = runs[i].Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
                            for (int j = 0; j < texts.Count; j++)
                            {
                                if (texts[j].Text.Contains("{QR}"))
                                {
                                    // Разбить текст, если в нем есть другие символы
                                    string[] splitTexts = texts[j].Text.Split(new string[] { "{QR}" }, StringSplitOptions.None);

                                    // Создать новый Run для текста до {QR}
                                    if (!string.IsNullOrEmpty(splitTexts[0]))
                                    {
                                        DocumentFormat.OpenXml.Drawing.Text beforeText = new DocumentFormat.OpenXml.Drawing.Text(splitTexts[0]);
                                        Run beforeRun = new Run(beforeText);
                                        paragraph.InsertBefore(beforeRun, runs[i]);
                                    }

                                    // Добавить изображение QR-кода
                                    AddImageToBody(mainPart, paragraph, qrImageBytes, (long)2.5 * 360000, (long)2.5 * 360000);

                                    // Создать новый Run для текста после {QR}
                                    if (splitTexts.Length > 1 && !string.IsNullOrEmpty(splitTexts[1]))
                                    {
                                        DocumentFormat.OpenXml.Drawing.Text afterText = new DocumentFormat.OpenXml.Drawing.Text(splitTexts[1]);
                                        Run afterRun = new Run(afterText);
                                        paragraph.InsertAfter(afterRun, runs[i]);
                                    }

                                    // Удалить оригинальный Run с {QR}
                                    runs[i].Remove();
                                    // Так как мы заменили текущий Run, выходим из внутреннего цикла
                                    break;
                                }
                            }
                        }
                    }

                    mainPart.Document.Save();
                }

                return outputDocPath;
            }
            catch (Exception ex)
            {

                throw new ApplicationException("Ошибка при замене тега QR в Word-файле", ex);
            }
        }

        private byte[] GenerateQrCodeImage(string qrText)
        {
            var writer = new BarcodeWriter<SKBitmap>
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new EncodingOptions
                {
                    Height = 100,
                    Width = 100,
                    Margin = 0
                },
                Renderer = new SKBitmapRenderer()
            };

            using (var bitmap = writer.Write(qrText))
            {
                using (var image = SKImage.FromBitmap(bitmap))
                {
                    using (var data = image.Encode(SKEncodedImageFormat.Png, 100))
                    {
                        return data.ToArray();
                    }
                }
            }
        }


        public string ReplaceTagWithImageInWordDocument(string inputDocPath, string tagToReplace, string imagePath)
        {
            try
            {
                string outputDocPath = Path.Combine(Path.GetDirectoryName(inputDocPath), $"{Guid.NewGuid()}_image.docx");
                File.Copy(inputDocPath, outputDocPath, true);

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputDocPath, true))
                {
                    MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                    // Читаем изображение из файла
                    byte[] imageBytes = File.ReadAllBytes(imagePath);

                    // Получаем размеры изображения
                    long imageWidthEmu;
                    long imageHeightEmu;
                    using (var imageStream = new MemoryStream(imageBytes))
                    {
                        using (var bitmap = SKBitmap.Decode(imageStream))
                        {
                            // Конвертируем пиксели в EMU
                            const int emusPerInch = 914400;
                            const int emusPerCm = 360000;
                            float dpiX = bitmap.Width / (bitmap.Width / 96.0f); // Предполагаем 96 DPI
                            float dpiY = bitmap.Height / (bitmap.Height / 96.0f);

                            // Рассчитываем размеры в EMU
                            imageWidthEmu = (long)(bitmap.Width / dpiX * emusPerInch);
                            imageHeightEmu = (long)(bitmap.Height / dpiY * emusPerInch);
                        }
                    }

                    // Найти все абзацы в документе
                    var paragraphs = mainPart.Document.Descendants<Paragraph>();

                    foreach (var paragraph in paragraphs)
                    {
                        var runs = paragraph.Descendants<Run>().ToList();
                        for (int i = 0; i < runs.Count; i++)
                        {
                            var texts = runs[i].Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
                            for (int j = 0; j < texts.Count; j++)
                            {
                                if (texts[j].Text.Contains(tagToReplace))
                                {
                                    // Разбиваем текст, если есть другие символы
                                    string[] splitTexts = texts[j].Text.Split(new string[] { tagToReplace }, StringSplitOptions.None);

                                    // Создаем новый Run для текста до тега
                                    if (!string.IsNullOrEmpty(splitTexts[0]))
                                    {
                                        DocumentFormat.OpenXml.Drawing.Text beforeText = new DocumentFormat.OpenXml.Drawing.Text(splitTexts[0]);
                                        Run beforeRun = new Run(beforeText);
                                        paragraph.InsertBefore(beforeRun, runs[i]);
                                    }

                                    // Добавляем изображение
                                    AddImageToBody(mainPart, paragraph, imageBytes, imageWidthEmu, imageHeightEmu);

                                    // Создаем новый Run для текста после тега
                                    if (splitTexts.Length > 1 && !string.IsNullOrEmpty(splitTexts[1]))
                                    {
                                        DocumentFormat.OpenXml.Drawing.Text afterText = new DocumentFormat.OpenXml.Drawing.Text(splitTexts[1]);
                                        Run afterRun = new Run(afterText);
                                        paragraph.InsertAfter(afterRun, runs[i]);
                                    }

                                    // Удаляем оригинальный Run с тегом
                                    runs[i].Remove();
                                    // Выходим из внутреннего цикла, так как текущий Run был заменен
                                    break;
                                }
                            }
                        }
                    }

                    mainPart.Document.Save();
                }

                return outputDocPath;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Ошибка при замене тега на изображение в Word-файле", ex);
            }
        }

        private void AddImageToBody(MainDocumentPart mainPart, Paragraph paragraph, byte[] imageBytes, long imageWidthEmu, long imageHeightEmu)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(stream);
            }

            AddImageToParagraph(mainPart.GetIdOfPart(imagePart), paragraph, imageWidthEmu, imageHeightEmu);
        }

        private void AddImageToParagraph(string relationshipId, Paragraph paragraph, long imageWidthEmu, long imageHeightEmu)
        {
            var element =
                 new Drawing(
                     new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = imageWidthEmu, Cy = imageHeightEmu },
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Inserted Image"
                         },
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                             new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                         new DocumentFormat.OpenXml.Drawing.Graphic(
                             new DocumentFormat.OpenXml.Drawing.GraphicData(
                                 new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                     new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                         new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "Inserted Image"
                                         },
                                         new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                     new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                         new DocumentFormat.OpenXml.Drawing.Blip()
                                         {
                                             Embed = relationshipId,
                                             CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                         },
                                         new DocumentFormat.OpenXml.Drawing.Stretch(
                                             new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                     new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                         new DocumentFormat.OpenXml.Drawing.Transform2D(
                                             new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                             new DocumentFormat.OpenXml.Drawing.Extents() { Cx = imageWidthEmu, Cy = imageHeightEmu }),
                                         new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                             new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                         )
                                         { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            Run run = new Run(element);
            paragraph.AppendChild(run);
        }
    }

}
