﻿using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ZXing;
using ZXing.Common;
using System.Drawing;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml;
using SkiaSharp;
using ZXing.SkiaSharp;
using ZXing.SkiaSharp.Rendering;

namespace ECMDocumentHelper.Helpers
{
    public class WordHelper
    {
        public string ReplaceQrTagInWordDocument(string inputDocPath, string qrText, int qrCodeSize = 100)
        {
            try
            {
                string outputDocPath = Path.Combine(Path.GetDirectoryName(inputDocPath), $"{Guid.NewGuid()}_qr.docx");
                File.Copy(inputDocPath, outputDocPath, true);

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputDocPath, true))
                {
                    MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                    // Сгенерировать QR-код с указанным размером
                    byte[] qrImageBytes = GenerateQrCodeImage(qrText, qrCodeSize);

                    // Найти все текстовые элементы в документе
                    var paragraphs = mainPart.Document.Body.Descendants<Paragraph>();

                    foreach (var paragraph in paragraphs)
                    {
                        var runs = paragraph.Descendants<Run>().ToList();
                        for (int i = 0; i < runs.Count; i++)
                        {
                            var texts = runs[i].Descendants<Text>().ToList();
                            for (int j = 0; j < texts.Count; j++)
                            {
                                if (texts[j].Text.Contains("{QR}"))
                                {
                                    // Разбить текст, если в нем есть другие символы
                                    string[] splitTexts = texts[j].Text.Split(new string[] { "{QR}" }, StringSplitOptions.None);

                                    // Создать новый Run для текста до {QR}
                                    if (!string.IsNullOrEmpty(splitTexts[0]))
                                    {
                                        Text beforeText = new Text(splitTexts[0]);
                                        Run beforeRun = new Run(beforeText);
                                        paragraph.InsertBefore(beforeRun, runs[i]);
                                    }

                                    // Добавить изображение QR-кода
                                    double qrCodeSizeEmu = qrCodeSize * 914400 / 96; // Convert pixels to EMU
                                    AddImageToBody(mainPart, paragraph, qrImageBytes, qrCodeSizeEmu, qrCodeSizeEmu);

                                    // Создать новый Run для текста после {QR}
                                    if (splitTexts.Length > 1 && !string.IsNullOrEmpty(splitTexts[1]))
                                    {
                                        Text afterText = new Text(splitTexts[1]);
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

        private byte[] GenerateQrCodeImage(string qrText, int size)
        {
            var writer = new BarcodeWriter<SKBitmap>
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new EncodingOptions
                {
                    Height = size,
                    Width = size,
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

                    // Найти все абзацы в документе
                    var paragraphs = mainPart.Document.Body.Descendants<Paragraph>();

                    foreach (var paragraph in paragraphs)
                    {
                        var runs = paragraph.Descendants<Run>().ToList();
                        for (int i = 0; i < runs.Count; i++)
                        {
                            var texts = runs[i].Descendants<Text>().ToList();
                            for (int j = 0; j < texts.Count; j++)
                            {
                                if (texts[j].Text.Contains(tagToReplace))
                                {
                                    // Разбиваем текст, если есть другие символы
                                    string[] splitTexts = texts[j].Text.Split(new string[] { tagToReplace }, StringSplitOptions.None);

                                    // Создаем новый Run для текста до тега
                                    if (!string.IsNullOrEmpty(splitTexts[0]))
                                    {
                                        Text beforeText = new Text(splitTexts[0]);
                                        Run beforeRun = new Run(beforeText);
                                        paragraph.InsertBefore(beforeRun, runs[i]);
                                    }

                                    // Добавляем изображение
                                    AddImageToBody(mainPart, paragraph, imageBytes, 2.5 * 914400, 2.5 * 914400);

                                    // Создаем новый Run для текста после тега
                                    if (splitTexts.Length > 1 && !string.IsNullOrEmpty(splitTexts[1]))
                                    {
                                        Text afterText = new Text(splitTexts[1]);
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

        private void AddImageToBody(MainDocumentPart mainPart, Paragraph paragraph, byte[] imageBytes, double imageWidthEmu, double imageHeightEmu)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(stream);
            }

            AddImageToParagraph(mainPart.GetIdOfPart(imagePart), paragraph, (long)imageWidthEmu, (long)imageHeightEmu);
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
                                             new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                                         { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                         )
                     )
                 );

            paragraph.AppendChild(new Run(element));
        }
    }
}
