using System;
using System.IO;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Outlook= Microsoft.Office.Interop.Outlook;

namespace ECMDocumentHelper.Helpers
{
    public class OfficeInteropHelper
    {
        private Word.Application _wordApp;
        private Excel.Application _excelApp;
        private PowerPoint.Application _powerPointApp;
        private Outlook.Application _outlookApp;

        public OfficeInteropHelper()
        {
            _wordApp = new Word.Application();
            _excelApp = new Excel.Application();
            _powerPointApp = new PowerPoint.Application();
            _outlookApp = new Outlook.Application();
        }

        public async Task ConvertWordToPdfAsync(string inputFilePath, string outputFilePath)
        {
            try
            {
                Word.Document document = _wordApp.Documents.Open(inputFilePath);
                document.ExportAsFixedFormat(outputFilePath, Word.WdExportFormat.wdExportFormatPDF);
                document.Close();
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error converting Word to PDF: {ex.Message}", ex);
            }
            finally
            {
                _wordApp.Quit();
            }
        }

        public async Task ConvertExcelToPdfAsync(string inputFilePath, string outputFilePath)
        {
            try
            {
                Excel.Workbook workbook = _excelApp.Workbooks.Open(inputFilePath);
                workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFilePath);
                workbook.Close();
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error converting Excel to PDF: {ex.Message}", ex);
            }
            finally
            {
                _excelApp.Quit();
            }
        }

        public async Task ConvertPowerPointToPdfAsync(string inputFilePath, string outputFilePath)
        {
            try
            {
                PowerPoint.Presentation presentation = _powerPointApp.Presentations.Open(inputFilePath);
                presentation.SaveAs(outputFilePath, PowerPoint.PpSaveAsFileType.ppSaveAsPDF);
                presentation.Close();
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error converting PowerPoint to PDF: {ex.Message}", ex);
            }
            finally
            {
                _powerPointApp.Quit();
            }
        }

        public async Task ConvertOutlookMsgToPdfAsync(string inputFilePath, string outputFilePath)
        {
            try
            {
                Outlook.MailItem mailItem = (Outlook.MailItem)_outlookApp.Session.OpenSharedItem(inputFilePath);
                mailItem.SaveAs(outputFilePath, Outlook.OlSaveAsType.olDoc);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error converting Outlook MSG to PDF: {ex.Message}", ex);
            }
        }

       
    }
}
