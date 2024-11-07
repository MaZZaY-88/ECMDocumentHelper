namespace ECMDocumentHelper.Models
{

    public class WordTemplateRequest
    {
        public string Template { get; set; } // Путь до шаблона Word-документа
        public Dictionary<string, string> Data { get; set; } // Словарь тегов и их значений
    }


}
