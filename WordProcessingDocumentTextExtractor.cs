using DocumentFormat.OpenXml.Packaging;
using System.Text;

namespace UmbracoExamine.OpenXml
{
    public class WordProcessingDocumentTextExtractor : BaseOpenXmlTextExtractor, IWordProcessingDocumentTextExtractor
    {
        public string GetText(Stream fileStream)
        {
            StringBuilder builder = new StringBuilder();

            var wordprocessingDocument = WordprocessingDocument.Open(fileStream, false);

            if (wordprocessingDocument.MainDocumentPart != null)
            {
                builder.AppendLine(GetTextFromPart(wordprocessingDocument.MainDocumentPart));
            }

            return builder.ToString();
        }
    }
}