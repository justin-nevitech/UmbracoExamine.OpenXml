using DocumentFormat.OpenXml.Packaging;
using System.Text;

namespace UmbracoExamine.OpenXml
{
    public class SpreadsheetDocumentTextExtractor : BaseOpenXmlTextExtractor, ISpreadsheetDocumentTextExtractor
    {
        public string GetText(Stream fileStream)
        {
            StringBuilder builder = new StringBuilder();

            var spreadsheetDocument = SpreadsheetDocument.Open(fileStream, false);

            if (spreadsheetDocument.WorkbookPart != null)
            {
                builder.AppendLine(GetTextFromPart(spreadsheetDocument.WorkbookPart));
                builder.AppendLine(GetTextFromWorksheetParts(spreadsheetDocument.WorkbookPart));                
            }

            return builder.ToString();
        }
    }
}