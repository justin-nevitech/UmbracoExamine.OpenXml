using DocumentFormat.OpenXml.Packaging;
using System.Text;

namespace UmbracoExamine.OpenXml
{
    public class PresentationDocumentTextExtractor : BaseOpenXmlTextExtractor, IPresentationDocumentTextExtractor
    {
        public string GetText(Stream fileStream)
        {
            StringBuilder builder = new StringBuilder();

            var presentationDocument = PresentationDocument.Open(fileStream, false);

            if (presentationDocument.PresentationPart != null)
            {
                builder.AppendLine(GetTextFromPart(presentationDocument.PresentationPart));

                if (presentationDocument.PresentationPart.SlideParts != null)
                {
                    foreach (var slidePart in presentationDocument.PresentationPart.SlideParts)
                    {
                        builder.AppendLine(GetTextFromPart(slidePart));
                    }
                }
            }

            return builder.ToString();
        }
    }
}