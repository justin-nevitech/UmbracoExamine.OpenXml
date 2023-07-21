namespace UmbracoExamine.OpenXml
{
    public class OpenXmlTextExtractorFactory : IOpenXmlTextExtractorFactory
    {
        private readonly Dictionary<string, IOpenXmlTextExtractor> _openXmlTextExtractors;

        private readonly IWordProcessingDocumentTextExtractor _wordProcessingDocumentTextExtractor;
        private readonly IPresentationDocumentTextExtractor _presentationDocumentTextExtractor;
        private readonly ISpreadsheetDocumentTextExtractor _spreadsheetDocumentTextExtractor;

        public OpenXmlTextExtractorFactory(IWordProcessingDocumentTextExtractor wordProcessingDocumentTextExtractor,
            IPresentationDocumentTextExtractor presentationDocumentTextExtractor,
            ISpreadsheetDocumentTextExtractor spreadsheetDocumentTextExtractor)
        {
            _wordProcessingDocumentTextExtractor = wordProcessingDocumentTextExtractor;
            _presentationDocumentTextExtractor = presentationDocumentTextExtractor;
            _spreadsheetDocumentTextExtractor = spreadsheetDocumentTextExtractor;

            _openXmlTextExtractors = new Dictionary<string, IOpenXmlTextExtractor>()
            {
                { OpenXmlIndexConstants.WordProcessingDocumentFileExtension, _wordProcessingDocumentTextExtractor },
                { OpenXmlIndexConstants.PresentationDocumentFileExtension, _presentationDocumentTextExtractor },
                { OpenXmlIndexConstants.SpreadsheetDocumentFileExtension, _spreadsheetDocumentTextExtractor }
            };
        }

        public IOpenXmlTextExtractor GetOpenXmlTextExtractor(string extension)
        {
            var openXmlTextExtractor = _openXmlTextExtractors.ContainsKey(extension) ? _openXmlTextExtractors[extension] : null;

            if (openXmlTextExtractor != null)
            {
                return openXmlTextExtractor;
            }
            else
            {
                throw new Exception($"No text extractor defined for extension '{extension}'");
            }
        }
    }
}