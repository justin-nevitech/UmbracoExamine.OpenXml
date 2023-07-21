using Microsoft.Extensions.Logging;
using Umbraco.Cms.Core.IO;

namespace UmbracoExamine.OpenXml
{
    /// <summary>
    /// Extracts the openXml from a OpenXml document
    /// </summary>
    public class OpenXmlService
    {
        private readonly IOpenXmlTextExtractorFactory _openXmlTextExtractorFactory;
        private readonly MediaFileManager _mediaFileSystem;
        private readonly ILogger<OpenXmlService> _logger;

        public OpenXmlService(
            IOpenXmlTextExtractorFactory openXmlTextExtractorFactory,
            MediaFileManager mediaFileSystem,
            ILogger<OpenXmlService> logger)
        {
            _openXmlTextExtractorFactory = openXmlTextExtractorFactory;
            _mediaFileSystem = mediaFileSystem;
            _logger = logger;
        }

        /// <summary>
        /// Extract openXml from a OpenXml file at the given path
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public string ExtractOpenXml(string filePath)
        {
            using (var fileStream = _mediaFileSystem.FileSystem.OpenFile(filePath))
            {
                if (fileStream != null)
                {
                    var openXmlTextExtractor = _openXmlTextExtractorFactory.GetOpenXmlTextExtractor(Path.GetExtension(filePath).Substring(1));
                    return openXmlTextExtractor.GetText(fileStream);
                }
                else
                {
                    _logger.LogError(new Exception($"Unable to open file {filePath}"), "Unable to open file");
                    return String.Empty;
                }
            }
        }
    }
}