namespace UmbracoExamine.OpenXml
{
    public interface IOpenXmlTextExtractorFactory
    {
        IOpenXmlTextExtractor GetOpenXmlTextExtractor(string extension);
    }
}