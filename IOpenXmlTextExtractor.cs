namespace UmbracoExamine.OpenXml
{
    public interface IOpenXmlTextExtractor
    {
        string GetText(Stream fileStream);
    }
}