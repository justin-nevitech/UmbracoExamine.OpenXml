using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace UmbracoExamine.OpenXml
{
    public abstract class BaseOpenXmlTextExtractor
    {
        protected string GetTextFromPart(OpenXmlPart part)
        {
            StringBuilder builder = new StringBuilder();

            using (var reader = OpenXmlReader.Create(part, false))
            {
                while (reader.Read())
                {
                    var line = reader.GetText().Trim();

                    if (String.IsNullOrEmpty(line) == false)
                    {
                        builder.AppendLine(line).Append(" ");
                    }
                }
            }

            return builder.ToString();
        }

        protected string GetTextFromWorksheetParts(WorkbookPart workbookPart)
        {
            StringBuilder builder = new StringBuilder();

            foreach (var worksheet in workbookPart.WorksheetParts)
            {
                using (var reader = OpenXmlReader.Create(worksheet, false))
                {
                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Cell))
                        {
                            var cell = reader.LoadCurrentElement() as Cell;

                            if (cell != null)
                            {
                                string? value;

                                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                                {
                                    var sharedString = workbookPart?.SharedStringTablePart?.SharedStringTable.Elements<SharedStringItem>().ElementAt(Int32.Parse(cell?.CellValue?.InnerText ?? String.Empty));

                                    value = sharedString?.Text?.Text;
                                }
                                else
                                {
                                    value = cell?.CellValue?.InnerText;
                                }

                                value = value?.Trim();

                                if (String.IsNullOrEmpty(value) == false)
                                {
                                    builder.AppendLine(value).Append(" ");
                                }
                            }
                        }
                    }
                }
            }

            return builder.ToString();
        }
    }
}