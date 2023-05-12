using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main()
    {
        string filePath = "templateletter.docx";

        using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
        {
            string[] templateFields = GetTemplateFields(document);

            Dictionary<string, string> fieldValues = new Dictionary<string, string>
            {
                { "Name", "Julie" },
                { "count", "20" },
                { "sender", "Miguel" }
            };

            PopulateDocumentFields(document, fieldValues);
        }
    }

    static string[] GetTemplateFields(WordprocessingDocument document)
    {
        List<string> fieldNames = new List<string>();

        {
            MainDocumentPart mainPart = document.MainDocumentPart;

            IEnumerable<BookmarkStart> fields = mainPart.RootElement.Descendants<BookmarkStart>();

            foreach (BookmarkStart field in fields)
            {
                string fieldName = field.Name;
                fieldNames.Add(fieldName);
            }
        }


        return fieldNames.ToArray();
    }


    static void PopulateDocumentFields(WordprocessingDocument document, Dictionary<string, string> fieldValues)
    {
        MainDocumentPart mainPart = document.MainDocumentPart;
        var textElements = mainPart.RootElement.Descendants<Text>();

        foreach (var fieldValue in fieldValues)
        {
            string fieldName = fieldValue.Key;
            string value = fieldValue.Value;
            var matchedField = mainPart.RootElement.Descendants<BookmarkStart>()
                .FirstOrDefault(b => { return b?.Name == fieldName; });
            var textField = matchedField.NextSibling().NextSibling().NextSibling().Descendants<Text>().FirstOrDefault();
            if (textField != null)
            {
                textField.Text = value;
            }
        }

        document.SaveAs("editedDoc.docx");
    }
}