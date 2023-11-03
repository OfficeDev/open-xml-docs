
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;


    // To add a new document part to a package.
    public static void AddNewPart(string document, string fileName)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;

            CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                myXmlPart.FeedData(stream);
            }
        }
    }
