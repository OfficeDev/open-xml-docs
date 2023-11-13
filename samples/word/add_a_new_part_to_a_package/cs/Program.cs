
using DocumentFormat.OpenXml.Packaging;
using System.IO;

AddNewPart(args[0], args[1]);

// To add a new document part to a package.
static void AddNewPart(string document, string fileName)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
        MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();

        CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

        using (FileStream stream = new FileStream(fileName, FileMode.Open))
        {
            myXmlPart.FeedData(stream);
        }
    }
}
