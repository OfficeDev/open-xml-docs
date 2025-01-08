
using DocumentFormat.OpenXml.Packaging;
using System.IO;

// <Snippet3>
string document = args[0];
string fileName = args[1];

AddNewPart(args[0], args[1]);
// </Snippet3>

// To add a new document part to a package.
// <Snippet0>
static void AddNewPart(string document, string fileName)
{
    // <Snippet1>
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
        // </Snippet1>
    {
        // <Snippet2>
        MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();

        CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

        using (FileStream stream = new FileStream(fileName, FileMode.Open))
        {
            myXmlPart.FeedData(stream);
        }
        // </Snippet2>
    }
}
// </Snippet0>
