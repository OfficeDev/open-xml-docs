
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Xml;


static void AddNewPart(string document)
{
    // Create a new word processing document.
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
    {

        // Add the MainDocumentPart part in the new word processing document.
        var mainDocPart = wordDoc.AddNewPart<MainDocumentPart>
    ("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1");
        mainDocPart.Document = new Document();

        // Add the CustomFilePropertiesPart part in the new word processing document.
        var customFilePropPart = wordDoc.AddCustomFilePropertiesPart();
        customFilePropPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();

        // Add the CoreFilePropertiesPart part in the new word processing document.
        var coreFilePropPart = wordDoc.AddCoreFilePropertiesPart();
        using (XmlTextWriter writer = new
    XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
        {
            writer.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<cp:coreProperties xmlns:cp=\"https://schemas.openxmlformats.org/package/2006/metadata/core-properties\"></cp:coreProperties>");
            writer.Flush();
        }

        // Add the DigitalSignatureOriginPart part in the new word processing document.
        wordDoc.AddNewPart<DigitalSignatureOriginPart>("rId4");

        // Add the ExtendedFilePropertiesPart part in the new word processing document.
        var extendedFilePropPart = wordDoc.AddNewPart<ExtendedFilePropertiesPart>("rId5");
        extendedFilePropPart.Properties =
    new DocumentFormat.OpenXml.ExtendedProperties.Properties();

        // Add the ThumbnailPart part in the new word processing document.
        wordDoc.AddNewPart<ThumbnailPart>("image/jpeg", "rId6");
    }
}
