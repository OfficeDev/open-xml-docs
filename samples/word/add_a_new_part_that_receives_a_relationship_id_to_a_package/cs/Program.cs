
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Xml;

if (File.Exists(args[0]))
{
    File.Delete(args[0]);
}

AddNewPart(args[0]);
// <Snippet0>
// <Snippet1>
static void AddNewPart(string document)
{
    // Create a new word processing document.
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
// </Snippet1>
    {
        // <Snippet2>
        // Add the MainDocumentPart part in the new word processing document.
        MainDocumentPart mainDocPart = wordDoc.AddMainDocumentPart();
        mainDocPart.Document = new Document();

        // Add the CustomFilePropertiesPart part in the new word processing document.
        var customFilePropPart = wordDoc.AddCustomFilePropertiesPart();
        customFilePropPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();

        // Add the CoreFilePropertiesPart part in the new word processing document.
        var coreFilePropPart = wordDoc.AddCoreFilePropertiesPart();
        using (XmlTextWriter writer = new XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
        {
            writer.WriteRaw("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" />
                """);
            writer.Flush();
        }
        // </Snippet2>
        // <Snippet3>
        // Add the DigitalSignatureOriginPart part in the new word processing document.
        wordDoc.AddNewPart<DigitalSignatureOriginPart>("rId4");

        // Add the ExtendedFilePropertiesPart part in the new word processing document.
        var extendedFilePropPart = wordDoc.AddNewPart<ExtendedFilePropertiesPart>("rId5");
        extendedFilePropPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();

        // Add the ThumbnailPart part in the new word processing document.
        wordDoc.AddNewPart<ThumbnailPart>("image/jpeg", "rId6");
        // </Snippet3>
    }
}
// </Snippet0>
