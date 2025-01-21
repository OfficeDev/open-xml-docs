using DocumentFormat.OpenXml.Packaging;

RemovePart(args[0]);

// <Snippet0>
// To remove a document part from a package.
static void RemovePart(string document)
{
    // <Snippet1>
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
        // </Snippet1>
    {
        // <Snippet2>
        MainDocumentPart? mainPart = wordDoc.MainDocumentPart;

        if (mainPart is not null && mainPart.DocumentSettingsPart is not null)
        {
            mainPart.DeletePart(mainPart.DocumentSettingsPart);
        }
        // </Snippet2>
    }
}
// </Snippet0>
