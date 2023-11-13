using DocumentFormat.OpenXml.Packaging;

RemovePart(args[0]);

// To remove a document part from a package.
static void RemovePart(string document)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
        MainDocumentPart? mainPart = wordDoc.MainDocumentPart;

        if (mainPart is not null && mainPart.DocumentSettingsPart is not null)
        {
            mainPart.DeletePart(mainPart.DocumentSettingsPart);
        }
    }
}
