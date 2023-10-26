#nullable disable

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

static void AddHeaderFromTo(string filepathFrom, string filepathTo)
{
    // Replace header in target document with header of source document.
    using (WordprocessingDocument
        wdDoc = WordprocessingDocument.Open(filepathTo, true))
    {
        MainDocumentPart mainPart = wdDoc.MainDocumentPart;

        // Delete the existing header part.
        mainPart.DeleteParts(mainPart.HeaderParts);

        // Create a new header part.
        DocumentFormat.OpenXml.Packaging.HeaderPart headerPart =
    mainPart.AddNewPart<HeaderPart>();

        // Get Id of the headerPart.
        string rId = mainPart.GetIdOfPart(headerPart);

        // Feed target headerPart with source headerPart.
        using (WordprocessingDocument wdDocSource =
            WordprocessingDocument.Open(filepathFrom, true))
        {
            DocumentFormat.OpenXml.Packaging.HeaderPart firstHeader =
    wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

            wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

            if (firstHeader != null)
            {
                headerPart.FeedData(firstHeader.GetStream());
            }
        }

        // Get SectionProperties and Replace HeaderReference with new Id.
        IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SectionProperties> sectPrs =
    mainPart.Document.Body.Elements<SectionProperties>();
        foreach (var sectPr in sectPrs)
        {
            // Delete existing references to headers.
            sectPr.RemoveAllChildren<HeaderReference>();

            // Create the new header reference node.
            sectPr.PrependChild<HeaderReference>(new HeaderReference() { Id = rId });
        }
    }
}