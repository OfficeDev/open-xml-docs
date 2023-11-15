using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

AddHeaderFromTo(args[0], args[1]);

static void AddHeaderFromTo(string filepathFrom, string filepathTo)
{
    // Replace header in target document with header of source document.
    using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(filepathTo, true))
    using (WordprocessingDocument wdDocSource = WordprocessingDocument.Open(filepathFrom, true))
    {
        if (wdDocSource.MainDocumentPart is null || wdDocSource.MainDocumentPart.HeaderParts is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or HeaderParts is null.");
        }

        if (wdDoc.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart is null.");
        }

        MainDocumentPart mainPart = wdDoc.MainDocumentPart;

        // Delete the existing header part.
        mainPart.DeleteParts(mainPart.HeaderParts);

        // Create a new header part.
        DocumentFormat.OpenXml.Packaging.HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();

        // Get Id of the headerPart.
        string rId = mainPart.GetIdOfPart(headerPart);

        // Feed target headerPart with source headerPart.

        DocumentFormat.OpenXml.Packaging.HeaderPart? firstHeader = wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

        wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

        if (firstHeader is not null)
        {
            headerPart.FeedData(firstHeader.GetStream());
        }

        if (mainPart.Document.Body is null)
        {
            throw new ArgumentNullException("Body is null.");
        }

        // Get SectionProperties and Replace HeaderReference with new Id.
        IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SectionProperties> sectPrs = mainPart.Document.Body.Elements<SectionProperties>();
        foreach (var sectPr in sectPrs)
        {
            // Delete existing references to headers.
            sectPr.RemoveAllChildren<HeaderReference>();

            // Create the new header reference node.
            sectPr.PrependChild<HeaderReference>(new HeaderReference() { Id = rId });
        }
    }
}