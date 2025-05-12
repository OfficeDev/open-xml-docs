using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

AcceptAllRevisions(args[0], args[1]);

static void AcceptAllRevisions(string fileName, string authorName)
{
    using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(fileName, true))
    {
        if (wdDoc.MainDocumentPart is null || wdDoc.MainDocumentPart.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        Body body = wdDoc.MainDocumentPart.Document.Body;

        // Handle the formatting changes.
        RemoveElements(body.Descendants<ParagraphPropertiesChange>().Where(c => c.Author?.Value == authorName));

        // Handle the deletions.
        RemoveElements(body.Descendants<Deleted>().Where(c => c.Author?.Value == authorName));
        RemoveElements(body.Descendants<DeletedRun>().Where(c => c.Author?.Value == authorName));
        RemoveElements(body.Descendants<DeletedMathControl>().Where(c => c.Author?.Value == authorName));

        // Handle the insertions.
        HandleInsertions(body, authorName);

        // Handle move from elements.
        RemoveElements(body.Descendants<Paragraph>()
            .Where(p => p.Descendants<MoveFrom>()
            .Any(m => m.Author?.Value == authorName)));
        RemoveElements(body.Descendants<MoveFromRangeEnd>());

        // Handle move to elements.
        HandleMoveToElements(body, authorName);
    }
}

// Method to remove elements from the document body
static void RemoveElements(IEnumerable<OpenXmlElement> elements)
{
    foreach (var element in elements.ToList())
    {
        element.Remove();
    }
}

// Method to handle insertions in the document body
static void HandleInsertions(Body body, string authorName)
{
    // Collect all insertion elements by the specified author
    var insertions = body.Descendants<Inserted>().Cast<OpenXmlElement>().ToList();
    insertions.AddRange(body.Descendants<InsertedRun>().Where(c => c.Author?.Value == authorName));
    insertions.AddRange(body.Descendants<InsertedMathControl>().Where(c => c.Author?.Value == authorName));

    foreach (var insertion in insertions)
    {
        // Promote new content to the same level as the node and then delete the node
        foreach (var run in insertion.Elements<Run>())
        {

            if (run == insertion.FirstChild)
            {
                insertion.InsertAfterSelf(new Run(run.OuterXml));
            }
            else
            {
                OpenXmlElement nextSibling = insertion.NextSibling()!;
                nextSibling.InsertAfterSelf(new Run(run.OuterXml));
            }
        }

        // Remove specific attributes and the insertion element itself
        insertion.RemoveAttribute("rsidR", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
        insertion.RemoveAttribute("rsidRPr", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
        insertion.Remove();
    }
}

// Method to handle move-to elements in the document body
static void HandleMoveToElements(Body body, string authorName)
{
    // Collect all move-to elements by the specified author
    var paragraphs = body.Descendants<Paragraph>()
        .Where(p => p.Descendants<MoveFrom>()
          .Any(m => m.Author?.Value == authorName));
    var moveToRun = body.Descendants<MoveToRun>();
    var moveToRangeEnd = body.Descendants<MoveToRangeEnd>();

    List<OpenXmlElement> moveToElements = [.. paragraphs, .. moveToRun, .. moveToRangeEnd];

    foreach (var toElement in moveToElements)
    {
        // Promote new content to the same level as the node and then delete the node
        foreach (var run in toElement.Elements<Run>())
        {
            toElement.InsertBeforeSelf(new Run(run.OuterXml));
        }
        // Remove the move-to element itself
        toElement.Remove();
    }
}
