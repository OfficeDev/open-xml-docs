#nullable disable
﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

string fileName = args[0];
string authorName = args[1];

using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(fileName, true))
{
    Body body = wdDoc.MainDocumentPart.Document.Body;

    // Handle the formatting changes.
    List<OpenXmlElement> changes = body.Descendants<ParagraphPropertiesChange>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

    foreach (OpenXmlElement change in changes)
    {
        change.Remove();
    }

    // Handle the deletions.
    List<OpenXmlElement> deletions = body
        .Descendants<Deleted>()
        .Where(c => c.Author.Value == authorName)
        .Cast<OpenXmlElement>().ToList();

    deletions.AddRange(body.Descendants<DeletedRun>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

    deletions.AddRange(body.Descendants<DeletedMathControl>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

    foreach (OpenXmlElement deletion in deletions)
    {
        deletion.Remove();
    }

    // Handle the insertions.
    List<OpenXmlElement> insertions =
        body.Descendants<Inserted>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

    insertions.AddRange(body.Descendants<InsertedRun>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

    insertions.AddRange(body.Descendants<InsertedMathControl>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

    foreach (OpenXmlElement insertion in insertions)
    {
        // Found new content.
        // Promote them to the same level as node, and then delete the node.
        foreach (var run in insertion.Elements<Run>())
        {
            if (run == insertion.FirstChild)
            {
                insertion.InsertAfterSelf(new Run(run.OuterXml));
            }
            else
            {
                insertion.NextSibling().InsertAfterSelf(new Run(run.OuterXml));
            }
        }

        insertion.RemoveAttribute("rsidR", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
        insertion.RemoveAttribute("rsidRPr", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
        insertion.Remove();
    }
}