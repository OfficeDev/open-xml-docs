// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;


static void WDDeleteHiddenText(string docName)
{
    // Given a document name, delete all the hidden text.

    // <Snippet1>
    using (WordprocessingDocument doc = WordprocessingDocument.Open(docName, true))
    {
        // </Snippet1>

        // <Snippet2>
        if (doc.MainDocumentPart is null || doc.MainDocumentPart.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        // Get a list of all the Vanish elements
        List<Vanish> vanishes = doc.MainDocumentPart.Document.Body.Descendants<Vanish>().ToList();
        // </Snippet2>

        // <Snippet3>
        // Loop over the list of Vanish elements
        foreach (Vanish vanish in vanishes)
        {
            var parent = vanish?.Parent;
            var grandparent = parent?.Parent;

            // If the grandparent is a Run remove it
            if (grandparent is Run)
            {
                grandparent.Remove();
            }
            // If it's not a run remove the Vanish
            else if (parent is not null)
            {
                parent.RemoveAllChildren<Vanish>();
            }
        }
        // </Snippet3>
    }
}
// </Snippet0>

string fileName = args[0];

WDDeleteHiddenText(fileName);
