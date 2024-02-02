// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;


// Remove all of the headers and footers from a document.
// <Snippet1>
static void RemoveHeadersAndFooters(string filename)
// </Snippet1>
{
    // <Snippet3>
    // Given a document name, remove all of the headers and footers
    // from the document.
    using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true))
    // </Snippet3>
    {
        // <Snippet4>
        if (doc.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart is null.");
        }

        // Get a reference to the main document part.
        var docPart = doc.MainDocumentPart;

        // Count the header and footer parts and continue if there 
        // are any.
        if (docPart.HeaderParts.Count() > 0 || docPart.FooterParts.Count() > 0)
        {
            // </Snippet4>

            // <Snippet5>
            // Remove the header and footer parts.
            docPart.DeleteParts(docPart.HeaderParts);
            docPart.DeleteParts(docPart.FooterParts);
            // </Snippet5>

            // <Snippet6>
            // Get a reference to the root element of the main
            // document part.
            Document document = docPart.Document;

            // Remove all references to the headers and footers.

            // First, create a list of all descendants of type
            // HeaderReference. Then, navigate the list and call
            // Remove on each item to delete the reference.
            var headers = document.Descendants<HeaderReference>().ToList();

            foreach (var header in headers)
            {
                header.Remove();
            }

            // First, create a list of all descendants of type
            // FooterReference. Then, navigate the list and call
            // Remove on each item to delete the reference.
            var footers = document.Descendants<FooterReference>().ToList();

            foreach (var footer in footers)
            {
                footer.Remove();
            }

            // Save the changes.
            document.Save();
            // </Snippet6>
        }
    }
}
// </Snippet0>

// <Snippet2>
string filename = args[0];

RemoveHeadersAndFooters(filename);
// </Snippet2>

