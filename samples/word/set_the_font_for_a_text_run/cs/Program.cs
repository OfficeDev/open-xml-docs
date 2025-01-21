using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;


// Set the font for a text run.
// <Snippet0>
static void SetRunFont(string fileName)
{
    // Open a Wordprocessing document for editing.
    using (WordprocessingDocument package = WordprocessingDocument.Open(fileName, true))
    {
        // <Snippet1>
        // Set the font to Arial to the first Run.
        // Use an object initializer for RunProperties and rPr.
        RunProperties rPr = new RunProperties(
            new RunFonts()
            {
                Ascii = "Arial"
            });
        // </Snippet1>

        // <Snippet2>
        if (package.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart is null.");
        }

        Run r = package.MainDocumentPart.Document.Descendants<Run>().First();
        r.PrependChild<RunProperties>(rPr);
        // </Snippet2>
    }
}
// </Snippet0>

SetRunFont(args[0]);
