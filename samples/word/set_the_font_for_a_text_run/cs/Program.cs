#nullable disable

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

// Set the font for a text run.
static void SetRunFont(string fileName)
{
    // Open a Wordprocessing document for editing.
    using (WordprocessingDocument package = WordprocessingDocument.Open(fileName, true))
    {
        // Set the font to Arial to the first Run.
        // Use an object initializer for RunProperties and rPr.
        RunProperties rPr = new RunProperties(
            new RunFonts()
            {
                Ascii = "Arial"
            });

        Run r = package.MainDocumentPart.Document.Descendants<Run>().First();
        r.PrependChild<RunProperties>(rPr);

        // Save changes to the MainDocumentPart part.
        package.MainDocumentPart.Document.Save();
    }
}