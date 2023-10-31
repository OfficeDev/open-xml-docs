#nullable disable
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Xml;
using System.Xml.Linq;

// Extract the styles or stylesWithEffects part from a 
// word processing document as an XDocument instance.
static XDocument ExtractStylesPart(
  string fileName,
  bool getStylesWithEffectsPart = true)
{
    // Declare a variable to hold the XDocument.
    XDocument styles = null;

    // Open the document for read access and get a reference.
    using (var document =
        WordprocessingDocument.Open(fileName, false))
    {
        // Get a reference to the main document part.
        var docPart = document.MainDocumentPart;

        // Assign a reference to the appropriate part to the
        // stylesPart variable.
        StylesPart stylesPart = null;
        if (getStylesWithEffectsPart)
            stylesPart = docPart.StylesWithEffectsPart;
        else
            stylesPart = docPart.StyleDefinitionsPart;

        // If the part exists, read it into the XDocument.
        if (stylesPart != null)
        {
            using (var reader = XmlNodeReader.Create(
              stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
            {
                // Create the XDocument.
                styles = XDocument.Load(reader);
            }
        }
    }
    // Return the XDocument instance.
    return styles;
}