using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Xml;
using System.Xml.Linq;

if (args is [{ } fileName, { } getStyleWithEffectsPart])
{
    ExtractStylesPart(fileName, getStyleWithEffectsPart);
}
else if (args is [{ } fileName2])
{
    ExtractStylesPart(fileName2);
}

// Extract the styles or stylesWithEffects part from a 
// word processing document as an XDocument instance.
static XDocument ExtractStylesPart(string fileName, string getStylesWithEffectsPart = "true")
{
    // Declare a variable to hold the XDocument.
    XDocument? styles = null;

    // Open the document for read access and get a reference.
    using (var document = WordprocessingDocument.Open(fileName, false))
    {
        if (document.MainDocumentPart is null || document.MainDocumentPart.StyleDefinitionsPart is null || document.MainDocumentPart.StylesWithEffectsPart is null)
        {
            throw new System.NullReferenceException("MainDocumentPart and/or one or both of the Styles parts is null.");
        }

        // Get a reference to the main document part.
        var docPart = document.MainDocumentPart;

        // Assign a reference to the appropriate part to the
        // stylesPart variable.
        StylesPart? stylesPart = null;

        if (getStylesWithEffectsPart.ToLower() == "true")
            stylesPart = docPart.StylesWithEffectsPart;
        else
            stylesPart = docPart.StyleDefinitionsPart;

        using var reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read));

        // Create the XDocument.
        styles = XDocument.Load(reader);
    }
    // Return the XDocument instance.
    return styles;
}