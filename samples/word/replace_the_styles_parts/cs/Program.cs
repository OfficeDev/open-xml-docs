#nullable disable

using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Xml;
using System.Xml.Linq;

// Replace the styles in the "to" document with the styles in
// the "from" document.
static void ReplaceStyles(string fromDoc, string toDoc)
{
    // Extract and replace the styles part.
    var node = ExtractStylesPart(fromDoc, false);
    if (node is not null)
        ReplaceStylesPart(toDoc, node, false);

    // Extract and replace the stylesWithEffects part. To fully support 
    // round-tripping from Word 2010 to Word 2007, you should 
    // replace this part, as well.
    node = ExtractStylesPart(fromDoc);
    if (node is not null)
        ReplaceStylesPart(toDoc, node);
    return;
}

// Given a file and an XDocument instance that contains the content of 
// a styles or stylesWithEffects part, replace the styles in the file 
// with the styles in the XDocument.
static void ReplaceStylesPart(string fileName, XDocument newStyles, bool setStylesWithEffectsPart = true)
{
    // Open the document for write access and get a reference.
    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        // Get a reference to the main document part.
        var docPart = document.MainDocumentPart;

        // Assign a reference to the appropriate part to the
        // stylesPart variable.
        StylesPart stylesPart = null;
        if (setStylesWithEffectsPart)
            stylesPart = docPart.StylesWithEffectsPart;
        else
            stylesPart = docPart.StyleDefinitionsPart;

        // If the part exists, populate it with the new styles.
        if (stylesPart is not null)
        {
            newStyles.Save(new StreamWriter(stylesPart.GetStream(
              FileMode.Create, FileAccess.Write)));
        }
    }
}

// Extract the styles or stylesWithEffects part from a 
// word processing document as an XDocument instance.
static XDocument ExtractStylesPart(string fileName, bool getStylesWithEffectsPart = true)
{
    // Declare a variable to hold the XDocument.
    XDocument styles = null;

    // Open the document for read access and get a reference.
    using (var document = WordprocessingDocument.Open(fileName, false))
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
        if (stylesPart is not null)
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