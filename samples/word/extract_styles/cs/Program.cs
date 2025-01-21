// <Snippet>
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;


// Extract the styles or stylesWithEffects part from a 
// word processing document as an XDocument instance.
// <Snippet1>
static XDocument? ExtractStylesPart(string fileName, string getStylesWithEffectsPart = "true")
// </Snippet1>
{
    // <Snippet3>
    // Declare a variable to hold the XDocument.
    XDocument? styles = null;

    // Open the document for read access and get a reference.
    using (var document = WordprocessingDocument.Open(fileName, false))
    {
        if (
            document.MainDocumentPart is null ||
            (document.MainDocumentPart.StyleDefinitionsPart is null && document.MainDocumentPart.StylesWithEffectsPart is null)
        )
        {
            throw new ArgumentNullException("MainDocumentPart and/or one or both of the Styles parts is null.");
        }

        // Get a reference to the main document part.
        var docPart = document.MainDocumentPart;

        // Assign a reference to the appropriate part to the
        // stylesPart variable.
        StylesPart? stylesPart = null;
        // </Snippet3>

        // <Snippet4>
        if (getStylesWithEffectsPart.ToLower() == "true")
        {
            stylesPart = docPart.StylesWithEffectsPart;
        }
        else
        {
            stylesPart = docPart.StyleDefinitionsPart;
        }
        // </Snippet4>

        // <Snippet5>
        if (stylesPart is not null)
        {
            using var reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read));

            // Create the XDocument.
            styles = XDocument.Load(reader);
        }
        // </Snippet5>
    }
    // Return the XDocument instance.
    return styles;
}
// </Snippet>

// <Snippet2>
if (args is [{ } fileName, { } getStyleWithEffectsPart])
{
    var styles = ExtractStylesPart(fileName, getStyleWithEffectsPart);

    if (styles is not null)
    {
        Console.WriteLine(styles.ToString());
    }
}
else if (args is [{ } fileName2])
{
    var styles = ExtractStylesPart(fileName2);

    if (styles is not null)
    {
        Console.WriteLine(styles.ToString());
    }
}
// </Snippet2>
