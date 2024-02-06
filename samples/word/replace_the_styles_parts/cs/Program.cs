// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;


// Replace the styles in the "to" document with the styles in
// the "from" document.
// <Snippet1>
static void ReplaceStyles(string fromDoc, string toDoc)
// </Snippet1>
{

    // <Snippet3>
    // Extract and replace the styles part.
    var node = ExtractStylesPart(fromDoc, false);

    if (node is not null)
    {
        ReplaceStylesPart(toDoc, node, false);
    }
    // </Snippet3>

    // <Snippet4>
    // Extract and replace the stylesWithEffects part. To fully support 
    // round-tripping from Word 2010 to Word 2007, you should 
    // replace this part, as well.
    node = ExtractStylesPart(fromDoc);

    if (node is not null)
    {
        ReplaceStylesPart(toDoc, node);
    }

    return;
    // </Snippet4>
}

// Given a file and an XDocument instance that contains the content of 
// a styles or stylesWithEffects part, replace the styles in the file 
// with the styles in the XDocument.

// <Snippet5>
static void ReplaceStylesPart(string fileName, XDocument newStyles, bool setStylesWithEffectsPart = true)
// </Snippet5>
{

    // <Snippet6>
    // Open the document for write access and get a reference.
    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        if (document.MainDocumentPart is null || (document.MainDocumentPart.StyleDefinitionsPart is null && document.MainDocumentPart.StylesWithEffectsPart is null))
        {
            throw new ArgumentNullException("MainDocumentPart and/or one or both of the Styles parts is null.");
        }

        // Get a reference to the main document part.
        var docPart = document.MainDocumentPart;

        // Assign a reference to the appropriate part to the
        // stylesPart variable.

        StylesPart? stylesPart = null;
        // </Snippet6>

        // <Snippet7>
        if (setStylesWithEffectsPart)
        {
            stylesPart = docPart.StylesWithEffectsPart;
        }
        else
        {
            stylesPart = docPart.StyleDefinitionsPart;
        }
        // </Snippet7>

        // <Snippet8>
        // If the part exists, populate it with the new styles.
        if (stylesPart is not null)
        {
            newStyles.Save(new StreamWriter(stylesPart.GetStream(FileMode.Create, FileAccess.Write)));
        }
        // </Snippet8>
    }
}

// Extract the styles or stylesWithEffects part from a 
// word processing document as an XDocument instance.
static XDocument ExtractStylesPart(string fileName, bool getStylesWithEffectsPart = true)
{
    // Declare a variable to hold the XDocument.
    XDocument? styles = null;

    // Open the document for read access and get a reference.
    using (var document = WordprocessingDocument.Open(fileName, false))
    {
        // Get a reference to the main document part.
        var docPart = document.MainDocumentPart;

        if (docPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart is null.");
        }

        // Assign a reference to the appropriate part to the
        // stylesPart variable.
        StylesPart stylesPart;

        if (getStylesWithEffectsPart && docPart.StylesWithEffectsPart is not null)
        {
            stylesPart = docPart.StylesWithEffectsPart;
        }
        else if (docPart.StyleDefinitionsPart is not null)
        {
            stylesPart = docPart.StyleDefinitionsPart;
        }
        else
        {
            throw new ArgumentNullException("StyleWithEffectsPart and StyleDefinitionsPart are undefined");
        }

        using (var reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
        {
            // Create the XDocument.
            styles = XDocument.Load(reader);
        }
    }
    // Return the XDocument instance.
    return styles;
}
// </Snippet0>

// <Snippet2>
string fromDoc = args[0];
string toDoc = args[1];

ReplaceStyles(fromDoc, toDoc);
// </Snippet2>
