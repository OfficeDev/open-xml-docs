using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

CreateAndAddCharacterStyle(args[0], args[1], args[2], args[3]);

// <Snippet0>
// Create a new character style with the specified style id, style name and aliases and 
// add it to the specified file.
// <Snippet1>
static void CreateAndAddCharacterStyle(string filePath, string styleid, string stylename, string aliases = "")
// </Snippet1>
{
    // <Snippet3>
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filePath, true))
    {
        // Get access to the root element of the styles part.
        Styles? styles = wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles ?? AddStylesPartToPackage(wordprocessingDocument).Styles;
        // </Snippet3>

        if (styles is not null)
        {
            // <Snippet4>
            // Create a new character style and specify some of the attributes.
            Style style = new Style()
            {
                Type = StyleValues.Character,
                StyleId = styleid,
                CustomStyle = true
            };
            // </Snippet4>

            // <Snippet5>
            // Create and add the child elements (properties of the style).
            Aliases aliases1 = new Aliases() { Val = aliases };
            StyleName styleName1 = new StyleName() { Val = stylename };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "OverdueAmountPara" };

            if (!String.IsNullOrEmpty(aliases))
            {
                style.Append(aliases1);
            }

            style.Append(styleName1);
            style.Append(linkedStyle1);
            // </Snippet5>

            // <Snippet6>
            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
            RunFonts font1 = new RunFonts() { Ascii = "Tahoma" };
            Italic italic1 = new Italic();
            // Specify a 24 point size.
            FontSize fontSize1 = new FontSize() { Val = "48" };
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(italic1);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);

            // Add the style to the styles part.
            styles.Append(style);
            // </Snippet6>
        }
    }
}

// Add a StylesDefinitionsPart to the document.  Returns a reference to it.
static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument? doc)
{
    StyleDefinitionsPart part;

    if (doc?.MainDocumentPart is null)
    {
        throw new ArgumentNullException("MainDocumentPart is null.");
    }

    part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
    Styles root = new Styles();
    root.Save(part);
    return part;
}
// </Snippet0>

// <Snippet2>
static void AddStylesToPackage(string filePath)
{
    // Create and add the character style with the style id, style name, and
    // aliases specified.
    CreateAndAddCharacterStyle(
        filePath,
        "OverdueAmountChar",
        "Overdue Amount Char",
        "Late Due, Late Amount");

    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
    {

        // Add a paragraph with a run with some text.
        Paragraph p = new Paragraph(
                new Run(
                    new Text("this is some text ") { Space = SpaceProcessingModeValues.Preserve }));

        // Add another run with some text.
        p.AppendChild<Run>(new Run(new Text("in a run ") { Space = SpaceProcessingModeValues.Preserve }));

        // Add another run with some text.
        p.AppendChild<Run>(new Run(new Text("in a paragraph.") { Space = SpaceProcessingModeValues.Preserve }));

        // Add the paragraph as a child element of the w:body.
        doc?.MainDocumentPart?.Document?.Body?.AppendChild(p);

        // Get a reference to the second run (indexed starting with 0).
        Run r = p.Descendants<Run>().ElementAtOrDefault(1)!;

        // <Snippet7>
        // If the Run has no RunProperties object, create one and get a reference to it.
        RunProperties rPr = r.RunProperties ?? r.PrependChild(new RunProperties());

        // Set the character style of the run.
        if (rPr.RunStyle is null)
        {
            rPr.RunStyle = new RunStyle();
            rPr.RunStyle.Val = "OverdueAmountChar";
        }
        // </Snippet7>
    }
}
// </Snippet2>