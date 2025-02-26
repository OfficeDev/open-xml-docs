using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;


// <Snippet>
// Create a new paragraph style with the specified style ID, primary style name, and aliases and 
// add it to the specified style definitions part.
// <Snippet1>
static void CreateAndAddParagraphStyle(StyleDefinitionsPart styleDefinitionsPart, string styleid, string stylename, string aliases = "")
// </Snippet1>
{
    // <Snippet3>
    // Access the root element of the styles part.
    Styles? styles = styleDefinitionsPart.Styles;

    if (styles is null)
    {
        styleDefinitionsPart.Styles = new Styles();
        styles = styleDefinitionsPart.Styles;
    }
    // </Snippet3>

    // <Snippet4>
    // Create a new paragraph style element and specify some of the attributes.
    Style style = new Style()
    {
        Type = StyleValues.Paragraph,
        StyleId = styleid,
        CustomStyle = true,
        Default = false
    };
    // </Snippet4>

    // <Snippet5>
    // Create and add the child elements (properties of the style).
    Aliases aliases1 = new Aliases() { Val = aliases };
    AutoRedefine autoredefine1 = new AutoRedefine() { Val = OnOffOnlyValues.Off };
    BasedOn basedon1 = new BasedOn() { Val = "Normal" };
    LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "OverdueAmountChar" };
    Locked locked1 = new Locked() { Val = OnOffOnlyValues.Off };
    PrimaryStyle primarystyle1 = new PrimaryStyle() { Val = OnOffOnlyValues.On };
    StyleHidden stylehidden1 = new StyleHidden() { Val = OnOffOnlyValues.Off };
    SemiHidden semihidden1 = new SemiHidden() { Val = OnOffOnlyValues.Off };
    StyleName styleName1 = new StyleName() { Val = stylename };
    NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
    UIPriority uipriority1 = new UIPriority() { Val = 1 };
    UnhideWhenUsed unhidewhenused1 = new UnhideWhenUsed() { Val = OnOffOnlyValues.On };

    if (string.IsNullOrWhiteSpace(aliases))
    {
        style.Append(aliases1);
    }

    style.Append(autoredefine1);
    style.Append(basedon1);
    style.Append(linkedStyle1);
    style.Append(locked1);
    style.Append(primarystyle1);
    style.Append(stylehidden1);
    style.Append(semihidden1);
    style.Append(styleName1);
    style.Append(nextParagraphStyle1);
    style.Append(uipriority1);
    style.Append(unhidewhenused1);
    // </Snippet5>

    // <Snippet6>
    // Create the StyleRunProperties object and specify some of the run properties.
    StyleRunProperties styleRunProperties1 = new StyleRunProperties();
    Bold bold1 = new Bold();
    Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
    RunFonts font1 = new RunFonts() { Ascii = "Lucida Console" };
    Italic italic1 = new Italic();

    // Specify a 12 point size.
    FontSize fontSize1 = new FontSize() { Val = "24" };
    styleRunProperties1.Append(bold1);
    styleRunProperties1.Append(color1);
    styleRunProperties1.Append(font1);
    styleRunProperties1.Append(fontSize1);
    styleRunProperties1.Append(italic1);

    // Add the run properties to the style.
    style.Append(styleRunProperties1);

    // Add the style to the styles part.
    styles.Append(style);
    // </Snippet6>
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
    part.Styles = new Styles();

    return part;
}

// <Snippet2>
string strDoc = args[0];

using (WordprocessingDocument doc = WordprocessingDocument.Open(strDoc, true))
{
    if (doc is null)
    {
        throw new ArgumentNullException("document could not be opened");
    }

    MainDocumentPart mainDocumentPart = doc.MainDocumentPart ?? doc.AddMainDocumentPart();

    // Get the Styles part for this document.
    StyleDefinitionsPart? part = mainDocumentPart.StyleDefinitionsPart;

    // If the Styles part does not exist, add it and then add the style.
    if (part is null)
    {
        part = AddStylesPartToPackage(doc);
    }

    // Set up a variable to hold the style ID.
    string parastyleid = "OverdueAmountPara";

    // Create and add a paragraph style to the specified styles part 
    // with the specified style ID, style name and aliases.
    CreateAndAddParagraphStyle(part, parastyleid, "Overdue Amount Para", "Late Due, Late Amount");

    // Add a paragraph with a run and some text.
    Paragraph p =
        new Paragraph(
            new Run(
                new Text("This is some text in a run in a paragraph.")));

    // Add the paragraph as a child element of the w:body element.
    mainDocumentPart.Document ??= new Document();
    mainDocumentPart.Document.Body ??= new Body();

    mainDocumentPart.Document.Body.AppendChild(p);

    // <Snippet7>
    // If the paragraph has no ParagraphProperties object, create one.
    if (p.Elements<ParagraphProperties>().Count() == 0)
    {
        p.PrependChild(new ParagraphProperties());
    }

    // Get a reference to the ParagraphProperties object.
    p.ParagraphProperties ??= new ParagraphProperties();
    ParagraphProperties pPr = p.ParagraphProperties;

    // If a ParagraphStyleId object doesn't exist, create one.
    pPr.ParagraphStyleId ??= new ParagraphStyleId();

    // Set the style of the paragraph.
    pPr.ParagraphStyleId.Val = parastyleid;
    // </Snippet7>
}
// </Snippet2>
// </Snippet>
