#nullable disable
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// Create a new paragraph style with the specified style ID, primary style name, and aliases and 
// add it to the specified style definitions part.
static void CreateAndAddParagraphStyle(StyleDefinitionsPart styleDefinitionsPart,
    string styleid, string stylename, string aliases = "")
{
    // Access the root element of the styles part.
    Styles styles = styleDefinitionsPart.Styles;
    if (styles == null)
    {
        styleDefinitionsPart.Styles = new Styles();
        styleDefinitionsPart.Styles.Save();
    }

    // Create a new paragraph style element and specify some of the attributes.
    Style style = new Style()
    {
        Type = StyleValues.Paragraph,
        StyleId = styleid,
        CustomStyle = true,
        Default = false
    };

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
    if (aliases != "")
        style.Append(aliases1);
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
}

// Add a StylesDefinitionsPart to the document.  Returns a reference to it.
static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
{
    StyleDefinitionsPart part;
    part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
    Styles root = new Styles();
    root.Save(part);
    return part;
}