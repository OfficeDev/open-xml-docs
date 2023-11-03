
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;


    // Apply a style to a paragraph.
    public static void ApplyStyleToParagraph(WordprocessingDocument doc, 
        string styleid, string stylename, Paragraph p)
    {
        // If the paragraph has no ParagraphProperties object, create one.
        if (p.Elements<ParagraphProperties>().Count() == 0)
        {
            p.PrependChild<ParagraphProperties>(new ParagraphProperties());
        }

        // Get the paragraph properties element of the paragraph.
        ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();

        // Get the Styles part for this document.
        StyleDefinitionsPart part =
            doc.MainDocumentPart.StyleDefinitionsPart;

        // If the Styles part does not exist, add it and then add the style.
        if (part == null)
        {
            part = AddStylesPartToPackage(doc);
            AddNewStyle(part, styleid, stylename);
        }
        else
        {
            // If the style is not in the document, add it.
            if (IsStyleIdInDocument(doc, styleid) != true)
            {
                // No match on styleid, so let's try style name.
                string styleidFromName = GetStyleIdFromStyleName(doc, stylename);
                if (styleidFromName == null)
                {
                    AddNewStyle(part, styleid, stylename);
                }
                else
                    styleid = styleidFromName;
            }
        }

        // Set the style of the paragraph.
        pPr.ParagraphStyleId = new ParagraphStyleId(){Val = styleid};
    }
        
    // Return true if the style id is in the document, false otherwise.
    public static bool IsStyleIdInDocument(WordprocessingDocument doc, 
        string styleid)
    {
        // Get access to the Styles element for this document.
        Styles s = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

        // Check that there are styles and how many.
        int n = s.Elements<Style>().Count();
        if (n==0)
            return false;

        // Look for a match on styleid.
        Style style = s.Elements<Style>()
            .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
            .FirstOrDefault();
        if (style == null)
                return false;
        
        return true;
    }

    // Return styleid that matches the styleName, or null when there's no match.
    public static string GetStyleIdFromStyleName(WordprocessingDocument doc, string styleName)
    {
        StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
        string styleId = stylePart.Styles.Descendants<StyleName>()
            .Where(s => s.Val.Value.Equals(styleName) && 
                (((Style)s.Parent).Type == StyleValues.Paragraph))
            .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
        return styleId;
    }

    // Create a new style with the specified styleid and stylename and add it to the specified
    // style definitions part.
    private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, 
        string styleid, string stylename)
    {
        // Get access to the root element of the styles part.
        Styles styles = styleDefinitionsPart.Styles;

        // Create a new paragraph style and specify some of the properties.
        Style style = new Style() { Type = StyleValues.Paragraph, 
            StyleId = styleid, 
            CustomStyle = true };
        StyleName styleName1 = new StyleName() { Val = stylename };
        BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
        NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
        style.Append(styleName1);
        style.Append(basedOn1);
        style.Append(nextParagraphStyle1);

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
    public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
    {
        StyleDefinitionsPart part;
        part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        Styles root = new Styles();
        root.Save(part);
        return part;
    }
