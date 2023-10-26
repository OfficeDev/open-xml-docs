---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 73cbca2d-3603-45a5-8a73-c2e718376b01
title: 'How to: Create and add a paragraph style to a word processing document (Open XML SDK)'
description: 'Learn how to create and add a paragraph style to a word processing document using hte Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: high
---
# Create and add a paragraph style to a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create and add a paragraph style to a word
processing document. It contains an example
**CreateAndAddParagraphStyle** method to illustrate this task, plus a
supplemental example method to add the styles part when necessary.

To use the sample code in this topic, you must install the [Open XML SDK]
(https://www.nuget.org/packages/DocumentFormat.OpenXml). You
must also explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

---------------------------------------------------------------------------------

## CreateAndAddParagraphStyle Method

The **CreateAndAddParagraphStyle** sample method can be used to add a
style to a word processing document. You must first obtain a reference
to the style definitions part in the document to which you want to add
the style. For more information and an example of how to do this, see
the [Calling the Sample Method](#calling-the-sample-method)
section.

The method accepts four parameters that indicate: a reference to the
style definitions part, the style ID of the style (an internal
identifier), the name of the style (for external use in the user
interface), and optionally, any style aliases (alternate names for use
in the user interface).

```csharp
    public static void CreateAndAddParagraphStyle(StyleDefinitionsPart styleDefinitionsPart,
        string styleid, string stylename, string aliases="")
```

```vb
    Public Sub CreateAndAddParagraphStyle(ByVal styleDefinitionsPart As StyleDefinitionsPart, 
    ByVal styleid As String, ByVal stylename As String, Optional ByVal aliases As String = "")
```

The complete code listing for the method can be found in the [Sample Code](#sample-code) section.

---------------------------------------------------------------------------------

## About Style IDs, Style Names, and Aliases

The style ID is used by the document to refer to the style, and can be
thought of as its primary identifier. Typically you use the style ID to
identify a style in code. A style can also have a separate display name
in the user interface. Often, the style name therefore appears in proper
case and with spacing (for example, Heading 1), while the style ID is
more succinct (for example, heading1) and intended for internal use.
Aliases specify alternate style names that can be used by the user
interface of an application.

For example, consider the following XML code example taken from a style
definition.

```xml
    <w:style w:type="paragraph" w:styleId="OverdueAmountPara" . . .
      <w:aliases w:val="Late Due, Late Amount" />
      <w:name w:val="Overdue Amount Para" />
    . . .
    </w:style>
```

The styleId attribute of the style element holds the main internal
identifier of the style, the style ID (OverdueAmountPara). The aliases
element specifies two alternate style names, Late Due, and Late Amount,
which are comma separated. Each name must be separated by one or more
commas. Finally, the name element specifies the primary style name,
which is the one typically shown in the user interface of an
application.

---------------------------------------------------------------------------------

## Calling the Sample Method

Use the **CreateAndAddParagraphStyle** example
method to create and add a named style to a word processing document
using the Open XML SDK. The following code example shows how to open and
obtain a reference to a word processing document, retrieve a reference
to the style definitions part of the document, and then call the **CreateAndAddParagraphStyle** method.

To call the method, pass a reference to the style definitions part as
the first parameter, the style ID of the style as the second parameter,
the name of the style as the third parameter, and optionally, any style
aliases as the fourth parameter. For example, the following code creates
the "Overdue Amount Para" paragraph style in a sample file that is named
CreateAndAddParagraphStyle.docx. It also adds a paragraph of text, and
applies the style to the paragraph.

```csharp
    string strDoc = @"C:\Users\Public\Documents\CreateAndAddParagraphStyle.docx";

    using (WordprocessingDocument doc = 
        WordprocessingDocument.Open(strDoc, true))
    {
        // Get the Styles part for this document.
        StyleDefinitionsPart part =
            doc.MainDocumentPart.StyleDefinitionsPart;

        // If the Styles part does not exist, add it and then add the style.
        if (part == null)
        {
            part = AddStylesPartToPackage(doc);
        }

        // Set up a variable to hold the style ID.
        string parastyleid = "OverdueAmountPara";

        // Create and add a paragraph style to the specified styles part 
        // with the specified style ID, style name and aliases.
        CreateAndAddParagraphStyle(part,
            parastyleid,
            "Overdue Amount Para",
            "Late Due, Late Amount");
        
        // Add a paragraph with a run and some text.
        Paragraph p = 
            new Paragraph(
                new Run(
                    new Text("This is some text in a run in a paragraph.")));
        
        // Add the paragraph as a child element of the w:body element.
        doc.MainDocumentPart.Document.Body.AppendChild(p);

        // If the paragraph has no ParagraphProperties object, create one.
        if (p.Elements<ParagraphProperties>().Count() == 0)
        {
            p.PrependChild<ParagraphProperties>(new ParagraphProperties());
        }

        // Get a reference to the ParagraphProperties object.
        ParagraphProperties pPr = p.ParagraphProperties;
        
        // If a ParagraphStyleId object doesn't exist, create one.
        if (pPr.ParagraphStyleId == null)
            pPr.ParagraphStyleId = new ParagraphStyleId();

        // Set the style of the paragraph.
        pPr.ParagraphStyleId.Val = parastyleid;
    }
```

```vb
    Dim strDoc As String = "C:\Users\Public\Documents\CreateAndAddParagraphStyle.docx"

    Using doc As WordprocessingDocument =
        WordprocessingDocument.Open(strDoc, True)

        ' Get the Styles part for this document.
        Dim part As StyleDefinitionsPart =
            doc.MainDocumentPart.StyleDefinitionsPart

        ' If the Styles part does not exist, add it.
        If part Is Nothing Then
            part = AddStylesPartToPackage(doc)
        End If

        ' Set up a variable to hold the style ID.
        Dim parastyleid As String = "OverdueAmountPara"

        ' Create and add a paragraph style to the specified styles part 
        ' with the specified style ID, style name and aliases.
        CreateAndAddParagraphStyle(part,
            parastyleid,
            "Overdue Amount Para",
            "Late Due, Late Amount")

        ' Add a paragraph with a run and some text.
        Dim p As New Paragraph(
            New Run(
                New Text("This is some text in a run in a paragraph.")))

        ' Add the paragraph as a child element of the w:body element.
        doc.MainDocumentPart.Document.Body.AppendChild(p)

        ' If the paragraph has no ParagraphProperties object, create one.
        If p.Elements(Of ParagraphProperties)().Count() = 0 Then
            p.PrependChild(Of ParagraphProperties)(New ParagraphProperties())
        End If

        ' Get a reference to the ParagraphProperties object.
        Dim pPr As ParagraphProperties = p.ParagraphProperties

        ' If a ParagraphStyleId object doesn't exist, create one.
        If pPr.ParagraphStyleId Is Nothing Then
            pPr.ParagraphStyleId = New ParagraphStyleId()
        End If

        ' Set the style of the paragraph.
        pPr.ParagraphStyleId.Val = parastyleid
    End Using
```

---------------------------------------------------------------------------------

## Style Types

WordprocessingML supports six style types, four of which you can specify
using the type attribute on the style element. The following
information, from section 17.7.4.17 in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification,
introduces style types.

*Style types* refers to the property on a style which defines the type
of style created with this style definition. WordprocessingML supports
six types of style definitions by the values for the style definition's
type attribute:

- Paragraph styles

- Character styles

- Linked styles (paragraph + character) [*Note*: Accomplished via the
    link element (§17.7.4.6). *end note*]

- Table styles

- Numbering styles

- Default paragraph + character properties

*Example*: Consider a style called Heading 1 in a document as follows:

```xml
    <w:style w:type="paragraph" w:styleId="Heading1">
      <w:name w:val="heading 1"/>
      <w:basedOn w:val="Normal"/>
      <w:next w:val="Normal"/>
      <w:link w:val="Heading1Char"/>
      <w:uiPriority w:val="1"/>
      <w:qformat/>
      <w:rsid w:val="00F303CE"/>
      …
    </w:style>
```

The type attribute has a value of paragraph, which indicates that the
following style definition is a paragraph style.

© ISO/IEC29500: 2008.

You can set the paragraph, character, table and numbering styles types
by specifying the corresponding value in the type attribute of the style
element.

---------------------------------------------------------------------------------

## Paragraph Style Type

You specify paragraph as the style type by setting the value of the type
attribute on the style element to "paragraph".

The following information from section 17.7.8 of the ISO/IEC 29500
specification discusses paragraph styles. Note that section numbers
preceded by § indicate sections in the ISO specification.

## 17.7.8 Paragraph Styles

*Paragraph styles* are styles which apply to the contents of an entire
paragraph as well as the paragraph mark. This definition implies that
the style can define both character properties (properties which apply
to text within the document) as well as paragraph properties (properties
which apply to the positioning and appearance of the paragraph).
Paragraph styles cannot be referenced by runs within a document; they
shall be referenced by the **pStyle** element
(§17.3.1.27) within a paragraph's paragraph properties element.

A paragraph style has three defining style type-specific
characteristics:

-   The type attribute on the style has a value of paragraph, which
    indicates that the following style definition is a paragraph style.

-   The **next** element defines an editing
    behavior which supplies the paragraph style to be automatically
    applied to the next paragraph when ENTER is pressed at the end of a
    paragraph of this style.

-   The style specifies both paragraph-level and character-level
    properties using the **pPr** and **rPr** elements, respectively. In this case, the
    run properties are the set of properties applied to each run in the
    paragraph.

The paragraph style is then applied to paragraphs by referencing the
styleId attribute value for this style in the paragraph properties'
**pStyle** element.

© ISO/IEC29500: 2008.

---------------------------------------------------------------------------------

## How the Code Works

The **CreateAndAddParagraphStyle** method
begins by retrieving a reference to the styles element in the styles
part. The styles element is the root element of the part and contains
all of the individual style elements. If the reference is null, the
styles element is created and saved to the part.

```csharp
    // Access the root element of the styles part.
        Styles styles = styleDefinitionsPart.Styles;
        if (styles == null)
        {
            styleDefinitionsPart.Styles = new Styles();
            styleDefinitionsPart.Styles.Save();
        }
```

```vb
    ' Access the root element of the styles part.
        Dim styles As Styles = styleDefinitionsPart.Styles
        If styles Is Nothing Then
            styleDefinitionsPart.Styles = New Styles()
            styleDefinitionsPart.Styles.Save()
        End If
```

---------------------------------------------------------------------------------

## Creating the Style

To create the style, the code instantiates the **[Style](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.style.aspx)** class and sets certain properties,
such as the **[Type](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.style.type.aspx)** of style (paragraph), the **[StyleId](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.style.styleid.aspx)**, whether the style is a **[CustomStyle](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.style.customstyle.aspx)**, and whether the style is the
**[Default](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.style.default.aspx)** style for its type.

```csharp
    // Create a new paragraph style element and specify some of the attributes.
    Style style = new Style() { Type = StyleValues.Paragraph,
        StyleId = styleid,
        CustomStyle = true,
        Default = false
    };
```

```vb
    ' Create a new paragraph style element and specify some of the attributes.
    Dim style As New Style() With { .Type = StyleValues.Paragraph, _
     .StyleId = styleid, _
     .CustomStyle = True, _
     .[Default] = False}
```

The code results in the following XML.

```xml
    <w:styles xmlns:w="https://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="OverdueAmountPara" w:default="false" w:customStyle="true">
      </w:style>
    </w:styles>
```

The code next creates the child elements of the style, which define the
properties of the style. To create an element, you instantiate its
corresponding class, and then call the **[Append([])](https://msdn.microsoft.com/library/office/cc801361.aspx)** method add the child element to
the style. For more information about these properties, see section 17.7
of the [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification.

```csharp
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
```

```vb
    ' Create and add the child elements (properties of the style)
    Dim aliases1 As New Aliases() With {.Val = aliases}
    Dim autoredefine1 As New AutoRedefine() With {.Val = OnOffOnlyValues.Off}
    Dim basedon1 As New BasedOn() With {.Val = "Normal"}
    Dim linkedStyle1 As New LinkedStyle() With {.Val = "OverdueAmountChar"}
    Dim locked1 As New Locked() With {.Val = OnOffOnlyValues.Off}
    Dim primarystyle1 As New PrimaryStyle() With {.Val = OnOffOnlyValues.[On]}
    Dim stylehidden1 As New StyleHidden() With {.Val = OnOffOnlyValues.Off}
    Dim semihidden1 As New SemiHidden() With {.Val = OnOffOnlyValues.Off}
    Dim styleName1 As New StyleName() With {.Val = stylename}
    Dim nextParagraphStyle1 As New NextParagraphStyle() With { _
     .Val = "Normal"}
    Dim uipriority1 As New UIPriority() With {.Val = 1}
    Dim unhidewhenused1 As New UnhideWhenUsed() With { _
     .Val = OnOffOnlyValues.[On]}
    If aliases <> "" Then
        style.Append(aliases1)
    End If
    style.Append(autoredefine1)
    style.Append(basedon1)
    style.Append(linkedStyle1)
    style.Append(locked1)
    style.Append(primarystyle1)
    style.Append(stylehidden1)
    style.Append(semihidden1)
    style.Append(styleName1)
    style.Append(nextParagraphStyle1)
    style.Append(uipriority1)
    style.Append(unhidewhenused1)
```

Next, the code instantiates a **[StyleRunProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.stylerunproperties.aspx)** object to create a **rPr** (Run Properties) element. You specify the character properties that apply to the style, such as font and color, in this element. The properties are then appended as children of the **rPr** element.

When the run properties are created, the code appends the **rPr** element to the style, and the style element to the styles root element in the styles part.

```csharp
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
```

```vb
    ' Create the StyleRunProperties object and specify some of the run properties.
    Dim styleRunProperties1 As New StyleRunProperties()
    Dim bold1 As New Bold()
    Dim color1 As New Color() With { _
     .ThemeColor = ThemeColorValues.Accent2}
    Dim font1 As New RunFonts() With { _
     .Ascii = "Lucida Console"}
    Dim italic1 As New Italic()
    ' Specify a 12 point size.
    Dim fontSize1 As New FontSize() With { _
     .Val = "24"}
    styleRunProperties1.Append(bold1)
    styleRunProperties1.Append(color1)
    styleRunProperties1.Append(font1)
    styleRunProperties1.Append(fontSize1)
    styleRunProperties1.Append(italic1)

    ' Add the run properties to the style.
    style.Append(styleRunProperties1)

    ' Add the style to the styles part.
    styles.Append(style)
```

---------------------------------------------------------------------------------

## Applying the Paragraph Style

When you have the style created, you can apply it to a paragraph by
referencing the styleId attribute value for this style in the paragraph
properties' pStyle element. The following code example shows how to
apply a style to a paragraph referenced by the variable p. The style ID
of the style to apply is stored in the parastyleid variable, and the
ParagraphStyleId property represents the paragraph properties' **pStyle** element.

```csharp
    // If the paragraph has no ParagraphProperties object, create one.
    if (p.Elements<ParagraphProperties>().Count() == 0)
    {
        p.PrependChild<ParagraphProperties>(new ParagraphProperties());
    }

    // Get a reference to the ParagraphProperties object.
    ParagraphProperties pPr = p.ParagraphProperties;

    // If a ParagraphStyleId object does not exist, create one.
    if (pPr.ParagraphStyleId == null)
        pPr.ParagraphStyleId = new ParagraphStyleId();

    // Set the style of the paragraph.
    pPr.ParagraphStyleId.Val = parastyleid;
```

```vb
    ' If the paragraph has no ParagraphProperties object, create one.
    If p.Elements(Of ParagraphProperties)().Count() = 0 Then
        p.PrependChild(Of ParagraphProperties)(New ParagraphProperties())
    End If

    ' Get a reference to the ParagraphProperties object.
    Dim pPr As ParagraphProperties = p.ParagraphProperties

    ' If a ParagraphStyleId object does not exist, create one.
    If pPr.ParagraphStyleId Is Nothing Then
        pPr.ParagraphStyleId = New ParagraphStyleId()
    End If

    ' Set the style of the paragraph.
    pPr.ParagraphStyleId.Val = parastyleid
```

---------------------------------------------------------------------------------

## Sample Code

The following is the complete **CreateAndAddParagraphStyle** code sample in both
C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/word/create_and_add_a_paragraph_style/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/create_and_add_a_paragraph_style/vb/Program.vb)]

---------------------------------------------------------------------------------

## See also

- [Apply a style to a paragraph in a word processing document (Open XML SDK)](how-to-apply-a-style-to-a-paragraph-in-a-word-processing-document.md)
- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
