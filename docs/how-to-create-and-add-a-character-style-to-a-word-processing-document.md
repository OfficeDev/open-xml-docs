---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: c38f2c94-f0b5-4bb5-8c95-02e556d4e9f1
title: 'Create and add a character style to a word processing document'
description: 'Learn how to create and add a character style to a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: medium
---
# Create and add a character style to a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create and add a character style to a word
processing document. It contains an example
**CreateAndAddCharacterStyle** method to illustrate this task, plus a
supplemental example method to add the styles part when it is necessary.

To use the sample code in this topic, you must install the [Open XML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml). You
must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

## CreateAndAddCharacterStyle Method

The **CreateAndAddCharacterStyle** sample method can be used to add a
style to a word processing document. You must first obtain a reference
to the style definitions part in the document to which you want to add
the style. See the Calling the Sample Method section for an example that
shows how to do this.

The method accepts four parameters that indicate: a reference to the
style definitions part, the style id of the style (an internal
identifier), the name of the style (for external use in the user
interface), and optionally, any style aliases (alternate names for use
in the user interface).

```csharp
    public static void CreateAndAddCharacterStyle(StyleDefinitionsPart styleDefinitionsPart,
        string styleid, string stylename, string aliases="")
```

```vb
    Public Sub CreateAndAddCharacterStyle(ByVal styleDefinitionsPart As StyleDefinitionsPart, 
    ByVal styleid As String, ByVal stylename As String, Optional ByVal aliases As String = "")
```

The complete code listing for the method can be found in the [Sample Code](#sample-code) section.

## About Style IDs, Style Names, and Aliases

The style ID is used by the document to refer to the style, and can be
thought of as its primary identifier. Typically, you use the style ID to
identify a style in code. A style can also have a separate display name
shown in the user interface. Often, the style name, therefore, appears
in proper case and with spacing (for example, Heading 1), while the
style ID is more succinct (for example, heading1) and intended for
internal use. Aliases specify alternate style names that can be used in
an application's user interface.

For example, consider the following XML code example taken from a style
definition.

```xml
    <w:style w:type="character" w:styleId="OverdueAmountChar" . . .
      <w:aliases w:val="Late Due, Late Amount" />
      <w:name w:val="Overdue Amount Char" />
    . . .
    </w:style>
```

The style element styleId attribute defines the main internal identifier
of the style, the style ID (OverdueAmountChar). The aliases element
specifies two alternate style names, Late Due, and Late Amount, which
are comma separated. Each name must be separated by one or more commas.
Finally, the name element specifies the primary style name, which is the
one typically shown in an application's user interface.

## Calling the Sample Method

You can use the **CreateAndAddCharacterStyle**
example method to create and add a named style to a word processing
document using the Open XML SDK. The following code example shows how to
open and obtain a reference to a word processing document, retrieve a
reference to the document's style definitions part, and then call the
**CreateAndAddCharacterStyle** method.

To call the method, you pass a reference to the style definitions part
as the first parameter, the style ID of the style as the second
parameter, the name of the style as the third parameter, and optionally,
any style aliases as the fourth parameter. For example, the following
code example creates the "Overdue Amount Char" character style in a
sample file that is named CreateAndAddCharacterStyle.docx. It also
creates three runs of text in a paragraph, and applies the style to the
second run.

```csharp
    string strDoc = @"C:\Users\Public\Documents\CreateAndAddCharacterStyle.docx";

    using (WordprocessingDocument doc = 
        WordprocessingDocument.Open(strDoc, true))
    {
        // Get the Styles part for this document.
        StyleDefinitionsPart part =
            doc.MainDocumentPart.StyleDefinitionsPart;

        // If the Styles part does not exist, add it.
        if (part == null)
        {
            part = AddStylesPartToPackage(doc);
        }

        // Create and add the character style with the style id, style name, and
        // aliases specified.
        CreateAndAddCharacterStyle(part,
            "OverdueAmountChar",
            "Overdue Amount Char",
            "Late Due, Late Amount");
        
        // Add a paragraph with a run with some text.
        Paragraph p = 
            new Paragraph(
                new Run(
                    new Text("this is some text "){Space = SpaceProcessingModeValues.Preserve}));
        
        // Add another run with some text.
        p.AppendChild<Run>(new Run(new Text("in a run "){Space = SpaceProcessingModeValues.Preserve}));
        
        // Add another run with some text.
        p.AppendChild<Run>(new Run(new Text("in a paragraph."){Space = SpaceProcessingModeValues.Preserve}));

        // Add the paragraph as a child element of the w:body.
        doc.MainDocumentPart.Document.Body.AppendChild(p);

        // Get a reference to the second run (indexed starting with 0).
        Run r = p.Descendants<Run>().ElementAtOrDefault(1);

        // If the Run has no RunProperties object, create one.
        if (r.Elements<RunProperties>().Count() == 0)
        {
            r.PrependChild<RunProperties>(new RunProperties());
        }
        
        // Get a reference to the RunProperties.
        RunProperties rPr = r.RunProperties;
        
        // Set the character style of the run.
        if (rPr.RunStyle == null)
            rPr.RunStyle = new RunStyle();
        rPr.RunStyle.Val = "OverdueAmountChar";
```

```vb
    Dim strDoc As String = "C:\Users\Public\Documents\CreateAndAddCharacterStyle.docx"

    Using doc As WordprocessingDocument =
        WordprocessingDocument.Open(strDoc, True)

        ' Get the Styles part for this document.
        Dim part As StyleDefinitionsPart =
            doc.MainDocumentPart.StyleDefinitionsPart

        ' If the Styles part does not exist, add it.
        If part Is Nothing Then
            part = AddStylesPartToPackage(doc)
        End If

        ' Create and add the character style with the style id, style name, and
        ' aliases specified.
        CreateAndAddCharacterStyle(part,
            "OverdueAmountChar",
            "Overdue Amount Char",
            "Late Due, Late Amount")

        ' Add a paragraph with a run with some text.
        Dim p As New Paragraph(
            New Run(
                New Text("This is some text ") With { _
                    .Space = SpaceProcessingModeValues.Preserve}))

        ' Add another run with some text.
        p.AppendChild(Of Run)(New Run(New Text("in a run ") With { _
            .Space = SpaceProcessingModeValues.Preserve}))

        ' Add another run with some text.
        p.AppendChild(Of Run)(New Run(New Text("in a paragraph.") With { _
            .Space = SpaceProcessingModeValues.Preserve}))

        ' Add the paragraph as a child element of the w:body.
        doc.MainDocumentPart.Document.Body.AppendChild(p)

        ' Get a reference to the second run (indexed starting with 0).
        Dim r As Run = p.Descendants(Of Run)().ElementAtOrDefault(1)

        ' If the Run has no RunProperties object, create one.
        If r.Elements(Of RunProperties)().Count() = 0 Then
            r.PrependChild(Of RunProperties)(New RunProperties())
        End If

        ' Get a reference to the RunProperties.
        Dim rPr As RunProperties = r.RunProperties

        ' Set the character style of the run.
        If rPr.RunStyle Is Nothing Then
            rPr.RunStyle = New RunStyle()
        End If
        rPr.RunStyle.Val = "OverdueAmountChar"

    End Using
```

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
- Linked styles (paragraph + character) Accomplished via the link element (§17.7.4.6).
- Table styles
- Numbering styles
- Default paragraph + character properties

Consider a style called Heading 1 in a document as shown in the following code example.

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

The type attribute has a value of paragraph, which indicates that the following style definition is a paragraph style.

You can set the paragraph, character, table and numbering styles types by specifying the corresponding value in the style element's type attribute.

## Character Style Type

You specify character as the style type by setting the value of the type attribute on the style element to "character".

The following information from section 17.7.9 of the ISO/IEC 29500 specification discusses character styles. Be aware that section numbers preceded by § indicate sections in the ISO specification.

### 17.7.9 Run (Character) Styles

*Character styles* are styles which apply to the contents of one or more
runs of text within a document's contents. This definition implies that
the style can only define character properties (properties which apply
to text within a paragraph) because it cannot be applied to paragraphs.
Character styles can only be referenced by runs within a document, and
they shall be referenced by the rStyle element within a run's run
propertieselement.

A character style has two defining style type-specific characteristics:

- The type attribute on the style has a value of character, which indicates that the following style definition is a character style.

- The style specifies only character-level properties using the rPr element. In this case, the run properties are the set of properties applied to each run which is of this style.

The character style is then applied to runs by referencing the styleId
attribute value for this style in the run properties' rStyle element.

The following image shows some text that has had a character style
applied. A character style can only be applied to a sub-paragraph level
range of text.

Figure 1. Text with a character style applied

![A character style applied to some text](./media/OpenXmlCon_CreateCharacterStyle_Fig1.gif)

## How the Code Works

The **CreateAndAddCharacterStyle** method
begins by retrieving a reference to the styles element in the styles
part. The styles element is the root element of the part and contains
all of the individual style elements. If the reference is null, the
styles element is created and saved to the part.

```csharp
    // Get access to the root element of the styles part.
        Styles styles = styleDefinitionsPart.Styles;
        if (styles == null)
        {
            styleDefinitionsPart.Styles = new Styles();
            styleDefinitionsPart.Styles.Save();
        }
```

```vb
    ' Get access to the root element of the styles part.
        Dim styles As Styles = styleDefinitionsPart.Styles
        If styles Is Nothing Then
            styleDefinitionsPart.Styles = New Styles()
            styleDefinitionsPart.Styles.Save()
        End If
```

## Creating the Style

To create the style, the code instantiates the [Style](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.style.aspx) class and sets certain properties,
such as the [Type](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.style.type.aspx) of style (paragraph), the [StyleId](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.style.styleid.aspx), and whether the style is a [CustomStyle](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.style.customstyle.aspx).

```csharp
    // Create a new character style and specify some of the attributes.
    Style style = new Style()
    {
        Type = StyleValues.Character,
        StyleId = styleid,
        CustomStyle = true
    };
```

```vb
    ' Create a new character style and specify some of the attributes.
    Dim style As New Style() With { _
        .Type = StyleValues.Character, _
        .StyleId = styleid, _
        .CustomStyle = True}
```

The code results in the following XML.

```xml
    <w:style w:type="character" w:styleId="OverdueAmountChar" w:customStyle="true" xmlns:w="https://schemas.openxmlformats.org/wordprocessingml/2006/main">
    </w:style>
```

The code next creates the child elements of the style, which define the
properties of the style. To create an element, you instantiate its
corresponding class, and then call the [Append(\[\])](https://msdn.microsoft.com/library/office/cc801361.aspx) method to add the child element
to the style. For more information about these properties, see section
17.7 of the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification.

```csharp
    // Create and add the child elements (properties of the style).
    Aliases aliases1 = new Aliases() { Val = aliases };
    StyleName styleName1 = new StyleName() { Val = stylename };
    LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "OverdueAmountPara" };
    if (aliases != "")
        style.Append(aliases1);
    style.Append(styleName1);
    style.Append(linkedStyle1);
```

```vb
    ' Create and add the child elements (properties of the style).
    Dim aliases1 As New Aliases() With {.Val = aliases}
    Dim styleName1 As New StyleName() With {.Val = stylename}
    Dim linkedStyle1 As New LinkedStyle() With {.Val = "OverdueAmountPara"}
    If aliases <> "" Then
        style.Append(aliases1)
    End If
    style.Append(styleName1)
    style.Append(linkedStyle1)
```

Next, the code instantiates a [StyleRunProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.stylerunproperties.aspx) object to create a **rPr** (Run Properties) element. You specify the
character properties that apply to the style, such as font and color, in
this element. The properties are then appended as children of the **rPr** element.

When the run properties are created, the code appends the **rPr** element to the style, and the style element
to the styles root element in the styles part.

```csharp
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
```

```vb
    ' Create the StyleRunProperties object and specify some of the run properties.
    Dim styleRunProperties1 As New StyleRunProperties()
    Dim bold1 As New Bold()
    Dim color1 As New Color() With { _
        .ThemeColor = ThemeColorValues.Accent2}
    Dim font1 As New RunFonts() With {.Ascii = "Tahoma"}
    Dim italic1 As New Italic()
    ' Specify a 24 point size.
    Dim fontSize1 As New FontSize() With {.Val = "48"}
    styleRunProperties1.Append(font1)
    styleRunProperties1.Append(fontSize1)
    styleRunProperties1.Append(color1)
    styleRunProperties1.Append(bold1)
    styleRunProperties1.Append(italic1)

    ' Add the run properties to the style.
    style.Append(styleRunProperties1)

    ' Add the style to the styles part.
    styles.Append(style)
```

The following XML shows the final style generated by the code shown here.

```xml
    <w:style w:type="character" w:styleId="OverdueAmountChar" w:customStyle="true" xmlns:w="https://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:aliases w:val="Late Due, Late Amount" />
      <w:name w:val="Overdue Amount Char" />
      <w:link w:val="OverdueAmountPara" />
      <w:rPr>
        <w:rFonts w:ascii="Tahoma" />
        <w:sz w:val="48" />
        <w:color w:themeColor="accent2" />
        <w:b />
        <w:i />
      </w:rPr>
    </w:style>
```

## Applying the Character Style

Once you have the style created, you can apply it to a run by
referencing the styleId attribute value for this style in the run
properties' **rStyle** element. The following
code example shows how to apply a style to a run referenced by the
variable r. The style ID of the style to apply, OverdueAmountChar in
this example, is stored in the RunStyle property of the **rPr** object. This property represents the run
properties' **rStyle** element.

```csharp
    // If the Run has no RunProperties object, create one.
    if (r.Elements<RunProperties>().Count() == 0)
    {
        r.PrependChild<RunProperties>(new RunProperties());
    }

    // Get a reference to the RunProperties.
    RunProperties rPr = r.RunProperties;

    // Set the character style of the run.
    if (rPr.RunStyle == null)
        rPr.RunStyle = new RunStyle();
    rPr.RunStyle.Val = "OverdueAmountChar";
```

```vb
    ' If the Run has no RunProperties object, create one.
    If r.Elements(Of RunProperties)().Count() = 0 Then
        r.PrependChild(Of RunProperties)(New RunProperties())
    End If

    ' Get a reference to the RunProperties.
    Dim rPr As RunProperties = r.RunProperties

    ' Set the character style of the run.
    If rPr.RunStyle Is Nothing Then
        rPr.RunStyle = New RunStyle()
    End If
    rPr.RunStyle.Val = "OverdueAmountChar"
```

## Sample Code

The following is the complete **CreateAndAddCharacterStyle** code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/word/create_and_add_a_character_style/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/create_and_add_a_character_style/vb/Program.vb)]

## See also

[How to: Apply a style to a paragraph in a word processing document (Open XML SDK)](how-to-apply-a-style-to-a-paragraph-in-a-word-processing-document.md)
[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
