---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 73cbca2d-3603-45a5-8a73-c2e718376b01
title: 'How to: Create and add a paragraph style to a word processing document'
description: 'Learn how to create and add a paragraph style to a word processing document using hte Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 02/13/2024
ms.localizationpriority: high
---
# Create and add a paragraph style to a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create and add a paragraph style to a word
processing document. It contains an example
`CreateAndAddParagraphStyle` method to illustrate this task, plus a
supplemental example method to add the styles part when necessary.



---------------------------------------------------------------------------------

## CreateAndAddParagraphStyle Method

The `CreateAndAddParagraphStyle` sample method can be used to add a
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

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/create_and_add_a_paragraph_style/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/create_and_add_a_paragraph_style/vb/Program.vb#snippet1)]
***


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
    <w:style w:type="paragraph" w:styleId="OverdueAmountPara" . . .>
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

Use the `CreateAndAddParagraphStyle` example
method to create and add a named style to a word processing document
using the Open XML SDK. The following code example shows how to open and
obtain a reference to a word processing document, retrieve a reference
to the style definitions part of the document, and then call the `CreateAndAddParagraphStyle` method.

To call the method, pass a reference to the style definitions part as
the first parameter, the style ID of the style as the second parameter,
the name of the style as the third parameter, and optionally, any style
aliases as the fourth parameter. For example, the following code creates
the "Overdue Amount Para" paragraph style. It also adds a paragraph of text, and
applies the style to the paragraph.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/create_and_add_a_paragraph_style/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/create_and_add_a_paragraph_style/vb/Program.vb#snippet2)]
***


---------------------------------------------------------------------------------

## Style Types

WordprocessingML supports six style types, four of which you can specify
using the type attribute on the style element. The following
information, from section 17.7.4.17 in the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification,
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

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

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

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

---------------------------------------------------------------------------------

## How the Code Works

The `CreateAndAddParagraphStyle` method
begins by retrieving a reference to the styles element in the styles
part. The styles element is the root element of the part and contains
all of the individual style elements. If the reference is null, the
styles element is created and saved to the part.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/create_and_add_a_paragraph_style/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/create_and_add_a_paragraph_style/vb/Program.vb#snippet3)]
***


---------------------------------------------------------------------------------

## Creating the Style

To create the style, the code instantiates the <xref:DocumentFormat.OpenXml.Wordprocessing.Style>
class and sets certain properties, such as the <xref:DocumentFormat.OpenXml.Wordprocessing.Style.Type>
of style (paragraph), the <xref:DocumentFormat.OpenXml.Wordprocessing.Style.StyleId>, whether the
style is a <xref:DocumentFormat.OpenXml.Wordprocessing.Style.CustomStyle>, and whether the style is the
<xref:DocumentFormat.OpenXml.Wordprocessing.Style.Default> style for its type.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/create_and_add_a_paragraph_style/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/create_and_add_a_paragraph_style/vb/Program.vb#snippet4)]
***


The code results in the following XML.

```xml
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="OverdueAmountPara" w:default="false" w:customStyle="true">
      </w:style>
    </w:styles>
```

The code next creates the child elements of the style, which define the
properties of the style. To create an element, you instantiate its
corresponding class, and then call the <xref:DocumentFormat.OpenXml.OpenXmlElement.Append%2A>
method add the child element to the style. For more information about these properties,
see section 17.7 of the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/create_and_add_a_paragraph_style/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/create_and_add_a_paragraph_style/vb/Program.vb#snippet5)]
***


Next, the code instantiates a <xref:DocumentFormat.OpenXml.Wordprocessing.StyleRunProperties>
object to create a `rPr` (Run Properties) element. You specify the character properties that 
apply to the style, such as font and color, in this element. The properties are then appended
as children of the `rPr` element.

When the run properties are created, the code appends the `rPr` element to the style, and the style element to the styles root element in the styles part.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/word/create_and_add_a_paragraph_style/cs/Program.cs#snippet6)]
### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/word/create_and_add_a_paragraph_style/vb/Program.vb#snippet6)]
***


---------------------------------------------------------------------------------

## Applying the Paragraph Style

When you have the style created, you can apply it to a paragraph by
referencing the styleId attribute value for this style in the paragraph
properties' pStyle element. The following code example shows how to
apply a style to a paragraph referenced by the variable p. The style ID
of the style to apply is stored in the parastyleid variable, and the
ParagraphStyleId property represents the paragraph properties' `pStyle` element.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/word/create_and_add_a_paragraph_style/cs/Program.cs#snippet7)]
### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/word/create_and_add_a_paragraph_style/vb/Program.vb#snippet7)]
***


---------------------------------------------------------------------------------

## Sample Code

The following is the complete `CreateAndAddParagraphStyle` code sample in both
C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/create_and_add_a_paragraph_style/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/create_and_add_a_paragraph_style/vb/Program.vb#snippet)]

---------------------------------------------------------------------------------

## See also

- [Apply a style to a paragraph in a word processing document](how-to-apply-a-style-to-a-paragraph-in-a-word-processing-document.md)
- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
