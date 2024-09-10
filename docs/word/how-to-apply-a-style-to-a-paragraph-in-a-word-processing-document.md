---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 8d465a77-6c1b-453a-8375-ecf80d2f1bdc
title: 'How to: Apply a style to a paragraph in a word processing document'
description: 'Learn how to apply a style to a paragraph in a word processing document using the Open XML SDK.'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/27/2024
ms.localizationpriority: high
---

# Apply a style to a paragraph in a word processing document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically apply a style to a paragraph within a word processing document. It contains an example `ApplyStyleToParagraph` method to illustrate this task, plus several supplemental example methods to check whether a style exists, add a new style, and add the styles part.



## ApplyStyleToParagraph method

The `ApplyStyleToParagraph` example method can be used to apply a style to a paragraph. You must first obtain a reference to the document as well as a reference to the paragraph that you want to style. The method accepts four parameters that indicate: the path to the word processing document to open, the styleid of the style to be applied, the name of the style to be applied, and the reference to the paragraph to which to apply the style.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/apply_a_style_to_a_paragraph/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/apply_a_style_to_a_paragraph/vb/Program.vb#snippet1)]
***


The following sections in this topic explain the implementation of this method and the supporting code, as well as how to call it. The complete sample code listing can be found in the [Sample Code](#sample-code) section at the end of this topic.

## Get a WordprocessingDocument object

The Sample Code section also shows the code required to set up for calling the sample method. To use the method to apply a style to a paragraph in a document, you first need a reference to the open document. In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class represents a Word document package. To open and work with a Word document, create an instance of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class from the document. After you create the instance, use it to obtain access to the main document part that contains the text of the document. The content in the main document part is represented in the package as XML using WordprocessingML markup.

To create the class instance, call one of the overloads of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open> method. The following sample code shows how to use the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open> overload. The first parameter takes a string that represents the full path to the document to open. The second parameter takes a value of `true` or `false` and represents whether to open the file for editing. In this example the parameter is `true` to enable read/write access to the file.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/apply_a_style_to_a_paragraph/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/apply_a_style_to_a_paragraph/vb/Program.vb#snippet2)]
***


[!include[Structure](../includes/word/structure.md)]

## Get the paragraph to style

After opening the file, the sample code retrieves a reference to the first paragraph. Because a typical word processing document body contains many types of elements, the code filters the descendants in the body of the document to those of type `Paragraph`. The <xref:System.Linq.Enumerable.ElementAtOrDefault> method is then employed to retrieve a reference to the paragraph. Because the elements are indexed starting at zero, you pass a zero to retrieve the reference to the first paragraph, as shown in the following code example.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/apply_a_style_to_a_paragraph/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/apply_a_style_to_a_paragraph/vb/Program.vb#snippet3)]
***


The reference to the found paragraph is stored in a variable named paragraph. If
a paragraph is not found at the specified index, the <xref:System.Linq.Enumerable.ElementAtOrDefault>
method returns null as the default value. This provides an opportunity
to test for null and throw an error with an appropriate error message.

Once you have the references to the document and the paragraph, you can
call the `ApplyStyleToParagraph` example
method to do the remaining work. To call the method, you pass the
reference to the document as the first parameter, the styleid of the
style to apply as the second parameter, the name of the style as the
third parameter, and the reference to the paragraph to which to apply
the style, as the fourth parameter.

## Add the paragraph properties element

The first step of the example method is to ensure that the paragraph has
a paragraph properties element. The paragraph properties element is a
child element of the paragraph and includes a set of properties that
allow you to specify the formatting for the paragraph.

The following information from the ISO/IEC 29500 specification
introduces the `pPr` (paragraph properties)
element used for specifying the formatting of a paragraph. Note that
section numbers preceded by § are from the ISO specification.

Within the paragraph, all rich formatting at the paragraph level is
stored within the `pPr` element (§17.3.1.25; §17.3.1.26). [Note: Some
examples of paragraph properties are alignment, border, hyphenation
override, indentation, line spacing, shading, text direction, and
widow/orphan control.

Among the properties is the `pStyle` element
to specify the style to apply to the paragraph. For example, the
following sample markup shows a pStyle element that specifies the
"OverdueAmount" style.

```xml
    <w:p  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:pPr>
        <w:pStyle w:val="OverdueAmount" />
      </w:pPr>
      ... 
    </w:p>
```

In the Open XML SDK, the `pPr` element is
represented by the <xref:DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties> class. The code
determines if the element exists, and creates a new instance of the
`ParagraphProperties` class if it does not.
The `pPr` element is a child of the `p` (paragraph) element; consequently, the <xref:DocumentFormat.OpenXml.OpenXmlElement.PrependChild> method is used to add
the instance, as shown in the following code example.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/apply_a_style_to_a_paragraph/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/apply_a_style_to_a_paragraph/vb/Program.vb#snippet4)]
***


## Add the Styles part

With the paragraph found and the paragraph properties element present,
now ensure that the prerequisites are in place for applying the style.
Styles in WordprocessingML are stored in their own unique part. Even
though it is typically true that the part as well as a set of base
styles are created automatically when you create the document by using
an application like Microsoft Word, the styles part is not required for
a document to be considered valid. If you create the document
programmatically using the Open XML SDK, the styles part is not created
automatically; you must explicitly create it. Consequently, the
following code verifies that the styles part exists, and creates it if
it does not.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/apply_a_style_to_a_paragraph/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/apply_a_style_to_a_paragraph/vb/Program.vb#snippet5)]
***


The `AddStylesPartToPackage` example method
does the work of adding the styles part. It creates a part of the `StyleDefinitionsPart` type, adding it as a child
to the main document part. The code then appends the `Styles` root element, which is the parent element
that contains all of the styles. The `Styles`
element is represented by the <xref:DocumentFormat.OpenXml.WordProcessing.Styles> class in the Open XML SDK. Finally,
the code saves the part.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/word/apply_a_style_to_a_paragraph/cs/Program.cs#snippet6)]
### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/word/apply_a_style_to_a_paragraph/vb/Program.vb#snippet6)]
***


## Verify that the style exists

Applying a style that does not exist to a paragraph has no effect; there
is no exception generated and no formatting changes occur. The example
code verifies the style exists prior to attempting to apply the style.
Styles are stored in the styles part, therefore if the styles part does
not exist, the style itself cannot exist.

If the styles part exists, the code verifies a matching style by calling
the `IsStyleIdInDocument` example method and
passing the document and the styleid. If no match is found on styleid,
the code then tries to lookup the styleid by calling the `GetStyleIdFromStyleName` example method and
passing it the style name.

If the style does not exist, either because the styles part did not
exist, or because the styles part exists, but the style does not, the
code calls the `AddNewStyle` example method
to add the style.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/word/apply_a_style_to_a_paragraph/cs/Program.cs#snippet7)]
### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/word/apply_a_style_to_a_paragraph/vb/Program.vb#snippet7)]
***


Within the `IsStyleInDocument` example
method, the work begins with retrieving the `Styles` element through the `Styles` property of the 
<xref:DocumentFormat.OpenXml.Packaging.MainDocumentPart.StyleDefinitionsPart> of the main document
part, and then determining whether any styles exist as children of that
element. All style elements are stored as children of the styles
element.

If styles do exist, the code looks for a match on the styleid. The
styleid is an attribute of the style that is used in many places in the
document to refer to the style, and can be thought of as its primary
identifier. Typically you use the styleid to identify a style in code.
The <xref:System.Linq.Enumerable.FirstOrDefault>
method defaults to null if no match is found, so the code verifies for
null to see whether a style was matched, as shown in the following
excerpt.

### [C#](#tab/cs-7)
[!code-csharp[](../../samples/word/apply_a_style_to_a_paragraph/cs/Program.cs#snippet8)]
### [Visual Basic](#tab/vb-7)
[!code-vb[](../../samples/word/apply_a_style_to_a_paragraph/vb/Program.vb#snippet8)]
***


When the style cannot be found based on the styleid, the code attempts
to find a match based on the style name instead. The `GetStyleIdFromStyleName` example method does this
work, looking for a match on style name and returning the styleid for
the matching element if found, or null if not.

## Add the style to the styles part

The `AddNewStyle` example method takes three
parameters. The first parameter takes a reference to the styles part.
The second parameter takes the styleid of the style, and the third
parameter takes the style name. The `AddNewStyle` code creates the named style
definition within the specified part.

To create the style, the code instantiates the <xref:DocumentFormat.OpenXml.Wordprocessing.Style> class and sets certain properties,
such as the <xref:DocumentFormat.OpenXml.Wordprocessing.Style.Type> of style (paragraph) and
<xref:DocumentFormat.OpenXml.Wordprocessing.Style.StyleId>. As mentioned above, the styleid is
used by the document to refer to the style, and can be thought of as its
primary identifier. Typically you use the styleid to identify a style in
code. A style can also have a separate user friendly style name to be
shown in the user interface. Often the style name therefore appears in
proper case and with spacing (for example, Heading 1), while the styleid
is more succinct (for example, heading1) and intended for internal use.
In the following sample code, the styleid and style name take their
values from the styleid and stylename parameters.

The next step is to specify a few additional properties, such as the
style upon which the new style is based, and the style to be
automatically applied to the next paragraph. The code specifies both of
these as the "Normal" style. Be aware that the value to specify here is
the styleid for the normal style. The code appends these properties as
children of the style element.

Once the code has finished instantiating the style and setting up the
basic properties, now work on the style formatting. Style formatting is
performed in the paragraph properties (`pPr`)
and run properties (`rPr`) elements. To set
the font and color characteristics for the runs in a paragraph, you use
the run properties.

To create the `rPr` element with the
appropriate child elements, the code creates an instance of the
<xref:DocumentFormat.OpenXml.Wordprocessing.StyleRunProperties> class and then appends
instances of the appropriate property classes. For this code example,
the style specifies the Lucida Console font, a point size of 12,
rendered in bold and italic, using the Accent2 color from the document
theme part. The point size is specified in half-points, so a value of 24
is equivalent to 12 point.

When the style definition is complete, the code appends the style to the
styles element in the styles part, as shown in the following code
example.

### [C#](#tab/cs-8)
[!code-csharp[](../../samples/word/apply_a_style_to_a_paragraph/cs/Program.cs#snippet9)]
### [Visual Basic](#tab/vb-10)
[!code-vb[](../../samples/word/apply_a_style_to_a_paragraph/vb/Program.vb#snippet9)]
***


The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/apply_a_style_to_a_paragraph/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/apply_a_style_to_a_paragraph/vb/Program.vb#snippet0)]

## See also

- [Open XML SDK class library reference](\/office/open-xml/open-xml-sdk)
