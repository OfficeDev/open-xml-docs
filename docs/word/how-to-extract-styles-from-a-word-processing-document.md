---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 20258c39-9411-41f2-8463-e94a4b0fa326
title: 'How to: Extract styles from a word processing document'
description: 'Learn how to extract styles from a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 02/09/2024
ms.localizationpriority: medium
---
# Extract styles from a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically extract the styles or stylesWithEffects part
from a word processing document to an <xref:System.Xml.Linq.XDocument>
instance. It contains an example `ExtractStylesPart` method to
illustrate this task.



---------------------------------------------------------------------------------

## ExtractStylesPart Method

You can use the `ExtractStylesPart` sample method to retrieve an `XDocument` instance that contains the styles or
stylesWithEffects part for a Microsoft Word document. Be aware that in a document created in Word 2010, there will
only be a single styles part; Word 2013+ adds a second stylesWithEffects
part. To provide for "round-tripping" a document from Word 2013+ to Word
2010 and back, Word 2013+ maintains both the original styles part and the
new styles part. (The Office Open XML File Formats specification
requires that Microsoft Word ignore any parts that it does not
recognize; Word 2010 does not notice the stylesWithEffects part that
Word 2013+ adds to the document.) You (and your application) must
interpret the results of retrieving the styles or stylesWithEffects
part.

The `ExtractStylesPart` procedure accepts a two parameters: the first
parameter contains a string indicating the path of the file from which
you want to extract styles, and the second indicates whether you want to
retrieve the styles part, or the newer stylesWithEffects part
(basically, you must call this procedure two times for Word 2013+
documents, retrieving each the part). The procedure returns an `XDocument` instance that contains the complete
styles or stylesWithEffects part that you requested, with all the style
information for the document (or a null reference, if the part you
requested does not exist).

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/extract_styles/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/extract_styles/vb/Program.vb#snippet1)]
***


The complete code listing for the method can be found in the [Sample Code](#sample-code) section.

---------------------------------------------------------------------------------

## Calling the Sample Method

To call the sample method, pass a string for the first parameter that
contains the file name of the document from which to extract the styles,
and a Boolean for the second parameter that specifies whether the type
of part to retrieve is the styleWithEffects part (`true`), or the styles part (`false`). The following sample code shows an example.
When you have the `XDocument` instance you
can do what you want with it; in the following sample code the content
of the `XDocument` instance is displayed to
the console.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/extract_styles/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/extract_styles/vb/Program.vb#snippet2)]
***


---------------------------------------------------------------------------------

## How the Code Works

The code starts by creating a variable named `styles` to contain the return value for the method.
The code continues by opening the document by using the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A>
method and indicating that the document should be open for read-only access (the final false
parameter). Given the open document, the code uses the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart>
property to navigate to the main document part, and then prepares a variable named `stylesPart` to hold a reference to the styles part.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/extract_styles/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/extract_styles/vb/Program.vb#snippet3)]
***

---------------------------------------------------------------------------------

## Find the Correct Styles Part

The code next retrieves a reference to the requested styles part by
using the `getStylesWithEffectsPart` <xref:System.Boolean> parameter.
Based on this value, the code retrieves a specific property
of the `docPart` variable, and stores it in the
`stylesPart` variable.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/extract_styles/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/extract_styles/vb/Program.vb#snippet4)]
***


---------------------------------------------------------------------------------

## Retrieve the Part Contents

If the requested styles part exists, the code must return the contents
of the part in an `XDocument` instance. Each part provides a
<xref:DocumentFormat.OpenXml.Packaging.OpenXmlPart.GetStream> method, which returns a Stream.
The code passes the Stream instance to the <xref:System.Xml.XmlReader.Create%2A>
method, and then calls the <xref:System.Xml.Linq.XDocument.Load%2A>
method, passing the `XmlNodeReader` as a parameter.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/word/extract_styles/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/word/extract_styles/vb/Program.vb#snippet5)]
***


---------------------------------------------------------------------------------

## Sample Code

The following is the complete **ExtractStylesPart** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/extract_styles/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/extract_styles/vb/Program.vb#snippet)]

---------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
