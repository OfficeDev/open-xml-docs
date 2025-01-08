---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 8a9117f7-066e-409c-8681-a26610c0eede
title: Working with paragraphs
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/12/2024
ms.localizationpriority: high
---
# Working with paragraphs

This topic discusses the Open XML SDK <xref:DocumentFormat.OpenXml.Wordprocessing.Paragraph> class and how it relates to the
Open XML File Format WordprocessingML schema.


--------------------------------------------------------------------------------
## Paragraphs in WordprocessingML
The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces the Open XML WordprocessingML element used to represent a
paragraph in a WordprocessingML document.

The most basic unit of block-level content within a WordprocessingML
document, paragraphs are stored using the `<p>` element. A paragraph
defines a distinct division of content that begins on a new line. A
paragraph can contain three pieces of information: optional paragraph
properties, inline content (typically runs), and a set of optional
revision IDs used to compare the content of two documents.

A paragraph's properties are specified via the `<pPr>` element. Some
examples of paragraph properties are alignment, border, hyphenation
override, indentation, line spacing, shading, text direction, and
widow/orphan control.

&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The following table lists the most common Open XML SDK classes used when
working with paragraphs.


| **WordprocessingML element** | **Open XML SDK Class** |
|------------------------------|----------------------------|
|            **p**             |         Paragraph          |
|           **pPr**            |    ParagraphProperties     |
|            **r**             |            Run             |
|            **t**             |            Text            |

---------------------------------------------------------------------------------
## Paragraph Class
The Open XML SDK <xref:DocumentFormat.OpenXml.Wordprocessing.Paragraph> class represents the paragraph
`<p>` element defined in the Open XML
File Format schema for WordprocessingML documents as discussed above.
Use the **Paragraph** object to manipulate
individual `<p>` elements in a
WordprocessingML document.

### ParagraphProperties Class

In WordprocessingML, a paragraph's properties are specified via the
paragraph properties `<pPr>` element.
Some examples of paragraph properties are alignment, border, hyphenation
override, indentation, line spacing, shading, text direction, and
widow/orphan control. The OXML SDK <xref:DocumentFormat.OpenXml.Wordprocessing.ParagraphPropertie> class represents the
`<pPr>` element.

### Run Class

Paragraphs in a word-processing document most often contain text. In the
OXML File Format schema for WordprocessingML documents, the run `<r>` element is provided to demarcate a region of
text. The OXML SDK <xref:DocumentFormat.OpenXml.Wordprocessing.Run> class represents the `<r>` element.

### Text Object

With the `<r>` element, the text `<t>` element is the container for the text that
makes up the document content. The OXML SDK <xref:DocumentFormat.OpenXml.Wordprocessing.Text> class represents the `<t>` element.


--------------------------------------------------------------------------------
## Open XML SDK Code Example
The following code instantiates an Open XML SDK**Paragraph** object and then uses it to add text to
a WordprocessingML document.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/working_with_paragraphs/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/working_with_paragraphs/vb/Program.vb)]

When this code is run, the following XML is written to the
WordprocessingML document referenced in the code.

```xml
    <w:p>
      <w:r>
        <w:t>String from WriteToWordDoc method.</w:t>
      </w:r>
    </w:p>
```

--------------------------------------------------------------------------------
## See also


[About the Open XML SDK for Office](../about-the-open-xml-sdk.md)  

[Working with runs](working-with-runs.md)  

[How to: Apply a style to a paragraph in a word processing document](how-to-apply-a-style-to-a-paragraph-in-a-word-processing-document.md)  

[How to: Open and add text to a word processing document](how-to-open-and-add-text-to-a-word-processing-document.md)  
