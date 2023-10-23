---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 1fbc6d30-bfe4-4b2b-8fd8-0c5a400d1e03
title: Working with runs (Open XML SDK)
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Working with runs (Open XML SDK)

This topic discusses the Open XML SDK **[Run](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.run.aspx)** class and how it relates to the Open
XML File Format WordprocessingML schema.


---------------------------------------------------------------------------------
## Runs in WordprocessingML 
The following text from the [ISO/IEC
29500](https://www.iso.org/standard/71691.html) specification
introduces the Open XML WordprocessingML run element.

The next level of the document hierarchy [after the paragraph] is the
run, which defines a region of text with a common set of properties. A
run is represented by an r element, which allows the producer to combine
breaks, styles, or formatting properties, applying the same information
to all the parts of the run.

Just as a paragraph can have properties, so too can a run. All of the
elements inside an r element have their properties controlled by a
corresponding optional rPr run properties element, which must be the
first child of the r element. In turn, the rPr element is a container
for a set of property elements that are applied to the rest of the
children of the r element. The elements inside the rPr container element
allow the consumer to control whether the text in the following t
elements is bold, underlined, or visible, for example. Some examples of
run properties are bold, border, character style, color, font, font
size, italic, kerning, disable spelling/grammar check, shading, small
caps, strikethrough, text direction, and underline.

© ISO/IEC29500: 2008.

The following table lists the most common Open XML SDK classes used when
working with runs.


| **XML element** | **Open XML SDK Class** |
|-----------------|----------------------------|
|      **p**      |         Paragraph          |
|     **rPr**     |       RunProperties        |
|      **t**      |            Text            |

---------------------------------------------------------------------------------
## Run Class 
The Open XML SDK<strong>[Run](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.run.aspx)</strong> class represents the run (\<**r**\>) element defined in the Open XML File Format
schema for WordprocessingML documents as discussed above. Use a **Run** object to manipulate an individual \<**r**\> element in a WordprocessingML document.

### RunProperties Class

In WordprocessingML, the properties for a run element are specified
using the run properties (\<**rPr**\>) element.
Some examples of run properties are bold, border, character style,
color, font, font size, italic, kerning, disable spelling/grammar check,
shading, small caps, strikethrough, text direction, and underline. Use a
**[RunProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.runproperties.aspx)** object to set the properties
for a run in a WordprocessingML document.

### Text Object

With the \<**r**\> element, the text (\<**t**\>) element is the container for the text that
makes up the document content. The OXML SDK **[Text](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.text.aspx)** class represents the \<**t**\> element. Use a **Text** object to place text in a Wordprocessing
document.


--------------------------------------------------------------------------------
## Open XML SDK Code Example 
The following code adds text to the main document surface of the
specified WordprocessingML document. A **Run**
object demarcates a region of text within the paragraph and then a **RunProperties** object is used to apply bold
formatting to the run.

### [C#](#tab/cs)
[!code-csharp[](../samples/word/working_with_runs/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/working_with_runs/vb/Program.vb)]
When this code is run, the following XML is written to the
WordprocessingML document specified in the preceding code.

```xml
    <w:p>
      <w:r>
        <w:rPr>
          <w:b />
        </w:rPr>
        <w:t>String from WriteToWordDoc method.</w:t>
      </w:r>
    </w:p>
```

--------------------------------------------------------------------------------
## See also 


[About the Open XML SDK for Office](about-the-open-xml-sdk.md)  

[Working with paragraphs (Open XML SDK)](working-with-paragraphs.md)  

[How to: Apply a style to a paragraph in a word processing document (Open XML SDK)](how-to-apply-a-style-to-a-paragraph-in-a-word-processing-document.md)  

[How to: Open and add text to a word processing document (Open XML SDK)](how-to-open-and-add-text-to-a-word-processing-document.md)  
