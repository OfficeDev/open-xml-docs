---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: ef817bef-27cd-4c2a-acf3-b7bba17e6e1e
title: 'How to: Move a paragraph from one presentation to another'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 12/02/2024
ms.localizationpriority: medium
---
# Move a paragraph from one presentation to another

This topic shows how to use the classes in the Open XML SDK for
Office to move a paragraph from one presentation to another presentation
programmatically.



## Getting a PresentationDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the
<xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)> method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read/write,
specify the value `true` for this parameter
as shown in the following `using` statement.
In this code, the `sourceFile` parameter is a string that represents the path
for the file from which you want to open the document.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/move_a_paragraph_from_one_presentation_to_another/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/move_a_paragraph_from_one_presentation_to_another/vb/Program.vb#snippet1)]
***


[!include[Using Statement](../includes/presentation/using-statement.md)] `sourceDoc`.

[!include[Structure](../includes/presentation/structure.md)]

## Structure of the Shape Text Body

The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces the structure of this element.

> This element specifies the existence of text to be contained within
> the corresponding shape. All visible text and visible text related
> properties are contained within this element. There can be multiple
> paragraphs and within paragraphs multiple runs of text.
> 
> &copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The following table lists the child elements of the shape text body and
the description of each.

| Child Element | Description |
|---|---|
| bodyPr | Body Properties |
| lstStyle | Text List Styles |
| p | Text Paragraphs |

The following XML Schema fragment defines the contents of this element:

```xml
    <complexType name="CT_TextBody">
       <sequence>
           <element name="bodyPr" type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
           <element name="lstStyle" type="CT_TextListStyle" minOccurs="0" maxOccurs="1"/>
           <element name="p" type="CT_TextParagraph" minOccurs="1" maxOccurs="unbounded"/>
       </sequence>
    </complexType>
```

## How the Sample Code Works

The code in this topic consists of two methods, `MoveParagraphToPresentation` and `GetFirstSlide`. The first method takes two string
parameters: one that represents the source file, which contains the
paragraph to move, and one that represents the target file, to which the
paragraph is moved. The method opens both presentation files and then
calls the `GetFirstSlide` method to get the
first slide in each file. It then gets the first `TextBody` shape in each slide and the first
paragraph in the source shape. It performs a `deep clone` of the source paragraph, 
copying not only the source `Paragraph` object itself, but also everything
contained in that object, including its text. It then inserts the cloned
paragraph in the target file and removes the source paragraph from the
source file, replacing it with a placeholder paragraph. Finally, it
saves the modified slides in both presentations.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/move_a_paragraph_from_one_presentation_to_another/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/move_a_paragraph_from_one_presentation_to_another/vb/Program.vb#snippet2)]
***


The `GetFirstSlide` method takes the `PresentationDocument` object passed in, gets its
presentation part, and then gets the ID of the first slide in its slide
list. It then gets the relationship ID of the slide, gets the slide part
from the relationship ID, and returns the slide part to the calling
method.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/move_a_paragraph_from_one_presentation_to_another/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/move_a_paragraph_from_one_presentation_to_another/vb/Program.vb#snippet3)]
***


## Sample Code

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/move_a_paragraph_from_one_presentation_to_another/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/move_a_paragraph_from_one_presentation_to_another/vb/Program.vb#snippet)]
***

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
