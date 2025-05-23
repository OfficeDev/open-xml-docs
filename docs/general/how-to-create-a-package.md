---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: fe261589-7b04-47df-8ee9-26b444e587b0
title: 'How to: Create a package'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/03/2025
ms.localizationpriority: medium
---

# Create a package

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create a word processing document package
from content in the form of `WordprocessingML` XML markup.

[!include[Structure](../includes/word/packages-and-document-parts.md)]

## Getting a WordprocessingDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class represents a Word document package. To create a Word document, you create an instance
of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class and
populate it with parts. At a minimum, the document must have a main
document part that serves as a container for the main text of the
document. The text is represented in the package as XML using `WordprocessingML` markup.

To create the class instance you call <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(System.String,DocumentFormat.OpenXml.WordprocessingDocumentType)>. Several <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create%2A> methods are
provided, each with a different signature. The first parameter takes a full path
string that represents the document that you want to create. The second
parameter is a member of the <xref:DocumentFormat.OpenXml.WordprocessingDocumentType> enumeration.
This parameter represents the type of document. For example, there is a
different member of the <xref:DocumentFormat.OpenXml.WordprocessingDocumentType> enumeration for each
of document, template, and the macro enabled variety of document and
template.

> [!NOTE]
> Carefully select the appropriate <xref:DocumentFormat.OpenXml.WordprocessingDocumentType> and verify that the persisted file has the correct, matching file extension. If the <xref:DocumentFormat.OpenXml.WordprocessingDocumentType> does not match the file extension, an error occurs when you open the file in Microsoft Word.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/create_a_package/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/create_a_package/vb/Program.vb#snippet1)]
***

[!include[Using Statement](../includes/word/using-statement.md)]

Once you have created the Word document package, you can add parts to
it. To add the main document part you call <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.AddMainDocumentPart%2A>. Having done that,
you can set about adding the document structure and text.

[!include[Structure](../includes/word/structure.md)]

## Sample Code

The following is the complete code sample that you can use to create an
Open XML word processing document package from XML content in the form
of `WordprocessingML` markup. 

After you run the program, open the created file and
examine its content; it should be one paragraph that contains the phrase
"Hello world!"

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/create_a_package/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/create_a_package/vb/Program.vb#snippet0)]
***

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
