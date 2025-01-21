---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: ec83a076-9d71-49d1-915f-e7090f74c13a
title: 'How to: Add a new document part to a package'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/08/2025
ms.localizationpriority: medium
---

# Add a new document part to a package

This topic shows how to use the classes in the Open XML SDK for Office to add a document part (file) to a word processing document programmatically.

[!include[Structure](../includes/word/packages-and-document-parts.md)]

## Get a WordprocessingDocument object

The code starts with opening a package file by passing a file name to one of the overloaded <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A> methods of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> that takes a string and a Boolean value that specifies whether the file should be opened for editing or for read-only access. In this case, the Boolean value is `true` specifying that the file should be opened in read/write mode.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/add_a_new_part_to_a_package/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/add_a_new_part_to_a_package/vb/Program.vb#snippet1)]
***


[!include[Using Statement](../includes/word/using-statement.md)]

[!include[Structure](../includes/word/structure.md)]

## How the sample code works

After opening the document for editing, in the `using` statement, as a <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> object, the code creates a reference to the `MainDocumentPart` part and adds a new custom XML part. It then reads the contents of the external
file that contains the custom XML and writes it to the `CustomXmlPart` part.

> [!NOTE]
> To use the new document part in the document, add a link to the document part in the relationship part for the new part.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/add_a_new_part_to_a_package/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/add_a_new_part_to_a_package/vb/Program.vb#snippet2)]
***


## Sample code

The following code adds a new document part that contains custom XML from an external file and then populates the part. To call the `AddCustomXmlPart` method in your program, use the following example that modifies a file by adding a new document part to it.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/add_a_new_part_to_a_package/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/add_a_new_part_to_a_package/vb/Program.vb#snippet3)]
***


> [!NOTE]
> Before you run the program, change the Word file extension from .docx to .zip, and view the content of the zip file. Then change the extension back to .docx and run the program. After running the program, change the file extension again to .zip and view its content. You will see an extra folder named &quot;customXML.&quot; This folder contains the XML file that represents the added part

Following is the complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/add_a_new_part_to_a_package/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/add_a_new_part_to_a_package/vb/Program.vb#snippet0)]
***

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
