---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: c9b2ce55-548c-4443-8d2e-08fe1f06b7d7
title: 'How to: Add a new document part that receives a relationship ID to a package'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/08/2025
ms.localizationpriority: medium
---

# Add a new document part that receives a relationship ID to a package

This topic shows how to use the classes in the Open XML SDK for
Office to add a document part (file) that receives a relationship `Id` parameter for a word
processing document.



-----------------------------------------------------------------------------
[!include[Structure](../includes/word/packages-and-document-parts.md)]


-----------------------------------------------------------------------------

[!include[Structure](../includes/word/structure.md)]

-----------------------------------------------------------------------------

## How the Sample Code Works

The sample code, in this how-to, starts by passing in a parameter that represents the path to the Word document. It then creates
a new WordprocessingDocument object within a using statement.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/add_a_new_part_that_receives_a_relationship_id_to_a_package/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/add_a_new_part_that_receives_a_relationship_id_to_a_package/vb/Program.vb#snippet1)]
***

It then adds the MainDocumentPart part in the new word processing document, with the relationship ID, rId1. It also adds the `CustomFilePropertiesPart` part and a `CoreFilePropertiesPart` in the new word processing document.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/add_a_new_part_that_receives_a_relationship_id_to_a_package/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/add_a_new_part_that_receives_a_relationship_id_to_a_package/vb/Program.vb#snippet2)]
***

The code then adds the `DigitalSignatureOriginPart` part, the `ExtendedFilePropertiesPart` part, and the `ThumbnailPart` part in the new word processing document with realtionship IDs rId4, rId5, and rId6.

> [!NOTE]
> The <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.AddNewPart*> method creates a relationship from the current document part to the new document part. This method returns the new document part. Also, you can use the <DocumentFormat.OpenXml.Packaging.DataPart.FeedData*> method to fill the document part.

## Sample Code

The following code, adds a new document part that contains custom XML
from an external file and then populates the document part. Below is the
complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/add_a_new_part_that_receives_a_relationship_id_to_a_package/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/add_a_new_part_that_receives_a_relationship_id_to_a_package/vb/Program.vb#snippet0)]
***

-----------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)



