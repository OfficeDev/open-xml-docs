---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 6079a1ae-4567-4d99-b350-b819fd06fe5c
title: 'How to: Insert a new slide into a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 12/03/2024
ms.localizationpriority: high
---
# Insert a new slide into a presentation

This topic shows how to use the classes in the Open XML SDK to
insert a new slide into a presentation programmatically.



## Getting a PresentationDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)> method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read/write,
specify the value `true` for this parameter
as shown in the following `using` statement.
In this code segment, the `presentationFile` parameter is a string that
represents the full path for the file from which you want to open the
document.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/insert_a_new_slideto/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/insert_a_new_slideto/vb/Program.vb#snippet1)]
***


[!include[Using Statement](../includes/presentation/using-statement.md)] `presentationDocument`.

[!include[Structure](../includes/presentation/structure.md)]

## Sample Code

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/insert_a_new_slideto/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/insert_a_new_slideto/vb/Program.vb#snippet)]
***

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
