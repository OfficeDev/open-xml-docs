---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: debad542-5915-45ad-a71c-eeb95b40ec1a
title: 'How to: Get all the text in all slides in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 12/04/2024
ms.localizationpriority: medium
---
# Get all the text in all slides in a presentation

This topic shows how to use the classes in the Open XML SDK to get
all of the text in all of the slides in a presentation programmatically.



--------------------------------------------------------------------------------
## Getting a PresentationDocument object 

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the
<xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open>
method that uses a file path, and a Boolean value as the second
parameter to specify whether a document is editable. To open a document
for read/write access, assign the value `true` to this parameter; for read-only access
assign it the value `false` as shown in the
following `using` statement. In this code,
the `presentationFile` parameter is a string
that represents the path for the file from which you want to open the
document.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/get_all_the_text_all_slides/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/get_all_the_text_all_slides/vb/Program.vb#snippet1)]
***


[!include[Using Statement](../includes/presentation/using-statement.md)] `presentationDocument`.


--------------------------------------------------------------------------------

[!include[Structure](../includes/presentation/structure.md)]

## Sample Code 
The following code gets all the text in all the slides in a specific
presentation file. For example, you can pass the name of the file as an argument, 
and then use a `foreach` loop in your program to get the array of
strings returned by the method `GetSlideIdAndText` as shown in the following
example.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/get_all_the_text_all_slides/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/get_all_the_text_all_slides/vb/Program.vb#snippet2)]
***

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/get_all_the_text_all_slides/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/get_all_the_text_all_slides/vb/Program.vb#snippet)]
***

--------------------------------------------------------------------------------
## See also 


[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
