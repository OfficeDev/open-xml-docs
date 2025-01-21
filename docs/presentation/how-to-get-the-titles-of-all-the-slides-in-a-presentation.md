---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b7d5d1fd-dcdf-4f88-9d57-884562c8144f
title: 'How to: Get the titles of all the slides in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 12/04/2024
ms.localizationpriority: medium
---
# Get the titles of all the slides in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to get the titles of all slides in a presentation
programmatically.



---------------------------------------------------------------------------------
## Getting a PresentationDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the
<xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*>
method that uses a file path, and a Boolean value as the second
parameter to specify whether a document is editable. To open a document
for read-only, specify the value `false` for
this parameter as shown in the following `using` statement. In this code, the `presentationFile` parameter is a string that
represents the path for the file from which you want to open the
document.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/get_the_titles_of_all_the_slides/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/get_the_titles_of_all_the_slides/vb/Program.vb#snippet1)]
***


[!include[Using Statement](../includes/presentation/using-statement.md)] `presentationDocument`.


[!include[Structure](../includes/presentation/structure.md)]

## Sample Code 

The following sample code gets all the
titles of the slides in a presentation file. For example you can use the
following `foreach` statement in your program
to return all the titles in the presentation file located at
the first argument.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/get_the_titles_of_all_the_slides/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/get_the_titles_of_all_the_slides/vb/Program.vb#snippet2)]
***


The result would be a list of the strings that represent the titles in
the presentation, each on a separate line.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/get_the_titles_of_all_the_slides/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/get_the_titles_of_all_the_slides/vb/Program.vb#snippet)]
***

--------------------------------------------------------------------------------
## See also 



[Open XML SDK class library
reference](/office/open-xml/open-xml-sdk)
