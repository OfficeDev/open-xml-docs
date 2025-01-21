---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 6de46612-f864-413f-a504-11ea85f1f88f
title: 'How to: Get all the text in a slide in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 12/06/2024
ms.localizationpriority: medium
---
# Get all the text in a slide in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to get all the text in a slide in a presentation
programmatically.



--------------------------------------------------------------------------------
## Getting a PresentationDocument object
In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the
<xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*>
method that uses a file path, and a Boolean value as the second
parameter to specify whether a document is editable. To open a document
for read/write access, assign the value `true` to this parameter; for read-only access
assign it the value `false` as shown in the
following `using` statement. In this code,
the `file` parameter is a string that
represents the path for the file from which you want to open the
document.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/get_all_the_text_a_slide/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/get_all_the_text_a_slide/vb/Program.vb#snippet1)]
***


[!include[Using Statement](../includes/presentation/using-statement.md)] `presentationDocument`.


--------------------------------------------------------------------------------

[!include[Structure](../includes/presentation/structure.md)]

## How the Sample Code Works

The sample code consists of three overloads of the `GetAllTextInSlide` method. In the following
segment, the first overloaded method opens the source presentation that
contains the slide with text to get, and passes the presentation to the
second overloaded method, which gets the slide part. This method returns
the array of strings that the second method returns to it, each of which
represents a paragraph of text in the specified slide.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/get_all_the_text_a_slide/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/get_all_the_text_a_slide/vb/Program.vb#snippet2)]
***


The second overloaded method takes the presentation document passed in
and gets a slide part to pass to the third overloaded method. It returns
to the first overloaded method the array of strings that the third
overloaded method returns to it, each of which represents a paragraph of
text in the specified slide.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/get_all_the_text_a_slide/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/get_all_the_text_a_slide/vb/Program.vb#snippet3)]
***

The following code segment shows the third overloaded method, which
takes takes the slide part passed in, and returns to the second
overloaded method a string array of text paragraphs. It starts by
verifying that the slide part passed in exists, and then it creates a
linked list of strings. It iterates through the paragraphs in the slide
passed in, and using a `StringBuilder` object
to concatenate all the lines of text in a paragraph, it assigns each
paragraph to a string in the linked list. It then returns to the second
overloaded method an array of strings that represents all the text in
the specified slide in the presentation.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/presentation/get_all_the_text_a_slide/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/presentation/get_all_the_text_a_slide/vb/Program.vb#snippet4)]
***



--------------------------------------------------------------------------------
## Sample Code

Following is the complete sample code that you can use to get all the
text in a specific slide in a presentation file. For example, you can
use the following `foreach` loop in your
program to get the array of strings returned by the method `GetAllTextInSlide`, which represents the text in
the slide at the index of `slideIndex` of the presentation file found at the `filePath`.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/presentation/get_all_the_text_a_slide/cs/Program.cs#snippet5)]

### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/presentation/get_all_the_text_a_slide/vb/Program.vb#snippet5)]
***


Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/get_all_the_text_a_slide/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/get_all_the_text_a_slide/vb/Program.vb#snippet)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
