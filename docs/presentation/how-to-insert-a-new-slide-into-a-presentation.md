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

## How the Sample Code Works 

The sample code consists of two overloads of the `InsertNewSlide` method. The first overloaded
method takes three parameters: the full path to the presentation file to
which to add a slide, an integer that represents the zero-based slide
index position in the presentation where to add the slide, and the
string that represents the title of the new slide. It opens the
presentation file as read/write, gets a `PresentationDocument` object, and then passes that
object to the second overloaded `InsertNewSlide` method, which performs the
insertion.

### [C#](#tab/cs-10)
[!code-csharp[](../../samples/presentation/insert_a_new_slideto/cs/Program.cs#snippet10)]

### [Visual Basic](#tab/vb-10)
[!code-vb[](../../samples/presentation/insert_a_new_slideto/vb/Program.vb#snippet10)]
***

The second overloaded `InsertNewSlide` method
creates a new `Slide` object, sets its
properties, and then inserts it into the slide order in the
presentation. The first section of the method creates the slide and sets
its properties.

### [C#](#tab/cs-11)
[!code-csharp[](../../samples/presentation/insert_a_new_slideto/cs/Program.cs#snippet11)]

### [Visual Basic](#tab/vb-11)
[!code-vb[](../../samples/presentation/insert_a_new_slideto/vb/Program.vb#snippet11)]
***

The next section of the second overloaded `InsertNewSlide` method adds a title shape to the
slide and sets its properties, including its text.

### [C#](#tab/cs-12)
[!code-csharp[](../../samples/presentation/insert_a_new_slideto/cs/Program.cs#snippet12)]

### [Visual Basic](#tab/vb-12)
[!code-vb[](../../samples/presentation/insert_a_new_slideto/vb/Program.vb#snippet12)]
***

The next section of the second overloaded `InsertNewSlide` method adds a body shape to the
slide and sets its properties, including its text.

### [C#](#tab/cs-13)
[!code-csharp[](../../samples/presentation/insert_a_new_slideto/cs/Program.cs#snippet13)]

### [Visual Basic](#tab/vb-13)
[!code-vb[](../../samples/presentation/insert_a_new_slideto/vb/Program.vb#snippet13)]
***

The final section of the second overloaded `InsertNewSlide` method creates a new slide part,
finds the specified index position where to insert the slide, and then
inserts it and assigns the new slide to the new slide part.

### [C#](#tab/cs-14)
[!code-csharp[](../../samples/presentation/insert_a_new_slideto/cs/Program.cs#snippet14)]

### [Visual Basic](#tab/vb-14)
[!code-vb[](../../samples/presentation/insert_a_new_slideto/vb/Program.vb#snippet14)]
***


## Sample Code

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/insert_a_new_slideto/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/insert_a_new_slideto/vb/Program.vb#snippet0)]
***

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
