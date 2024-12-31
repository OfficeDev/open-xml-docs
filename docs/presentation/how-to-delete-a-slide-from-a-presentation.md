---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 98781b17-8de4-46e9-b29a-5b4033665491
title: 'How to: Delete a slide from a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Delete a slide from a presentation

This topic shows how to use the Open XML SDK for Office to delete a
slide from a presentation programmatically. It also shows how to delete
all references to the slide from any custom shows that may exist. To
delete a specific slide in a presentation file you need to know first
the number of slides in the presentation. Therefore the code in this
how-to is divided into two parts. The first is counting the number of
slides, and the second is deleting a slide at a specific index.

> [!NOTE]
> Deleting a slide from more complex presentations, such as those that contain outline view settings, for example, may require additional steps.



--------------------------------------------------------------------------------
## Getting a Presentation Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a presentation document package. To work with a presentation document, first create an instance of the `PresentationDocument` class, and then work with that instance. To create the class instance from the document call one of the `Open` method overloads. The code in this topic uses the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)> method, which takes a file path as the first parameter to specify the file to open, and a Boolean value as the second parameter to specify whether a document is editable. Set this second parameter to `false` to open the file for read-only access, or `true` if you want to open the file for read/write access. The code in this topic opens the file twice, once to count the number of slides and once to delete a specific slide. When you count the number of slides in a presentation, it is best to open the file for read-only access to protect the file against accidental writing. The following `using` statement opens the file for read-only access. In this code example, the `presentationFile` parameter is a string that represents the path for the file from which you want to open the document.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/delete_a_slide_from/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/delete_a_slide_from/vb/Program.vb#snippet1)]
***


To delete a slide from the presentation file, open it for read/write
access as shown in the following `using`
statement.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/delete_a_slide_from/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/delete_a_slide_from/vb/Program.vb#snippet2)]
***


[!include[Using Statement](../includes/presentation/using-statement.md)] `presentationDocument`.

[!include[Structure](../includes/presentation/structure.md)]

## Counting the Number of Slides

The sample code consists of two overloads of the `CountSlides` method. The first overload uses a `string` parameter and the second overload uses a `PresentationDocument` parameter. In the first `CountSlides` method, the sample code opens the presentation document in the `using` statement. Then it passes the `PresentationDocument` object to the second `CountSlides` method, which returns an integer number that represents the number of slides in the presentation.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/delete_a_slide_from/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/delete_a_slide_from/vb/Program.vb#snippet3)]
***


In the second `CountSlides` method, the code
verifies that the `PresentationDocument`
object passed in is not `null`, and if it is
not, it gets a `PresentationPart` object from
the `PresentationDocument` object. By using
the `SlideParts` the code gets the slideCount
and returns it.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/presentation/delete_a_slide_from/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/presentation/delete_a_slide_from/vb/Program.vb#snippet4)]
***

--------------------------------------------------------------------------------
## Deleting a Specific Slide

The code for deleting a slide uses two overloads of the `DeleteSlide` method. The first overloaded `DeleteSlide` method takes two parameters: a string
that represents the presentation file name and path, and an integer that
represents the zero-based index position of the slide to delete. It
opens the presentation file for read/write access, gets a `PresentationDocument` object, and then passes that
object and the index number to the next overloaded `DeleteSlide` method, which performs the deletion.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/presentation/delete_a_slide_from/cs/Program.cs#snippet5)]

### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/presentation/delete_a_slide_from/vb/Program.vb#snippet5)]
***


The first section of the second overloaded `DeleteSlide` method uses the `CountSlides` method to get the number of slides in
the presentation. Then, it gets the list of slide IDs in the
presentation, identifies the specified slide in the slide list, and
removes the slide from the slide list.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/presentation/delete_a_slide_from/cs/Program.cs#snippet6)]

### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/presentation/delete_a_slide_from/vb/Program.vb#snippet6)]
***


The next section of the second overloaded `DeleteSlide` method removes all references to the
deleted slide from custom shows. It does that by iterating through the
list of custom shows and through the list of slides in each custom show.
It then declares and instantiates a linked list of slide list entries,
and finds references to the deleted slide by using the relationship ID
of that slide. It adds those references to the list of slide list
entries, and then removes each such reference from the slide list of its
respective custom show.

### [C#](#tab/cs-7)
[!code-csharp[](../../samples/presentation/delete_a_slide_from/cs/Program.cs#snippet7)]

### [Visual Basic](#tab/vb-7)
[!code-vb[](../../samples/presentation/delete_a_slide_from/vb/Program.vb#snippet7)]
***


Finally, the code deletes the slide part for the deleted slide.

### [C#](#tab/cs-8)
[!code-csharp[](../../samples/presentation/delete_a_slide_from/cs/Program.cs#snippet8)]

### [Visual Basic](#tab/vb-8)
[!code-vb[](../../samples/presentation/delete_a_slide_from/vb/Program.vb#snippet8)]
***


--------------------------------------------------------------------------------
## Sample Code

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/delete_a_slide_from/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/delete_a_slide_from/vb/Program.vb#snippet0)]
***

--------------------------------------------------------------------------------
## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
