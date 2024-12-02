---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 95fd9dcd-41e9-4e83-9191-2f3110ae73d5
title: 'How to: Move a slide to a new position in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 12/02/2024
ms.localizationpriority: medium
---
# Move a slide to a new position in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to move a slide to a new position in a presentation
programmatically.



--------------------------------------------------------------------------------
## Getting a Presentation Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the
<xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)> method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. In order to count the number of slides
in a presentation, it is best to open the file for read-only access in
order to avoid accidental writing to the file. To do that, specify the
value `false` for the Boolean parameter as
shown in the following `using` statement. In
this code, the `presentationFile` parameter
is a string that represents the path for the file from which you want to
open the document.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/presentation/move_a_slide_to_a_new_position/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/presentation/move_a_slide_to_a_new_position/vb/Program.vb#snippet1)]
***


The `using` statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the `Dispose` method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the `using` statement establishes a scope for the
object that is created or named in the `using` statement, in this case `presentationDocument`.


--------------------------------------------------------------------------------

[!include[Structure](../includes/presentation/structure.md)]

## How the Sample Code Works

In order to move a specific slide in a presentation file to a new
position, you need to know first the number of slides in the
presentation. Therefore, the code in this topic is divided into two
parts. The first is counting the number of slides, and the second is
moving a slide to a new position.


--------------------------------------------------------------------------------

## Counting the Number of Slides

The sample code for counting the number of slides consists of two
overloads of the method `CountSlides`. The
first overload uses a `string` parameter and
the second overload uses a `PresentationDocument` parameter. In the first
`CountSlides` method, the sample code opens
the presentation document in the `using`
statement. Then it passes the `PresentationDocument` object to the second `CountSlides` method, which returns an integer
number that represents the number of slides in the presentation.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/move_a_slide_to_a_new_position/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/move_a_slide_to_a_new_position/vb/Program.vb#snippet2)]
***


In the second `CountSlides` method, the code
verifies that the `PresentationDocument`
object passed in is not `null`, and if it is
not, it gets a `PresentationPart` object from
the `PresentationDocument` object. By using
the `SlideParts` the code gets the `slideCount` and returns it.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/move_a_slide_to_a_new_position/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/move_a_slide_to_a_new_position/vb/Program.vb#snippet3)]
***


--------------------------------------------------------------------------------

## Moving a Slide from one Position to Another

Moving a slide to a new position requires opening the file for
read/write access by specifying the value `true` to the Boolean parameter as shown in the
following `using` statement. The code for
moving a slide consists of two overloads of the `MoveSlide` method. The first overloaded `MoveSlide` method takes three parameters: a string
that represents the presentation file name and path and two integers
that represent the current index position of the slide and the index
position to which to move the slide respectively. It opens the
presentation file, gets a `PresentationDocument` object, and then passes that
object and the two integers, `from` and `to`, to the second overloaded
`MoveSlide` method, which performs the actual
move.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/move_a_slide_to_a_new_position/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/move_a_slide_to_a_new_position/vb/Program.vb#snippet4)]
***


In the second overloaded `MoveSlide` method,
the `CountSlides` method is called to get the
number of slides in the presentation. The code then checks if the
zero-based indexes, `from` and `to`, are within the range and different
from one another.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/presentation/move_a_slide_to_a_new_position/cs/Program.cs#snippet5)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/presentation/move_a_slide_to_a_new_position/vb/Program.vb#snippet5)]
***


A `PresentationPart` object is declared and
set equal to the presentation part of the `PresentationDocument` object passed in. The `PresentationPart` object is used to create a `Presentation` object, and then create a `SlideIdList` object that represents the list of
slides in the presentation from the `Presentation` object. A slide ID of the source
slide (the slide to move) is obtained, and then the position of the
target slide (the slide after which in the slide order to move the
source slide) is identified.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/presentation/move_a_slide_to_a_new_position/cs/Program.cs#snippet6)]

### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/presentation/move_a_slide_to_a_new_position/vb/Program.vb#snippet6)]
***


The `Remove` method of the `SlideID` object is used to remove the source slide
from its current position, and then the `InsertAfter` method of the `SlideIdList` object is used to insert the source
slide in the index position after the target slide. Finally, the
modified presentation is saved.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/presentation/move_a_slide_to_a_new_position/cs/Program.cs#snippet7)]

### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/presentation/move_a_slide_to_a_new_position/vb/Program.vb#snippet7)]
***


--------------------------------------------------------------------------------
## Sample Code
Following is the complete sample code that you can use to move a slide
from one position to another in the same presentation file in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/move_a_slide_to_a_new_position/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/move_a_slide_to_a_new_position/vb/Program.vb#snippet0)]
***

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
