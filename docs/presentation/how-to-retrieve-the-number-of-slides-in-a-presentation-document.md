---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b6f429a7-4489-4155-b713-2139f3add8c2
title: 'How to: Retrieve the number of slides in a presentation document'
description: 'Learn how to retrieve the number of slides in a presentation document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/27/2024
ms.localizationpriority: medium
---
# Retrieve the number of slides in a presentation document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve the number of slides in a
presentation document, either including hidden slides or not, without
loading the document into Microsoft PowerPoint. It contains an example
`RetrieveNumberOfSlides` method to illustrate
this task.



---------------------------------------------------------------------------------

## RetrieveNumberOfSlides Method

You can use the `RetrieveNumberOfSlides`
method to get the number of slides in a presentation document,
optionally including the hidden slides. The `RetrieveNumberOfSlides` method accepts two
parameters: a string that indicates the path of the file that you want
to examine, and an optional Boolean value that indicates whether to
include hidden slides in the count.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/presentation/retrieve_the_number_of_slides/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/presentation/retrieve_the_number_of_slides/vb/Program.vb#snippet1)]
***


---------------------------------------------------------------------------------
## Calling the RetrieveNumberOfSlides Method

The method returns an integer that indicates the number of slides,
counting either all the slides or only visible slides, depending on the
second parameter value. To call the method, pass all the parameter
values, as shown in the following code.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/retrieve_the_number_of_slides/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/retrieve_the_number_of_slides/vb/Program.vb#snippet2)]
***



---------------------------------------------------------------------------------

## How the Code Works

The code starts by creating an integer variable, `slidesCount`, to hold the number of slides. The code then opens the specified presentation by using the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*> method and indicating that the document should be open for read-only access (the
final `false` parameter value). Given the open presentation, the code uses the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.PresentationPart*> property to navigate to the main presentation part, storing the reference in a variable named `presentationPart`.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/retrieve_the_number_of_slides/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/retrieve_the_number_of_slides/vb/Program.vb#snippet3)]
***



---------------------------------------------------------------------------------

## Retrieving the Count of All Slides

If the presentation part reference is not null (and it will not be, for any valid presentation that loads correctly into PowerPoint), the code next calls the `Count` method on the value of the <xref:DocumentFormat.OpenXml.Packaging.PresentationPart.SlideParts*> property of the presentation part. If you requested all slides, including hidden slides, that is all there is to do. There is slightly more work to be done if you want to exclude hidden slides, as shown in the following code.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/retrieve_the_number_of_slides/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/retrieve_the_number_of_slides/vb/Program.vb#snippet4)]
***



---------------------------------------------------------------------------------

## Retrieving the Count of Visible Slides

If you requested that the code should limit the return value to include
only visible slides, the code must filter its collection of slides to
include only those slides that have a <xref:DocumentFormat.OpenXml.Presentation.Slide.Show*> property that contains a value, and
the value is `true`. If the `Show` property is null, that also indicates that
the slide is visible. This is the most likely scenario. PowerPoint does
not set the value of this property, in general, unless the slide is to
be hidden. The only way the `Show` property
would exist and have a value of `true` would
be if you had hidden and then unhidden the slide. The following code
uses the <xref:System.Linq.Enumerable.Where*>
function with a lambda expression to do the work.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/retrieve_the_number_of_slides/cs/Program.cs#snippet5)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/retrieve_the_number_of_slides/vb/Program.vb#snippet5)]
***


---------------------------------------------------------------------------------

## Sample Code

The following is the complete `RetrieveNumberOfSlides` code sample in C\# and
Visual Basic.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/retrieve_the_number_of_slides/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/retrieve_the_number_of_slides/vb/Program.vb#snippet0)]

---------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
