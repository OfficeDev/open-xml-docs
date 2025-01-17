---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: bb5319c8-ee99-4862-937b-94dcae8deaca
title: 'How to: Change the print orientation of a word processing document'
description: 'Learn how to change the print orientation of a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 05/13/2024
ms.localizationpriority: medium
---

# Change the print orientation of a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically set the print orientation of a Microsoft Word document. It contains an example
`SetPrintOrientation` method to illustrate this task.



-----------------------------------------------------------------------------

## SetPrintOrientation Method

You can use the `SetPrintOrientation` method
to change the print orientation of a word processing document. The
method accepts two parameters that indicate the name of the document to
modify (string) and the new print orientation (<xref:DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues>).

The following code shows the `SetPrintOrientation` method.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/change_the_print_orientation/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/change_the_print_orientation/vb/Program.vb#snippet1)]
***


For each section in the document, if the new orientation differs from
the section's current print orientation, the code modifies the print
orientation for the section. In addition, the code must manually update
the width, height, and margins for each section.

-----------------------------------------------------------------------------

## Calling the Sample SetPrintOrientation Method

To call the sample `SetPrintOrientation`
method, pass a string that contains the name of the file to convert and the string "landscape" or "portrait"
depending on which orientation you want. The following code shows an example method call.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/change_the_print_orientation/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/change_the_print_orientation/vb/Program.vb#snippet2)]
***


-----------------------------------------------------------------------------

## How the Code Works

The following code first determines which orientation to apply and
then opens the document by using the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A>
method and sets the `isEditable` parameter to
`true` to indicate that the document should
be read/write. The code retrieves a reference to the main
document part, and then uses that reference to retrieve a collection of
all of the descendants of type <xref:DocumentFormat.OpenXml.Wordprocessing.SectionProperties> within the content of the
document. Later code will use this collection to set the orientation for
each section in turn.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/change_the_print_orientation/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/change_the_print_orientation/vb/Program.vb#snippet3)]
***


-----------------------------------------------------------------------------

## Iterating Through All the Sections

The next block of code iterates through all the sections in the collection of `SectionProperties` elements. For each section, the code initializes a variable that tracks whether the page orientation for the section was changed so the code can update the page size and margins. (If the new orientation matches the original orientation, the code will not update the page.) The code continues by retrieving a reference to the first <xref:DocumentFormat.OpenXml.Wordprocessing.PageSize> descendant of the `SectionProperties` element. If the reference is not null, the code updates the orientation as required.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/change_the_print_orientation/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/change_the_print_orientation/vb/Program.vb#snippet4)]
***


-----------------------------------------------------------------------------

## Setting the Orientation for the Section

The next block of code first checks whether the <xref:DocumentFormat.OpenXml.Wordprocessing.PageSize.Orient>
property of the `PageSize` element exists. As with many properties
of Open XML elements, the property or attribute might not exist yet. In
that case, retrieving the property returns a null reference. By default,
if the property does not exist, and the new orientation is Portrait, the
code will not update the page. If the `Orient` property already exists, and its value
differs from the new orientation value supplied as a parameter to the
method, the code sets the `Value` property of
the `Orient` property, and sets the
`pageOrientationChanged` flag. (The code uses the `pageOrientationChanged` flag to determine whether it
must update the page size and margins.)

> [!NOTE]
> If the code must create the `Orient` property, it must also create the value to store in the property, as a new <xref:DocumentFormat.OpenXml.EnumValue%601> instance, supplying the new orientation in the `EnumValue` constructor.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/change_the_print_orientation/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/change_the_print_orientation/vb/Program.vb#snippet5)]
***


-----------------------------------------------------------------------------

## Updating the Page Size

At this point in the code, the page orientation may have changed. If so,
the code must complete two more tasks. It must update the page size, and
update the page margins for the section. The first task is easy—the
following code just swaps the page height and width, storing the values
in the `PageSize` element.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/word/change_the_print_orientation/cs/Program.cs#snippet6)]
### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/word/change_the_print_orientation/vb/Program.vb#snippet6)]
***


-----------------------------------------------------------------------------

## Updating the Margins

The next step in the sample procedure handles margins for the section.
If the page orientation has changed, the code must rotate the margins to
match. To do so, the code retrieves a reference to the <xref:DocumentFormat.OpenXml.Wordprocessing.PageMargin> element for the section. If the element exists, the code rotates the margins. Note that the code rotates
the margins by 90 degrees—some printers rotate the margins by 270
degrees instead and you could modify the code to take that into account.
Also be aware that the <xref:DocumentFormat.OpenXml.Wordprocessing.PageMargin.Top> and <xref:DocumentFormat.OpenXml.Wordprocessing.PageMargin.Bottom> properties of the `PageMargin` object are signed values, and the
<xref:DocumentFormat.OpenXml.Wordprocessing.PageMargin.Left> and <xref:DocumentFormat.OpenXml.Wordprocessing.PageMargin.Right> properties are unsigned values. The code must convert between the two types of values as it rotates the
margin settings, as shown in the following code.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/word/change_the_print_orientation/cs/Program.cs#snippet7)]
### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/word/change_the_print_orientation/vb/Program.vb#snippet7)]
***


-----------------------------------------------------------------------------

## Sample Code

The following is the complete `SetPrintOrientation` code sample in C\# and Visual
Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/change_the_print_orientation/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/change_the_print_orientation/vb/Program.vb#snippet0)]
***

-----------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
