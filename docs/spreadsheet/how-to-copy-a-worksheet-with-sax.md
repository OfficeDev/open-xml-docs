---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 2ad4855c-1c83-4dab-b93f-2bae13fac644
title: 'How to: Copy a Worksheet Using SAX (Simple API for XML)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 04/01/2025
ms.localizationpriority: high
---
# Copy a Worksheet Using SAX (Simple API for XML)

This topic shows how to use the the Open XML SDK for Office to programmatically copy a large worksheet
using SAX (Simple API for XML). For more information about the basic structure of a `SpreadsheetML`
document, see [Structure of a SpreadsheetML document](structure-of-a-spreadsheetml-document.md).

------------------------------------
## Why Use the SAX Approach?

The Open XML SDK provides two ways to parse Office Open XML files: the Document Object Model (DOM) and
the Simple API for XML (SAX). The DOM approach is designed to make it easy to query and parse Open XML
files by using strongly-typed classes. However, the DOM approach requires loading entire Open XML parts into
memory, which can lead to slower processing and `Out of Memory` exceptions when working with very large parts.
The SAX approach reads in the XML in an Open XML part one element at a time without reading in the entire part
into memory giving noncached, forward-only access to XML data, which makes it a better choice when reading
very large parts, such as a <xref:DocumentFormat.OpenXml.Packaging.WorksheetPart> with hundreds of thousands of rows.

## Using the DOM Approach

Using the DOM approach, we can take advantage of the Open XML SDK's strongly typed classes. The first step
is to access the package's `WorksheetPart` and make sure that it is not null.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/copy_worksheet_with_sax/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/copy_worksheet_with_sax/vb/Program.vb#snippet1)]
***

Once it is determined that the `WorksheetPart` to be copied is not null, add a new `WorksheetPart` to copy it to.
Then clone the `WorksheetPart`'s <xref:DocumentFormat.OpenXml.Spreadsheet.Worksheet> and assign the cloned
`Worksheet` to the new `WorksheetPart`'s Worksheet property.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/copy_worksheet_with_sax/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/copy_worksheet_with_sax/vb/Program.vb#snippet2)]
***

At this point, the new `WorksheetPart` has been added, but a new <xref:DocumentFormat.OpenXml.Spreadsheet.Sheet>
element must be added to the  `WorkbookPart`'s <xref:DocumentFormat.OpenXml.Spreadsheet.Sheets>'s
child elements for it to display. To do this, first find the new `WorksheetPart`'s Id and
create a new sheet Id by incrementing the `Sheets` count by one then append a new `Sheet`
child to the `Sheets` element. With this, the copied Worksheet is added to the file.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/copy_worksheet_with_sax/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/copy_worksheet_with_sax/vb/Program.vb#snippet3)]
***

## Using the SAX Approach

The SAX approach works on parts, so using the SAX approach, the first step is the same.
Access the package's <xref:DocumentFormat.OpenXml.Packaging.WorksheetPart> and make sure
that it is not null.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/copy_worksheet_with_sax/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/copy_worksheet_with_sax/vb/Program.vb#snippet4)]
***

With SAX, we don't have access to the <xref:DocumentFormat.OpenXml.OpenXmlElement.Clone*>
method. So instead, start by adding a new `WorksheetPart` to the `WorkbookPart`.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/spreadsheet/copy_worksheet_with_sax/cs/Program.cs#snippet5)]

### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/spreadsheet/copy_worksheet_with_sax/vb/Program.vb#snippet5)]
***

Then create an instance of the <xref:DocumentFormat.OpenXml.OpenXmlPartReader> with the
original worksheet part and an instance of the <xref:DocumentFormat.OpenXml.OpenXmlPartWriter>
with the newly created worksheet part.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/spreadsheet/copy_worksheet_with_sax/cs/Program.cs#snippet6)]

### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/spreadsheet/copy_worksheet_with_sax/vb/Program.vb#snippet6)]
***

Then read the elements one by one with the <xref:DocumentFormat.OpenXml.OpenXmlPartReader.Read*>
method. If the element is a <xref:DocumentFormat.OpenXml.Spreadsheet.CellValue> the inner text
needs to be explicitly added using the <xref:DocumentFormat.OpenXml.OpenXmlPartReader.GetText*>
method to read the text, because the <xref:DocumentFormat.OpenXml.OpenXmlPartWriter.WriteStartElement*>
does not write the inner text of an element. For other elements we only need to use the `WriteStartElement`
method, because we don't need the other element's inner text.

### [C#](#tab/cs-7)
[!code-csharp[](../../samples/spreadsheet/copy_worksheet_with_sax/cs/Program.cs#snippet7)]

### [Visual Basic](#tab/vb-7)
[!code-vb[](../../samples/spreadsheet/copy_worksheet_with_sax/vb/Program.vb#snippet7)]
***

At this point, the worksheet part has been copied to the newly added part, but as with the DOM
approach, we still need to add a `Sheet` to the `Workbook`'s `Sheets` element. Because
the SAX approach gives noncached, **forward-only** access to XML data, it is only possible to
prepend element children, which in this case would add the new worksheet to the beginning instead
of the end, changing the order of the worksheets. So the DOM approach is
necessary here, because we want to append not prepend the new `Sheet` and since the `WorkbookPart` is
not usually a large part, the performance gains would be minimal.

### [C#](#tab/cs-8)
[!code-csharp[](../../samples/spreadsheet/copy_worksheet_with_sax/cs/Program.cs#snippet8)]

### [Visual Basic](#tab/vb-8)
[!code-vb[](../../samples/spreadsheet/copy_worksheet_with_sax/vb/Program.vb#snippet8)]
***

## Sample Code

Below is the sample code for both the DOM and SAX approaches to copying the data from one sheet
to a new one and adding it to the Spreadsheet document. While the DOM approach is simpler
and in many cases the preferred choice, with very large documents the SAX approach is better
given that it is faster and can prevent `Out of Memory` exceptions. To see the difference,
create a spreadsheet document with many (10,000+) rows and check the results of the
<xref:System.Diagnostics.Stopwatch> to check the difference in execution time. Increase the
number of rows to 100,000+ to see even more significant performance gains.

### DOM Approach

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/spreadsheet/copy_worksheet_with_sax/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/spreadsheet/copy_worksheet_with_sax/vb/Program.vb#snippet0)]
***

### SAX Approach

### [C#](#tab/cs-99)
[!code-csharp[](../../samples/spreadsheet/copy_worksheet_with_sax/cs/Program.cs#snippet99)]

### [Visual Basic](#tab/vb-99)
[!code-vb[](../../samples/spreadsheet/copy_worksheet_with_sax/vb/Program.vb#snippet99)]
***