---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 7fde676b-81b6-4210-82bf-f74d0d925dec
title: 'How to: Open a spreadsheet document from a stream'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/10/2025
ms.localizationpriority: high
---
# Open a spreadsheet document from a stream

This topic shows how to use the classes in the Open XML SDK for
Office to open a spreadsheet document from a stream programmatically.



---------------------------------------------------------------------------------
## When to Open From a Stream
If you have an application, such as Microsoft SharePoint Foundation
2010, that works with documents by using stream input/output, and you
want to use the Open XML SDK to work with one of the documents, this
is designed to be easy to do. This is especially true if the document
exists and you can open it using the Open XML SDK. However, suppose
that the document is an open stream at the point in your code where you
must use the SDK to work with it? That is the scenario for this topic.
The sample method in the sample code accepts an open stream as a
parameter and then adds text to the document behind the stream using the
Open XML SDK.


--------------------------------------------------------------------------------

[!include[Spreadsheet Object](../includes/spreadsheet/spreadsheet-object.md)]

--------------------------------------------------------------------------------
## Generating the SpreadsheetML Markup to Add a Worksheet

When you have access to the body of the main document part, you add a
worksheet by calling <xref:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.AddNewPart*> method to
create a new <xref:DocumentFormat.OpenXml.Spreadsheet.Worksheet.WorksheetPart*>. The following code example
adds the new `WorksheetPart`.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/open_from_a_stream/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/open_from_a_stream/vb/Program.vb#snippet1)]
***


--------------------------------------------------------------------------------
## Sample Code

In this example, the `OpenAndAddToSpreadsheetStream` method can be used
to open a spreadsheet document from an already open stream and append
some text to it. The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/open_from_a_stream/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/open_from_a_stream/vb/Program.vb#snippet2)]
***


Notice that the `OpenAddAndAddToSpreadsheetStream` method does not
close the stream passed to it. The calling code must do that manually
or with a `using` statement.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/open_from_a_stream/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/open_from_a_stream/vb/Program.vb#snippet0)]
***

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
