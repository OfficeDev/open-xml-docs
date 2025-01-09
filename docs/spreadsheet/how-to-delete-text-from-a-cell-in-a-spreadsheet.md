---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 4b395c48-b469-4d69-b229-d4bad3f3dd8b
title: 'How to: Delete text from a cell in a spreadsheet document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/09/2025
ms.localizationpriority: high
---
# Delete text from a cell in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to delete text from a cell in a spreadsheet document
programmatically.

[!include[Structure](../includes/spreadsheet/structure.md)]

## How the sample code works

In the following code example, you delete text from a cell in a <xref:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument> document package. Then, you verify if other cells within the spreadsheet document still reference the text removed from the row, and if they do not, you remove the text from the <xref:DocumentFormat.OpenXml.Packaging.SharedStringTablePart> object by using the <xref:DocumentFormat.OpenXml.OpenXmlElement.Remove*> method. Then you clean up the `SharedStringTablePart` object by calling the `RemoveSharedStringItem` method.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/delete_text_from_a_cell/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/delete_text_from_a_cell/vb/Program.vb#snippet1)]
***


In the following code example, you verify that the cell specified by the column name and row index exists. If so, the code returns the cell; otherwise, it returns `null`.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/delete_text_from_a_cell/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/delete_text_from_a_cell/vb/Program.vb#snippet2)]
***


In the following code example, you verify if other cells within the
spreadsheet document reference the text specified by the `shareStringId` parameter. If they do not reference
the text, you remove it from the `SharedStringTablePart` object. You do that by
passing a parameter that represents the ID of the text to remove and a
parameter that represents the `SpreadsheetDocument` document package. Then you
iterate through each `Worksheet` object and
compare the contents of each `Cell` object to
the shared string ID. If other cells within the spreadsheet document
still reference the <xref:DocumentFormat.OpenXml.Spreadsheet.SharedStringItem> object, you do not remove
the item from the `SharedStringTablePart`
object. If other cells within the spreadsheet document no longer
reference the `SharedStringItem` object, you
remove the item from the `SharedStringTablePart` object. Then you iterate
through each `Worksheet` object and `Cell` object and refresh the shared string
references. Finally, you save the worksheet and the <xref:DocumentFormat.OpenXml.Spreadsheet.SharedStringTable> object.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/delete_text_from_a_cell/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/delete_text_from_a_cell/vb/Program.vb#snippet3)]
***


## Sample code

The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/delete_text_from_a_cell/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/delete_text_from_a_cell/vb/Program.vb#snippet0)]

## See also

[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
