---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 5ded6212-e8d4-4206-9025-cb5991bd2f80
title: 'How to: Insert text into a cell in a spreadsheet document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/10/2025
ms.localizationpriority: high
---
# Insert text into a cell in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to insert text into a cell in a new worksheet in a spreadsheet
document programmatically.

--------------------------------------------------------------------------------

[!include[Structure](../includes/spreadsheet/structure.md)]

## How the Sample Code Works
After opening the `SpreadsheetDocument`
document for editing, the code inserts a blank <xref:DocumentFormat.OpenXml.Packaging.WorksheetPart.Worksheet*> object into a <xref:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument> document package. Then,
inserts a new <xref:DocumentFormat.OpenXml.Spreadsheet.Cell> object into the new worksheet and
inserts the specified text into that cell.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/insert_textto_a_cell/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/insert_textto_a_cell/vb/Program.vb#snippet1)]
***


The code passes in a parameter that represents the text to insert into
the cell and a parameter that represents the `SharedStringTablePart` object for the spreadsheet.
If the `ShareStringTablePart` object does not
contain a <xref:DocumentFormat.OpenXml.Spreadsheet.SharedStringTable> object, the code creates
one. If the text already exists in the `ShareStringTable` object, the code returns the
index for the <xref:DocumentFormat.OpenXml.Spreadsheet.SharedStringItem> object that represents the
text. Otherwise, it creates a new `SharedStringItem` object that represents the text.

The following code verifies if the specified text exists in the `SharedStringTablePart` object and add the text if
it does not exist.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/insert_textto_a_cell/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/insert_textto_a_cell/vb/Program.vb#snippet2)]
***


The code adds a new `WorksheetPart` object to
the `WorkbookPart` object by using the <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.AddNewPart*> method. It then adds a new `Worksheet` object to the `WorksheetPart` object, and gets a unique ID for
the new worksheet by selecting the maximum <xref:DocumentFormat.OpenXml.Spreadsheet.Sheet.SheetId*> object used within the spreadsheet
document and adding one to create the new sheet ID. It gives the
worksheet a name by concatenating the word "Sheet" with the sheet ID. It
then appends the new `Sheet` object to the
`Sheets` collection.

The following code inserts a new `Worksheet`
object by adding a new `WorksheetPart` object
to the <xref:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.WorkbookPart*> object.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/insert_textto_a_cell/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/insert_textto_a_cell/vb/Program.vb#snippet3)]
***


To insert a cell into a worksheet, the code determines where to insert
the new cell in the column by iterating through the row elements to find
the cell that comes directly after the specified row, in sequential
order. It saves that row in the `refCell`
variable. It then inserts the new cell before the cell referenced by
`refCell` using the <xref:DocumentFormat.OpenXml.OpenXmlCompositeElement.InsertBefore*> method.

In the following code, insert a new `Cell`
object into a `Worksheet` object.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/insert_textto_a_cell/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/insert_textto_a_cell/vb/Program.vb#snippet4)]
***


--------------------------------------------------------------------------------
## Sample Code

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/insert_textto_a_cell/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/insert_textto_a_cell/vb/Program.vb#snippet0)]

--------------------------------------------------------------------------------
## See also


[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
