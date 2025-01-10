---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 73747a65-0857-4fd4-8362-3613f4169203
title: 'How to: Merge two adjacent cells in a spreadsheet document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/10/2025
ms.localizationpriority: high
---
# Merge two adjacent cells in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to merge two adjacent cells in a spreadsheet document
programmatically.

--------------------------------------------------------------------------------

[!include[Open Spreadsheet](../includes/spreadsheet/open-spreadsheet.md)]

------------------------------------------------------

[!include[Structure](../includes/spreadsheet/structure.md)]


--------------------------------------------------------------------------------

## How the Sample Code Works

After you have opened the spreadsheet file for editing, the code
verifies that the specified cells exist, and if they do not exist, it
creates them by calling the `CreateSpreadsheetCellIfNotExist` method and append
it to the appropriate <xref:DocumentFormat.OpenXml.Spreadsheet.Row> object.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/merge_two_adjacent_cells/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/merge_two_adjacent_cells/vb/Program.vb#snippet1)]
***

In order to get a column name, the code creates a new regular expression
to match the column name portion of the cell name. This regular
expression matches any combination of uppercase or lowercase letters.
For more information about regular expressions, see [Regular Expression Language Elements](/dotnet/standard/base-types/regular-expressions). The
code gets the column name by calling the [Regex.Match](/dotnet/api/system.text.regularexpressions.regex.match#overloads).

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/merge_two_adjacent_cells/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/merge_two_adjacent_cells/vb/Program.vb#snippet2)]
***

To get the row index, the code creates a new regular expression to match the row index portion of the cell name. This regular expression matches any combination of decimal digits. The following code creates a regular expression to match the row index portion of the cell name, comprised of decimal digits.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/merge_two_adjacent_cells/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/merge_two_adjacent_cells/vb/Program.vb#snippet3)]
***

## Sample Code

The following code merges two adjacent cells in a <xref:DocumentFormat.OpenXml.Spreadsheet.Row> document package. When
merging two cells, only the content from one of the cells is preserved.
In left-to-right languages, the content in the upper-left cell is
preserved. In right-to-left languages, the content in the upper-right
cell is preserved.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/merge_two_adjacent_cells/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/merge_two_adjacent_cells/vb/Program.vb#snippet0)]
***

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
