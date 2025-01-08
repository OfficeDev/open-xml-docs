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
ms.date: 12/12/2023
ms.localizationpriority: high
---
# Merge two adjacent cells in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to merge two adjacent cells in a spreadsheet document
programmatically.

--------------------------------------------------------------------------------

[!include[Structure](../includes/spreadsheet/structure.md)]


--------------------------------------------------------------------------------
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

--------------------------------------------------------------------------------
## See also 

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
