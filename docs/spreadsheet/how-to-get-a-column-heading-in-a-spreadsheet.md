---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 56ba8cee-d789-4a03-b8ff-b161af0788ff
title: 'How to: Get a column heading in a spreadsheet document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/03/2024
ms.localizationpriority: high
---
# Get a column heading in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to retrieve a column heading in a spreadsheet document
programmatically.

[!include[Structure](../includes/spreadsheet/structure.md)]

## How the Sample Code Works

The code in this how-to consists of three methods (functions in Visual
Basic): **GetColumnHeading**, **GetColumnName**, and **GetRowIndex**. The last two methods are called from
within the **GetColumnHeading** method.

The **GetColumnName** method takes the cell
name as a parameter. It parses the cell name to get the column name by
creating a regular expression to match the column name portion of the
cell name. For more information about regular expressions, see [Regular Expression Language Elements](/dotnet/standard/base-types/regular-expression-language-quick-reference).

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/get_a_column_heading/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/get_a_column_heading/vb/Program.vb#snippet1)]
***


The **GetRowIndex** method takes the cell name
as a parameter. It parses the cell name to get the row index by creating
a regular expression to match the row index portion of the cell name.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/get_a_column_heading/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/get_a_column_heading/vb/Program.vb#snippet2)]
***


The **GetColumnHeading** method uses three
parameters, the full path to the source spreadsheet file, the name of
the worksheet that contains the specified column, and the name of a cell
in the column for which to get the heading.

The code gets the name of the column of the specified cell by calling
the **GetColumnName** method. The code also
gets the cells in the column and orders them by row using the **GetRowIndex** method.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/get_a_column_heading/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/get_a_column_heading/vb/Program.vb#snippet3)]
***


If the specified column exists, it gets the first cell in the column
using the
[IEnumerable(T).First](/dotnet/api/system.linq.enumerable.first)
method. The first cell contains the heading. Otherwise the specified column does not exist and the method returns `null` / `Nothing`

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/get_a_column_heading/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/get_a_column_heading/vb/Program.vb#snippet4)]
***


If the content of the cell is stored in the <xref:DocumentFormat.OpenXml.Packaging.SharedStringTablePart> object, it gets the
shared string items and returns the content of the column heading using
the
[M:System.Int32.Parse(System.String)](/dotnet/api/system.int32.parse)
method. If the content of the cell is not in the <xref:DocumentFormat.OpenXml.Spreadsheet.SharedStringTable> object, it returns the
content of the cell.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/spreadsheet/get_a_column_heading/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/spreadsheet/get_a_column_heading/vb/Program.vb#snippet5)]
***


## Sample Code

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/get_a_column_heading/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/get_a_column_heading/vb/Program.vb#snippet0)]

## See also



[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
