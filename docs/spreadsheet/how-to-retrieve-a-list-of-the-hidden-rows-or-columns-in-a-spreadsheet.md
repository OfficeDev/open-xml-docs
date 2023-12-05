---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 5adddb6e-545e-4fba-ae35-cc4682e3eda7
title: 'How to: Retrieve a list of the hidden rows or columns in a spreadsheet document'
description: 'Learn how to retrieve a list of the hidden rows or columns in a spreadsheet document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: high
---
# Retrieve a list of the hidden rows or columns in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically retrieve a list of hidden rows or columns in a Microsoft Excel worksheet. It contains an example **GetHiddenRowsOrCols** method to illustrate this task.



---------------------------------------------------------------------------------

## GetHiddenRowsOrCols Method

You can use the **GetHiddenRowsOrCols** method to retrieve a list of the hidden rows or columns in a worksheet. The method returns a list of unsigned integers that contain each index for the hidden rows or columns, if the specified worksheet contains any hidden rows or columns (rows and columns are numbered starting at 1, rather than 0). The **GetHiddenRowsOrCols** method accepts three parameters:

- The name of the document to examine (string).

- The name of the sheet to examine (string).

- Whether to detect rows (true) or columns (false) (Boolean).

---------------------------------------------------------------------------------

## How the Code Works

The code opens the document, by using the [SpreadsheetDocument.Open](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.open.aspx) method and indicating that the document should be open for read-only access (the final **false** parameter value). Next the code retrieves a reference to the workbook part, by using the [WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.workbookpart.aspx) property of the document.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/vb/Program.vb#snippet1)]
***


To find the hidden rows or columns, the code must first retrieve a reference to the specified sheet, given its name. This is not as easy as you might think. The code must look through all the sheet-type descendants of the workbook part's [Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.workbookpart.workbook.aspx) property, examining the [Name](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.name.aspx) property of each sheet that it finds.
Note that this search simply looks through the relations of the workbook, and does not actually find a worksheet part. It simply finds a reference to a [Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx) object, which contains information such as the name and [Id](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.id.aspx) property of the sheet. The simplest way to accomplish this is to use a LINQ query.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/vb/Program.vb#snippet2)]
***

The sheet information you already retrieved provides an **Id** property, and given that **Id** property, the code can retrieve a reference to the corresponding [WorksheetPart](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.worksheet.worksheetpart.aspx) property by calling the [GetPartById](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.openxmlpartcontainer.getpartbyid.aspx) method of the [WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.workbookpart.aspx) object.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/vb/Program.vb#snippet3)]
***


---------------------------------------------------------------------------------

## Retrieving the List of Hidden Row or Column Index Values

The code uses the **detectRows** parameter that you specified when you called the method to determine whether to retrieve information about rows or columns.The code that actually retrieves the list of hidden rows requires only a single line of code.

### [C#](#tab/cs-7)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-7)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/vb/Program.vb#snippet4)]
***

Retrieving the list of hidden columns is a bit trickier, because Excel collapses groups of hidden columns into a single element, and provides [Min](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.column.min.aspx) and [Max](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.column.max.aspx) properties that describe the first and last columns in the group. Therefore, the code that retrieves the list of hidden columns starts the same as the code that retrieves hidden rows. However, it must iterate through the index values (looping each item in the collection of hidden columns, adding each index from the **Min** to the **Max** value, inclusively).

### [C#](#tab/cs-8)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/cs/Program.cs#snippet5)]

### [Visual Basic](#tab/vb-8)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/vb/Program.vb#snippet5)]
***


---------------------------------------------------------------------------------

## Sample Code

The following is the complete **GetHiddenRowsOrCols** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/vb/Program.vb#snippet0)]

---------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
