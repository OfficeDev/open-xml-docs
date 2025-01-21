---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: a6d35b76-d12a-460c-9d9d-2334abde759e
title: 'How to: Retrieve a list of the hidden worksheets in a spreadsheet document'
description: 'Learn how to retrieve a list of the hidden worksheets in a spreadsheet document using the Open XML SDK.'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/10/2025
ms.localizationpriority: medium
---

# Retrieve a list of the hidden worksheets in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically retrieve a list of hidden worksheets in a Microsoft Excel workbook, without loading the document into Excel. It contains an example `GetHiddenSheets` method to illustrate this task.

## GetHiddenSheets method

You can use the `GetHiddenSheets` method, to retrieve a list of the hidden worksheets in a workbook. The `GetHiddenSheets` method accepts a single parameter, a string that indicates the path of the file that you want to examine. The method works with the workbook you specify, filling a <xref:System.Collections.Generic.List`1> instance with a reference to each hidden `Sheet` object.

## Retrieve the collection of worksheets

The `WorkbookPart` class provides a `Workbook` property, which in turn contains the XML content of the workbook. Although the Open XML SDK provides the `Sheets` property, which returns a collection of the `Sheet` parts, all the information that you need is provided by the `Sheet` elements within the `Workbook` XML content.
The following code uses the `Descendants` generic method of the `Workbook` object to retrieve a collection of `Sheet` objects that contain information about all the sheet child elements of the workbook's XML content.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_worksheets/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_worksheets/vb/Program.vb#snippet1)]
***

## Retrieve hidden sheets

It's important to be aware that Excel supports two levels of worksheets. You can hide a worksheet by using the Excel user interface by right-clicking the worksheets tab and opting to hide the worksheet.
For these worksheets, the `State` property of the `Sheet` object contains an enumerated value of `Hidden`. You can also make a worksheet very hidden by writing code (either in VBA or in another language) that sets the sheet's `Visible` property to the enumerated value `xlSheetVeryHidden`. For worksheets hidden in this manner, the `State` property of the `Sheet` object contains the enumerated value `VeryHidden`.

Given the collection that contains information about all the sheets, the following code uses the <xref:System.Linq.Enumerable.Where*> function to filter the collection so that it contains only the sheets in which the `State` property is not null. If the `State` property is not null, the code looks for the `Sheet` objects in which the `State` property as a value, and where the value is either `SheetStateValues.Hidden` or `SheetStateValues.VeryHidden`.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_worksheets/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_worksheets/vb/Program.vb#snippet2)]
***

## Sample code

The following is the complete `GetHiddenSheets` code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_worksheets/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_hidden_worksheets/vb/Program.vb#snippet0)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
