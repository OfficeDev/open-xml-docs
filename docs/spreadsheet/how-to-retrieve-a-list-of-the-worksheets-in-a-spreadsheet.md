---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: a0c1e144-2080-4470-bd4b-ed98f1399374
title: 'How to: Retrieve a list of the worksheets in a spreadsheet document'
description: 'Learn how to retrieve a list of the worksheets in a spreadsheet document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/27/2023
ms.localizationpriority: high
---
# Retrieve a list of the worksheets in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve a list of the worksheets in a
Microsoft Excel workbook, without loading
the document into Excel. It contains an example **GetAllWorksheets** method to illustrate this task.



--------------------------------------------------------------------------------

## GetAllWorksheets Method

You can use the **GetAllWorksheets** method,
which is shown in the following code, to retrieve a list of the
worksheets in a workbook. The **GetAllWorksheets** method accepts a single
parameter, a string that indicates the path of the file that you want to
examine.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/vb/Program.vb#snippet2)]
***


The method works with the workbook you specify, returning an instance of
the **[Sheets](/dotnet/api/documentformat.openxml.spreadsheet.sheets)** object, from which you can retrieve
a reference to each **[Sheet](/dotnet/api/documentformat.openxml.spreadsheet.sheet)** object.

--------------------------------------------------------------------------------

## Calling the GetAllWorksheets Method

To call the **GetAllWorksheets** method, pass
the required value, as shown in the following code.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/vb/Program.vb#snippet2)]
***


--------------------------------------------------------------------------------

## How the Code Works

The sample method, **GetAllWorksheets**,
creates a variable that will contain a reference to the **Sheets** collection of the workbook. At the end of
its work, the method returns the variable, which contains either a
reference to the **Sheets** collection, or
null/Nothing if there were no sheets (this cannot occur in a well-formed
workbook).

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/vb/Program.vb#snippet3)]
***


The code then continues by opening the document in read-only mode, and
retrieving a reference to the **[WorkbookPart](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.workbookpart)**.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/vb/Program.vb#snippet4)]
***


To get access to the **[Workbook](/dotnet/api/documentformat.openxml.spreadsheet.workbook)** object, the code retrieves the value of the **[Workbook](/dotnet/api/documentformat.openxml.packaging.workbookpart.workbook)** property from the **WorkbookPart**, and then retrieves a reference to the **Sheets** object from the **[Sheets](/dotnet/api/documentformat.openxml.spreadsheet.workbook.sheets)** property of the **Workbook**. The **Sheets** object contains the collection of **[Sheet](/dotnet/api/documentformat.openxml.spreadsheet.sheet)** objects that provide the method's return value.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/cs/Program.cs#snippet5)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/vb/Program.vb#snippet5)]
***


--------------------------------------------------------------------------------

## Sample Code

The following is the complete **GetAllWorksheets** code sample in C\# and Visual
Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/vb/Program.vb#snippet0)]

--------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
