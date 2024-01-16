---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 0aa2aef3-b329-4ccc-8f25-9660c083e14e
title: 'How to: Retrieve a dictionary of all named ranges in a spreadsheet document'
description: 'Learn how to retrieve a dictionary of all named ranges in a spreadsheet document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 12/05/2023
ms.localizationpriority: medium
---
# Retrieve a dictionary of all named ranges in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve a dictionary that contains the names
and ranges of all defined names in a Microsoft Excel workbook. It contains an example **GetDefinedNames** method
to illustrate this task.

## GetDefinedNames Method

The **GetDefinedNames** method accepts a
single parameter that indicates the name of the document from which to
retrieve the defined names. The method returns an
[Dictionary](https://msdn.microsoft.com/library/xfhwa508.aspx)
instance that contains information about the defined names within the
specified workbook, which may be empty if there are no defined names.

## How the Code Works

The code opens the spreadsheet document, using the **Open** method, indicating that the
document should be open for read-only access with the final false parameter. Given the open workbook, the code uses the **WorkbookPart** property to navigate to the main workbook part. The code stores this reference in a variable named **wbPart**.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_dictionary_of_all_named_ranges/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/retrieve_a_dictionary_of_all_named_ranges/vb/Program.vb#snippet1)]
***


## Retrieving the Defined Names

Given the workbook part, the next step is simple. The code uses the
**Workbook** property of the workbook part to retrieve a reference to the content of the workbook, and then retrieves the **DefinedNames** collection provided by the Open XML SDK. This property returns a collection of all of the
defined names that are contained within the workbook. If the property returns a non-null value, the code then iterates through the collection, retrieving information about each named part and adding the key  name) and value (range description) to the dictionary for each defined name.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_dictionary_of_all_named_ranges/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/retrieve_a_dictionary_of_all_named_ranges/vb/Program.vb#snippet2)]
***


## Sample Code

The following is the complete **GetDefinedNames** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_dictionary_of_all_named_ranges/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/retrieve_a_dictionary_of_all_named_ranges/vb/Program.vb#snippet0)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
