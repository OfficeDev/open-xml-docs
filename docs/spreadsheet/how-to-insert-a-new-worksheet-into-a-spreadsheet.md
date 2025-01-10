---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 944036fa-9251-408f-86cb-2351a5f8cd48
title: 'How to: Insert a new worksheet into a spreadsheet document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/09/2025
ms.localizationpriority: high
---
# Insert a new worksheet into a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to insert a new worksheet into a spreadsheet document
programmatically.

## Getting a SpreadsheetDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument> class represents an
Excel document package. To open and work with an Excel document, you
create an instance of the `SpreadsheetDocument` class from the document.
After you create the instance from the document, you can then obtain
access to the main workbook part that contains the worksheets. The text
in the document is represented in the package as XML using `SpreadsheetML` markup.

To create the class instance from the document that you call one of the
<xref:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open*> methods. Several are provided, each
with a different signature. The sample code in this topic uses the [Open(String, Boolean)](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open?view=openxml-3.0.1#documentformat-openxml-packaging-spreadsheetdocument-open(system-string-system-boolean)) method with a
signature that requires two parameters. The first parameter takes a full
path string that represents the document that you want to open. The
second parameter is either `true` or `false` and represents whether you want the file to
be opened for editing. Any changes that you make to the document will
not be saved if this parameter is `false`.

The code that calls the `Open` method is
shown in the following `using` statement.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/insert_a_new_worksheet/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/insert_a_new_worksheet/vb/Program.vb#snippet1)]
***

--------------------------------------------------------------------------------

[!include[Structure](../includes/spreadsheet/structure.md)]

--------------------------------------------------------------------------------
## Sample Code

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/insert_a_new_worksheet/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/insert_a_new_worksheet/vb/Program.vb#snippet0)]
***

--------------------------------------------------------------------------------
## See also


[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
