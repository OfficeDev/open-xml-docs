---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: dd28d239-42be-42a9-893e-b65338fe184e
title: 'How to: Parse and read a large spreadsheet document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/10/2025
ms.localizationpriority: high
---
# Parse and read a large spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically read a large Excel file. For more information
about the basic structure of a `SpreadsheetML` document, see [Structure of a SpreadsheetML document](structure-of-a-spreadsheetml-document.md).

[!include[Add-ins note](../includes/addinsnote.md)]

--------------------------------------------------------------------------------
## Approaches to Parsing Open XML Files

The Open XML SDK provides two approaches to parsing Open XML files. You
can use the SDK Document Object Model (DOM), or the Simple API for XML
(SAX) reading and writing features. The SDK DOM is designed to make it
easy to query and parse Open XML files by using strongly-typed classes.
However, the DOM approach requires loading entire Open XML parts into
memory, which can cause an `Out of Memory`
exception when you are working with really large files. Using the SAX
approach, you can employ an OpenXMLReader to read the XML in the file
one element at a time, without having to load the entire file into
memory. Consider using SAX when you need to handle very large files.

The following code segment is used to read a very large Excel file using
the DOM approach.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/parse_and_read_a_large_spreadsheet/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/parse_and_read_a_large_spreadsheet/vb/Program.vb#snippet1)]
***


The following code segment performs an identical task to the preceding
sample (reading a very large Excel file), but uses the SAX approach.
This is the recommended approach for reading very large files.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/parse_and_read_a_large_spreadsheet/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/parse_and_read_a_large_spreadsheet/vb/Program.vb#snippet2)]
***


--------------------------------------------------------------------------------
## Sample Code

You can imagine a scenario where you work for a financial company that
handles very large Excel spreadsheets. Those spreadsheets are updated
daily by analysts and can easily grow to sizes exceeding hundreds of
megabytes. You need a solution to read and extract relevant data from
every spreadsheet. The following code example contains two methods that
correspond to the two approaches, DOM and SAX. The latter technique will
avoid memory exceptions when using very large files. To try them, you
can call them in your code one after the other or you can call each
method separately by commenting the call to the one you would like to
exclude.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/parse_and_read_a_large_spreadsheet/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/parse_and_read_a_large_spreadsheet/vb/Program.vb#snippet3)]
***


The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/parse_and_read_a_large_spreadsheet/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/parse_and_read_a_large_spreadsheet/vb/Program.vb#snippet0)]

--------------------------------------------------------------------------------
## See also


[Structure of a SpreadsheetML document](structure-of-a-spreadsheetml-document.md)



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
