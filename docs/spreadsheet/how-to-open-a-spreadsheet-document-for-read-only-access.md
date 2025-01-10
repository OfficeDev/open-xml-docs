---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 625bf571-5630-47f8-953f-e9e1a93e3229
title: 'How to: Open a spreadsheet document for read-only access'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/10/2025
ms.localizationpriority: high
---
# Open a spreadsheet document for read-only access

This topic shows how to use the classes in the Open XML SDK for
Office to open a spreadsheet document for read-only access
programmatically.



---------------------------------------------------------------------------------
## When to Open a Document for Read-Only Access

Sometimes you want to open a document to inspect or retrieve some
information, and you want to do this in a way that ensures the document
remains unchanged. In these instances, you want to open the document for
read-only access. This How To topic discusses several ways to
programmatically open a read-only spreadsheet document.

--------------------------------------------------------------------------------
[!include[Spreadsheet Object](../includes/spreadsheet/spreadsheet-object.md)]

--------------------------------------------------------------------------------
## Getting a SpreadsheetDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument> class represents an
Excel document package. To create an Excel document, you create an
instance of the `SpreadsheetDocument` class
and populate it with parts. At a minimum, the document must have a
workbook part that serves as a container for the document, and at least
one worksheet part. The text is represented in the package as XML using
SpreadsheetML markup.

To create the class instance from the document that you call one of the
<xref:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open*> overload methods. Several `Open` methods are provided, each with a different
signature. The methods that let you specify whether a document is
editable are listed in the following table.

|Open|Class Library Reference Topic|Description|
--|--|--
Open(String, Boolean)|[Open(String, Boolean)](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open?view=openxml-3.0.1#documentformat-openxml-packaging-spreadsheetdocument-open(system-string-system-boolean))|Create an instance of the SpreadsheetDocument class from the specified file.
Open(Stream, Boolean)|[Open(Stream, Boolean](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open?view=openxml-3.0.1#documentformat-openxml-packaging-spreadsheetdocument-open(system-io-stream-system-boolean))|Create an instance of the SpreadsheetDocument class from the specified IO stream.
Open(String, Boolean, OpenSettings)|[Open(String, Boolean, OpenSettings)](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open?view=openxml-3.0.1#documentformat-openxml-packaging-spreadsheetdocument-open(system-string-system-boolean-documentformat-openxml-packaging-opensettings))|Create an instance of the SpreadsheetDocument class from the specified file.
Open(Stream, Boolean, OpenSettings)|[Open(Stream, Boolean, OpenSettings)](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open?view=openxml-3.0.1#documentformat-openxml-packaging-spreadsheetdocument-open(system-io-stream-system-boolean-documentformat-openxml-packaging-opensettings))|Create an instance of the SpreadsheetDocument class from the specified I/O stream.

The table earlier in this topic lists only those `Open` methods that accept a Boolean value as the
second parameter to specify whether a document is editable. To open a
document for read-only access, specify `False` for this parameter.

Notice that two of the `Open` methods create
an instance of the SpreadsheetDocument class based on a string as the
first parameter. The first example in the sample code uses this
technique. It uses the first `Open` method in
the table earlier in this topic; with a signature that requires two
parameters. The first parameter takes a string that represents the full
path file name from which you want to open the document. The second
parameter is either `true` or `false`. This example uses `false` and indicates that you want to open the
file as read-only.

The following code example calls the `Open`
Method.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/spreadsheet/open_for_read_only_access/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/spreadsheet/open_for_read_only_access/vb/Program.vb#snippet1)]
***


The other two `Open` methods create an
instance of the SpreadsheetDocument class based on an input/output
stream. You might use this approach, for example, if you have a
Microsoft SharePoint Foundation 2010 application that uses stream
input/output, and you want to use the Open XML SDK to work with a
document.

The following code example opens a document based on a stream.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/open_for_read_only_access/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/open_for_read_only_access/vb/Program.vb#snippet2)]
***


Suppose you have an application that uses the Open XML support in the
System.IO.Packaging namespace of the .NET Framework Class Library, and
you want to use the Open XML SDK to work with a package as
read-only. Whereas the Open XML SDK includes method overloads that
accept a `Package` as the first parameter,
there is not one that takes a Boolean as the second parameter to
indicate whether the document should be opened for editing.

The recommended method is to open the package as read-only at first,
before creating the instance of the `SpreadsheetDocument` class, as shown in the second
example in the sample code. The following code example performs this
operation.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/open_for_read_only_access/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/open_for_read_only_access/vb/Program.vb#snippet3)]

---------------------------------------------------------------------------------

## Sample Code

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/open_for_read_only_access/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/open_for_read_only_access/vb/Program.vb#snippet0)]
***

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
