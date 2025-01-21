---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: c811c2c7-1066-45a5-a724-33d0fbfd5284
title: 'How to: Open a word processing document for read-only access'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/16/2025
ms.localizationpriority: high
---
# Open a word processing document for read-only access

This topic describes how to use the classes in the Open XML SDK for
Office to programmatically open a word processing document for read only
access.



---------------------------------------------------------------------------------
## When to Open a Document for Read-only Access

Sometimes you want to open a document to inspect or retrieve some
information, and you want to do so in a way that ensures the document
remains unchanged. In these instances, you want to open the document for
read-only access. This how-to topic discusses several ways to
programmatically open a read-only word processing document.


--------------------------------------------------------------------------------
## Create a WordprocessingDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class represents a
Word document package. To work with a Word document, first create an
instance of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument>
class from the document, and then work with that instance. Once you
create the instance from the document, you can then obtain access to the
main document part that contains the text of the document. Every Open
XML package contains some number of parts. At a minimum, a <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> must contain a main
document part that acts as a container for the main text of the
document. The package can also contain additional parts. Notice that in
a Word document, the text in the main document part is represented in
the package as XML using WordprocessingML markup.

To create the class instance from the document you call one of the `Open` methods. Several `Open` methods are provided, each with a different
signature. The methods that let you specify whether a document is
editable are listed in the following table.

Open Method|Class Library Reference Topic|Description
--|--|--
`Open(String, Boolean)`|<xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)> |Create an instance of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class from the specified file.
`Open(Stream, Boolean)`|<xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.IO.Stream,System.Boolean)> |Create an instance of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class from the specified IO stream.
`Open(String, Boolean, OpenSettings)`|<xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean,DocumentFormat.OpenXml.Packaging.OpenSettings)> |Create an instance of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class from the specified file.
`Open(Stream, Boolean, OpenSettings)`|<xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.IO.Stream,System.Boolean,DocumentFormat.OpenXml.Packaging.OpenSettings)> |Create an instance of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class from the specified I/O stream.

The table above lists only those `Open`
methods that accept a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read only access,
you specify false for this parameter.

Notice that two of the `Open` methods create
an instance of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument>
class based on a string as the first parameter. The first example in the
sample code uses this technique. It uses the first `Open` method in the table above; with a signature
that requires two parameters. The first parameter takes a string that
represents the full path filename from which you want to open the
document. The second parameter is either `true` or `false`; this
example uses `false` and indicates whether
you want to open the file for editing.

The following code example calls the `Open`
Method.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/open_for_read_only_access/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/open_for_read_only_access/vb/Program.vb#snippet1)]
***


The other two `Open` methods create an
instance of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument>
class based on an input/output stream. You might employ this approach,
for instance, if you have a Microsoft SharePoint Online
application that uses stream input/output, and you want to use the Open
XML SDK to work with a document.

The following code example opens a document based on a stream.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/open_for_read_only_access/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/open_for_read_only_access/vb/Program.vb#snippet3)]
***


Suppose you have an application that employs the Open XML support in the
System.IO.Packaging namespace of the .NET Framework Class Library, and
you want to use the Open XML SDK to work with a package read only.
While the Open XML SDK includes method overloads that accept a <xref:System.IO.Packaging.Package>
as the first parameter, there is not one that takes a Boolean as
the second parameter to indicate whether the document should be opened for editing.

The recommended method is to open the package as read-only to begin with
prior to creating the instance of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class, as shown in the
second example in the sample code. The following code example performs
this operation.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/open_for_read_only_access/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/open_for_read_only_access/vb/Program.vb#snippet2)]
***


Once you open the Word document package, you can access the main
document part. To access the body of the main document part, you assign
a reference to the existing document body, as shown in the following
code example.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/open_for_read_only_access/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/open_for_read_only_access/vb/Program.vb#snippet4)]
***


---------------------------------------------------------------------------------

[!include[Structure](../includes/word/structure.md)]

--------------------------------------------------------------------------------
## Generate the WordprocessingML Markup to Add Text and Attempt to Save
The sample code shows how you can add some text and attempt to save the
changes to show that access is read-only. Once you have access to the
body of the main document part, you add text by adding instances of the
<xref:DocumentFormat.OpenXml.Wordprocessing.Paragraph>,
<xref:DocumentFormat.OpenXml.Wordprocessing.Run>, and <xref:DocumentFormat.OpenXml.Wordprocessing.Text>
classes. This generates the required WordprocessingML markup. The
following code example adds the paragraph, run, and text.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/open_for_read_only_access/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/open_for_read_only_access/vb/Program.vb#snippet5)]
***


--------------------------------------------------------------------------------
## Sample Code
The first example method shown here, `OpenWordprocessingDocumentReadOnly`, opens a Word
document for read-only access. Call it by passing a full path to the
file that you want to open. For example, the following code example
opens the file path from the first command line argument for read-only access.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/word/open_for_read_only_access/cs/Program.cs#snippet6)]
### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/word/open_for_read_only_access/vb/Program.vb#snippet6)]
***


The second example method, `OpenWordprocessingPackageReadonly`, shows how to
open a Word document for read-only access from a <xref:System.IO.Packaging.Package>.
Call it by passing a full path to the file
that you want to open. For example, the following code example
opens the file path from the first command line argument for read-only access.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/word/open_for_read_only_access/cs/Program.cs#snippet7)]
### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/word/open_for_read_only_access/vb/Program.vb#snippet7)]
***

The third example method, `OpenWordprocessingStreamReadonly`, shows how to
open a Word document for read-only access from a a stream.
Call it by passing a full path to the file
that you want to open. For example, the following code example
opens the file path from the first command line argument for read-only access.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/word/open_for_read_only_access/cs/Program.cs#snippet8)]
### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/word/open_for_read_only_access/vb/Program.vb#snippet8)]
***

> [!IMPORTANT]
> If you uncomment the statement that saves the file, the program would throw an **IOException** because the file is opened for read-only access.

The following is the complete sample code in C\# and VB.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/open_for_read_only_access/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/open_for_read_only_access/vb/Program.vb#snippet0)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
