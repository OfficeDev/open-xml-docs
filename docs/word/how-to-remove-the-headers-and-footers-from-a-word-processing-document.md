---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 22f973f4-58d1-4dd4-943e-a15ac2571b7c
title: 'How to: Remove the headers and footers from a word processing document'
description: 'Learn how to remove the headers and footers from a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/17/2025
ms.localizationpriority: medium
---
# Remove the headers and footers from a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically remove all headers and footers in a word
processing document. It contains an example `RemoveHeadersAndFooters` method to illustrate this
task.



## RemoveHeadersAndFooters Method

You can use the `RemoveHeadersAndFooters`
method to remove all header and footer information from a word
processing document. Be aware that you must not only delete the header
and footer parts from the document storage, you must also delete the
references to those parts from the document too. The sample code
demonstrates both steps in the operation. The `RemoveHeadersAndFooters` method accepts a single
parameter, a string that indicates the path of the file that you want to
modify.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/remove_the_headers_and_footers/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/remove_the_headers_and_footers/vb/Program.vb#snippet1)]
***


The complete code listing for the method can be found in the [Sample Code](#sample-code) section.

## Calling the Sample Method

To call the sample method, pass a string for the first parameter that
contains the file name of the document that you want to modify as shown
in the following code example.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/remove_the_headers_and_footers/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/remove_the_headers_and_footers/vb/Program.vb#snippet2)]
***


## How the Code Works

The `RemoveHeadersAndFooters` method works
with the document you specify, deleting all of the header and footer
parts and references to those parts. The code starts by opening the
document, using the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A> method and indicating that the
document should be opened for read/write access (the final true
parameter). Given the open document, the code uses the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart> property to navigate to
the main document, storing the reference in a variable named `docPart`.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/remove_the_headers_and_footers/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/remove_the_headers_and_footers/vb/Program.vb#snippet3)]
***


## Confirm Header/Footer Existence

Given a reference to the document part, the code next determines if it
has any work to do, i.e. if the document contains any headers or
footers. To decide, the code calls the `Count` method of both the <xref:DocumentFormat.OpenXml.Packaging.MainDocumentPart.HeaderParts> and
<xref:DocumentFormat.OpenXml.Packaging.MainDocumentPart.FooterParts> properties of the document
part, and if either returns a value greater than 0, the code continues.
Be aware that the `HeaderParts` and `FooterParts` properties each return an
<xref:System.Collections.Generic.IEnumerable%601> of
<xref:DocumentFormat.OpenXml.Packaging.HeaderPart> or <xref:DocumentFormat.OpenXml.Packaging.FooterPart> objects, respectively.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/remove_the_headers_and_footers/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/remove_the_headers_and_footers/vb/Program.vb#snippet4)]
***


## Remove the Header and Footer Parts

Given a collection of references to header and footer parts, you could
write code to delete each one individually, but that is not necessary
because of the Open XML SDK. Instead, you can call the <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.DeleteParts%2A> method, passing in the
collection of parts to be deleted─this simple method provides a shortcut
for deleting a collection of parts. Therefore, the following few lines
of code take the place of the loop that you would otherwise have to
write yourself.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/remove_the_headers_and_footers/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/remove_the_headers_and_footers/vb/Program.vb#snippet5)]
***


## Delete the Header and Footer References

At this point, the code has deleted the header and footer parts, but the
document still contains orphaned references to those parts. Before the
orphaned references can be removed, the code must retrieve a reference
to the content of the document (that is, to the XML content contained
within the main document part).

To remove the stranded references, the code first retrieves a collection
of HeaderReference elements, converts the collection to a `List`, and then
loops through the collection, calling the <xref:DocumentFormat.OpenXml.OpenXmlElement.Remove> method for each element found. Note
that the code converts the `IEnumerable`
returned by the <xref:DocumentFormat.OpenXml.OpenXmlElement.Descendants> method into a `List` so that it
can delete items from the list, and that the <xref:DocumentFormat.OpenXml.Wordprocessing.HeaderReference> type that is provided by
the Open XML SDK makes it easy to refer to elements of type `HeaderReference` in the XML content. (Without that
additional help, you would have to work with the details of the XML
content directly.) Once it has removed all the headers, the code repeats
the operation with the footer elements.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/word/remove_the_headers_and_footers/cs/Program.cs#snippet6)]
### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/word/remove_the_headers_and_footers/vb/Program.vb#snippet6)]
***


## Sample Code

The following is the complete `RemoveHeadersAndFooters` code sample in C\# and
Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/remove_the_headers_and_footers/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/remove_the_headers_and_footers/vb/Program.vb#snippet0)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
