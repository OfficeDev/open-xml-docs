---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 2f6f0f89-0ac0-4d40-9f1a-222caf074cf1
title: 'How to: Replace text in a word document using SAX (Simple API for XML)'
description: 'Learn how to replace text in a Word document using SAX (Simple API for XML)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 04/03/2025
ms.localizationpriority: high
---
# Replace Text in a Word Document using SAX (Simple API for XML)

This topic shows how to use the Open XML SDK to search and replace text in a Word document with the
Open XML SDK using the Simple API for XML (SAX) approach. For more information about the basic structure
of a `WordprocessingML` document, see [Structure of a WordprocessingML document](./structure-of-a-wordprocessingml-document.md).

## Why Use the SAX Approach?

The Open XML SDK provides two ways to parse Office Open XML files: the Document Object Model (DOM) and the Simple API for XML (SAX). The DOM approach is designed to make it easy to query and parse Open XML files by using strongly-typed classes. However, the DOM approach requires loading entire Open XML parts into memory, which can lead to slower processing and Out of Memory exceptions when working with very large parts. The SAX approach reads in the XML in an Open XML part one element at a time without reading in the entire part into memory giving noncached, forward-only access to the XML data, which makes it a better choice when reading very large parts.

## Accessing the MainDocumentPart

The text of a Word document is stored in the <xref:DocumentFormat.OpenXml.Packaging.MainDocumentPart>, so the first step to
finding and replacing text is to access the Word document's `MainDocumentPart`. To do that we first use the `WordprocessingDocument.Open`
method passing in the path to the document as the first parameter and a second parameter `true` to indicate that we
are opening the file for editing. Then make sure that the `MainDocumentPart` is not null.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/replace_text_with_sax/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/replace_text_with_sax/vb/Program.vb#snippet1)]
***

## Create Memory Stream, OpenXmlReader, and OpenXmlWriter

With the DOM approach to editing documents, the entire part is read into memory, so we can use the Open XML SDK's
strongly typed classes to access the <xref:DocumentFormat.OpenXml.Wordprocessing.Text> class to access the
document's text and edit it. The SAX approach, however, uses the <xref:DocumentFormat.OpenXml.OpenXmlPartReader>
and <xref:DocumentFormat.OpenXml.OpenXmlPartWriter> classes, which access a part's stream with forward-only
access. The advantage of this is that the entire part does not need to be loaded into memory, which is faster
and uses less memory, but since the same part cannot be opened in multiple streams at the same time, we cannot create a
<xref:DocumentFormat.OpenXml.OpenXmlReader> to read a part and a <xref:DocumentFormat.OpenXml.OpenXmlWriter> to edit
the same part at the same time. The solution to this is to create an additional memory stream and write the
updated part to the new memory stream then use the stream to update the part when `OpenXmlReader` and `OpenXmlWriter`
have been disposed. In the code below we create the `MemoryStream` to store the updated part and create an
`OpenXmlReader` for the `MainDocumentPart` and a `OpenXmlWriter` to write to the `MemoryStream`

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/replace_text_with_sax/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/replace_text_with_sax/vb/Program.vb#snippet2)]
***

## Reading the Part and Writing to the New Stream

Now that we have an `OpenXmlReader` to read the part and an `OpenXmlWriter` to write to the new `MemoryStream`
we use the <xref:DocumentFormat.OpenXml.OpenXmlReader.Read*> method to read each element in the part. As
each element is read in we check if it is of type `Text` and if it is, we use the <xrefDocumentFormat.OpenXml.OpenXmlReader.GetText*>
method to access the text and use <xref:System.String.Replace*> to update the text. If it is not a
`Text` element, then we write it to the stream unchanged.

> [!Note]
> In a Word document text can be separated into multiple `Text` elements, so if you are replacing a
> phrase and not a single word, it's best to replace one word at a time.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/replace_text_with_sax/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/replace_text_with_sax/vb/Program.vb#snippet3)]
***

## Writing the New Stream to the MainDocumentPart

With the updated part written to the memory stream the last step is to set the `MemoryStream`'s
position to 0 and use the <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPart.FeedData*> method
to replace the `MainDocumentPart` with the updated stream.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/replace_text_with_sax/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/replace_text_with_sax/vb/Program.vb#snippet4)]
***

## Sample Code

Below is the complete sample code to replace text in a Word document using the SAX (Simple API for XML)
approach.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/replace_text_with_sax/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/replace_text_with_sax/vb/Program.vb#snippet0)]
***