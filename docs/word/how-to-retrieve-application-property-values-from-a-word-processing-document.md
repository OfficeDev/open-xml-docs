---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 3e9ca812-460e-442e-8257-38f523a53dc6
title: 'How to: Retrieve application property values from a word processing document'
description: 'Learn how to retrieve application property values from a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/31/2024
ms.localizationpriority: medium
---

# Retrieve application property values from a word processing document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically retrieve an application property from a Microsoft Word document, without loading the document into Word. It contains example code to illustrate this task.



## Retrieving Application Properties

To retrieve application document properties, you can retrieve the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.ExtendedFilePropertiesPart> property of a <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> object, and then retrieve the specific application property you need. To do this, you must first get a reference to the document, as shown in the following code.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/retrieve_application_property_values/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/retrieve_application_property_values/vb/Program.vb#snippet1)]
***


Given the reference to the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> object, you can retrieve a reference to the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.ExtendedFilePropertiesPart> property of the document. This object provides its own properties, each of which exposes one of the application document properties.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/retrieve_application_property_values/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/retrieve_application_property_values/vb/Program.vb#snippet2)]
***


Once you have the reference to the properties of <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.ExtendedFilePropertiesPart>, you can then retrieve any of the application properties, using simple code such as that shown
in the next example. Note that the code must confirm that the reference to each property isn't `null` of `Nothing` before retrieving its `Text` property. Unlike core properties, document properties aren't available if you (or the application) haven't specifically given them a value.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/retrieve_application_property_values/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/retrieve_application_property_values/vb/Program.vb#snippet3)]
***


## Sample Code

The following is the complete code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/retrieve_application_property_values/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/retrieve_application_property_values/vb/Program.vb#snippet0)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
