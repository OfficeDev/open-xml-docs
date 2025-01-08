---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 124cb0a0-cc47-433f-bad0-06b793890650
title: 'How to: Get worksheet information from an Open XML package'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/03/2024
ms.localizationpriority: high
---

# Get worksheet information from an Open XML package

This topic shows how to use the classes in the Open XML SDK for Office to programmatically retrieve information from a worksheet in a Spreadsheet document.

[!include[Structure](../includes/spreadsheet/structure.md)]

## How the Sample Code Works

After you have opened the file for read-only access, you instantiate the **Sheets** class.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/get_worksheetformation_from_a_package/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/get_worksheetformation_from_a_package/vb/Program.vb#snippet1)]
***


You then you iterate through the **Sheets** collection and display <xref:DocumentFormat.OpenXml.OpenXmlElement> and the
<xref:DocumentFormat.OpenXml.OpenXmlAttribute> in each element.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/get_worksheetformation_from_a_package/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/get_worksheetformation_from_a_package/vb/Program.vb#snippet2)]
***


By displaying the attribute information you get the name and ID for each worksheet in the spreadsheet file.

## Sample code

The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/get_worksheetformation_from_a_package/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/get_worksheetformation_from_a_package/vb/Program.vb#snippet0)]

## See also

[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
