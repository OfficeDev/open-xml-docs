---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b3890e64-51d1-4643-8d07-2c9d8e060000
title: 'How to: Remove a document part from a package'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/03/2025
ms.localizationpriority: medium
---
# Remove a document part from a package

This topic shows how to use the classes in the Open XML SDK for
Office to remove a document part (file) from a Wordprocessing document
programmatically.



--------------------------------------------------------------------------------
[!include[Structure](../includes/word/packages-and-document-parts.md)]


---------------------------------------------------------------------------------
## Getting a WordprocessingDocument Object

The code example starts with opening a package file by passing a file
name as an argument to one of the overloaded <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open*> methods of the 
<xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument>
that takes a string and a Boolean value that specifies whether the file
should be opened in read/write mode or not. In this case, the Boolean
value is `true` specifying that the file
should be opened in read/write mode.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/remove_a_part_from_a_package/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/remove_a_part_from_a_package/vb/Program.vb#snippet1)]
***


[!include[Using Statement](../includes/word/using-statement.md)]


---------------------------------------------------------------------------------

[!include[Structure](../includes/word/structure.md)]

--------------------------------------------------------------------------------
## Settings Element
The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces the settings element in a `PresentationML` package.

> This element specifies the settings that are applied to a
> WordprocessingML document. This element is the root element of the
> Document Settings part in a WordprocessingML document.   
> **Example**:
> Consider the following WordprocessingML fragment for the settings part
> of a document:

```xml
    <w:settings>
      <w:defaultTabStop w:val="720" />
      <w:characterSpacingControl w:val="dontCompress" />
    </w:settings>
```

> The **settings** element contains all of the
> settings for this document. In this case, the two settings applied are
> automatic tab stop increments of 0.5" using the **defaultTabStop** element, and no character level
> white space compression using the **characterSpacingControl** element. 
> 
> &copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]


--------------------------------------------------------------------------------
## How the Sample Code Works

After you have opened the document, in the `using` statement, as a <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> object, you create a
reference to the `DocumentSettingsPart` part.
You can then check if that part exists, if so, delete that part from the
package. In this instance, the `settings.xml`
part is removed from the package.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/remove_a_part_from_a_package/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/remove_a_part_from_a_package/vb/Program.vb#snippet2)]
***


--------------------------------------------------------------------------------
## Sample Code

Following is the complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/remove_a_part_from_a_package/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/remove_a_part_from_a_package/vb/Program.vb#snippet0)]
***

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
