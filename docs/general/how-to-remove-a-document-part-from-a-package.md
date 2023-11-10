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
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Remove a document part from a package

This topic shows how to use the classes in the Open XML SDK for
Office to remove a document part (file) from a Wordprocessing document
programmatically.



--------------------------------------------------------------------------------
## Packages and Document Parts
An Open XML document is stored as a package, whose format is defined by
[ISO/IEC 29500-2](https://www.iso.org/standard/71691.html). The
package can have multiple parts with relationships between them. The
relationship between parts controls the category of the document. A
document can be defined as a word-processing document if its
package-relationship item contains a relationship to a main document
part. If its package-relationship item contains a relationship to a
presentation part it can be defined as a presentation document. If its
package-relationship item contains a relationship to a workbook part, it
is defined as a spreadsheet document. In this how-to topic, you will use
a word-processing document package.


---------------------------------------------------------------------------------
## Getting a WordprocessingDocument Object
The code example starts with opening a package file by passing a file
name as an argument to one of the overloaded [Open()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx) methods of the [DocumentFormat.OpenXml.Packaging.WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx)
that takes a string and a Boolean value that specifies whether the file
should be opened in read/write mode or not. In this case, the Boolean
value is **true** specifying that the file
should be opened in read/write mode.

### [C#](#tab/cs-0)
```csharp
    // Open a Wordprocessing document for editing.
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
          // Insert other code here.
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    ' Open a Wordprocessing document for editing.
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        ' Insert other code here.
    End Using
```
***


The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **wordDoc**. Because the **WordprocessingDocument** class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because
the **Dispose** method is automatically called
when you exit the block; you do not have to explicitly call **Save** and **Close**─as
long as you use **using**.


---------------------------------------------------------------------------------

[!include[Structure](../includes/word/structure.md)]

--------------------------------------------------------------------------------
## Settings Element
The following text from the [ISO/IEC
29500](https://www.iso.org/standard/71691.html) specification
introduces the settings element in a **PresentationML** package.

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
> © ISO/IEC29500: 2008.


--------------------------------------------------------------------------------
## How the Sample Code Works
After you have opened the document, in the **using** statement, as a **WordprocessingDocument** object, you create a
reference to the **DocumentSettingsPart** part.
You can then check if that part exists, if so, delete that part from the
package. In this instance, the **settings.xml**
part is removed from the package.

### [C#](#tab/cs-1)
```csharp
    MainDocumentPart mainPart = wordDoc.MainDocumentPart;
    if (mainPart.DocumentSettingsPart != null)
    {
        mainPart.DeletePart(mainPart.DocumentSettingsPart);
    }
```

### [Visual Basic](#tab/vb-1)
```vb
    Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
    If mainPart.DocumentSettingsPart IsNot Nothing Then
        mainPart.DeletePart(mainPart.DocumentSettingsPart)
    End If
```
***


--------------------------------------------------------------------------------
## Sample Code
The following code removes a document part from a package. To run the
program, call the method **RemovePart** like
this example.

### [C#](#tab/cs-2)
```csharp
    string document = @"C:\Users\Public\Documents\MyPkg6.docx";
    RemovePart(document);
```

### [Visual Basic](#tab/vb-2)
```vb
    Dim document As String = "C:\Users\Public\Documents\MyPkg6.docx"
    RemovePart(document)
```
***

> [!NOTE]
> Before running the program on the test file, &quot;MyPkg6.docs,&quot; for example, open the file by using the Open XML SDK Productivity Tool for Microsoft Office and examine its structure. After running the program, examine the file again, and you will notice that the **DocumentSettingsPart** part was removed.

Following is the complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/remove_a_part_from_a_package/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/remove_a_part_from_a_package/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
