---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b3890e64-51d1-4643-8d07-2c9d8e060000
title: 'How to: Remove a document part from a package (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Remove a document part from a package (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to remove a document part (file) from a Wordprocessing document
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System
    Imports DocumentFormat.OpenXml.Packaging
```

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

```csharp
    // Open a Wordprocessing document for editing.
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
          // Insert other code here.
    }
```

```vb
    ' Open a Wordprocessing document for editing.
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        ' Insert other code here.
    End Using
```

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
## Basic Structure of a WordProcessingML Document
The basic document structure of a **WordProcessingML** document consists of the **document** and **body**
elements, followed by one or more block level elements such as **p**, which represents a paragraph. A paragraph
contains one or more **r** elements. The **r** stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The **t** element contains a range of text. The **WordprocessingML** markup for the document that the
sample code creates is shown in the following code example.

```xml
    <w:document xmlns:w="https://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Create text in body - CreateWordprocessingDocument</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

Using the Open XML SDK, you can create document structure and
content using strongly-typed classes that correspond to **WordprocessingML** elements. You can find these
classes in the [DocumentFormat.OpenXml.Wordprocessing](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.aspx)
namespace. The following table lists the class names of the classes that
correspond to the **document**, **body**, **p**, **r**, and **t** elements.

WordprocessingML Element|Open XML SDK Class|Description
--|--|--
document|[Document](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.document.aspx) |The root element for the main document part.
body|[Body](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.body.aspx) |The container for the block level structures such as paragraphs, tables, annotations, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification.
| p | [Paragraph](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) | A paragraph. |
| r | [Run](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.run.aspx) | A run. |
| t | [Text](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.text.aspx) | A range of text. |


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

```csharp
    MainDocumentPart mainPart = wordDoc.MainDocumentPart;
    if (mainPart.DocumentSettingsPart != null)
    {
        mainPart.DeletePart(mainPart.DocumentSettingsPart);
    }
```

```vb
    Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
    If mainPart.DocumentSettingsPart IsNot Nothing Then
        mainPart.DeletePart(mainPart.DocumentSettingsPart)
    End If
```

--------------------------------------------------------------------------------
## Sample Code
The following code removes a document part from a package. To run the
program, call the method **RemovePart** like
this example.

```csharp
    string document = @"C:\Users\Public\Documents\MyPkg6.docx";
    RemovePart(document);
```

```vb
    Dim document As String = "C:\Users\Public\Documents\MyPkg6.docx"
    RemovePart(document)
```
> [!NOTE]
> Before running the program on the test file, &quot;MyPkg6.docs,&quot; for example, open the file by using the Open XML SDK Productivity Tool for Microsoft Office and examine its structure. After running the program, examine the file again, and you will notice that the **DocumentSettingsPart** part was removed.

Following is the complete code example in both C\# and Visual Basic.

```csharp
    // To remove a document part from a package.
    public static void RemovePart(string document)
    {
      using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
      {
         MainDocumentPart mainPart = wordDoc.MainDocumentPart;
         if (mainPart.DocumentSettingsPart != null)
         {
            mainPart.DeletePart(mainPart.DocumentSettingsPart);
         }
      }
    }
```

```vb
    ' To remove a document part from a package.
    Public Sub RemovePart(ByVal document As String)
       Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, true)
       Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
       If (Not (mainPart.DocumentSettingsPart) Is Nothing) Then
          mainPart.DeletePart(mainPart.DocumentSettingsPart)
       End If
    End Sub
```

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
