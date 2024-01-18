---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b0d3d890-431a-4838-89dc-1f0dccd5dcd0
title: 'How to: Get the contents of a document part from a package'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Get the contents of a document part from a package

This topic shows how to use the classes in the Open XML SDK for
Office to retrieve the contents of a document part in a Wordprocessing
document programmatically.



--------------------------------------------------------------------------------
[!include[Structure](../includes/word/packages-and-document-parts.md)]


---------------------------------------------------------------------------------
## Getting a WordprocessingDocument Object
The code starts with opening a package file by passing a file name to
one of the overloaded [Open()](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument.open) methods (Visual Basic .NET Shared
method or C\# static method) of the [WordprocessingDocument](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument) class that takes a
string and a Boolean value that specifies whether the file should be
opened in read/write mode or not. In this case, the Boolean value is
**false** specifying that the file should be
opened in read-only mode to avoid accidental changes.

### [C#](#tab/cs-0)
```csharp
    // Open a Wordprocessing document for editing.
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
    {
          // Insert other code here.
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    ' Open a Wordprocessing document for editing.
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, False)
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
long as you use using.


---------------------------------------------------------------------------------

[!include[Structure](../includes/word/structure.md)]

--------------------------------------------------------------------------------
## Comments Element
In this how-to, you are going to work with comments. Therefore, it is
useful to familiarize yourself with the structure of the \<**comments**\> element. The following information
from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification can be useful when working with this element.

This element specifies all of the comments defined in the current
document. It is the root element of the comments part of a
WordprocessingML document.Consider the following WordprocessingML
fragment for the content of a comments part in a WordprocessingML
document:

```xml
    <w:comments>
      <w:comment … >
        …
      </w:comment>
    </w:comments>
```

The **comments** element contains the single
comment specified by this document in this example.

© ISO/IEC29500: 2008.

The following XML schema fragment defines the contents of this element.

```xml
    <complexType name="CT_Comments">
       <sequence>
           <element name="comment" type="CT_Comment" minOccurs="0" maxOccurs="unbounded"/>
       </sequence>
    </complexType>
```

--------------------------------------------------------------------------------
## How the Sample Code Works
After you have opened the source file for reading, you create a **mainPart** object by instantiating the **MainDocumentPart**. Then you can create a reference
to the **WordprocessingCommentsPart** part of
the document.

### [C#](#tab/cs-1)
```csharp
    // To get the contents of a document part.
    public static string GetCommentsFromDocument(string document)
    {
        string comments = null;

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            WordprocessingCommentsPart WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart;
```

### [Visual Basic](#tab/vb-1)
```vb
    ' To get the contents of a document part.
    Public Shared Function GetCommentsFromDocument(ByVal document As String) As String
        Dim comments As String = Nothing

        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
            Dim WordprocessingCommentsPart As WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart
```
***


You can then use a **StreamReader** object to
read the contents of the **WordprocessingCommentsPart** part of the document
and return its contents.

### [C#](#tab/cs-2)
```csharp
    using (StreamReader streamReader = new StreamReader(WordprocessingCommentsPart.GetStream()))
            {
                comments = streamReader.ReadToEnd();
            }
        }
        return comments;
```

### [Visual Basic](#tab/vb-2)
```vb
    Using streamReader As New StreamReader(WordprocessingCommentsPart.GetStream())
    comments = streamReader.ReadToEnd()
    End Using
    Return comments
```
***


--------------------------------------------------------------------------------
## Sample Code
The following code retrieves the contents of a **WordprocessingCommentsPart** part contained in a
**WordProcessing** document package. You can
run the program by calling the **GetCommentsFromDocument** method as shown in the
following example.

### [C#](#tab/cs-3)
```csharp
    string document = @"C:\Users\Public\Documents\MyPkg5.docx";
    GetCommentsFromDocument(document);
```

### [Visual Basic](#tab/vb-3)
```vb
    Dim document As String = "C:\Users\Public\Documents\MyPkg5.docx"
    GetCommentsFromDocument(document)
```
***


Following is the complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/get_the_contents_of_a_part_from_a_package/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/get_the_contents_of_a_part_from_a_package/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also


[Open XML SDK class library
reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
