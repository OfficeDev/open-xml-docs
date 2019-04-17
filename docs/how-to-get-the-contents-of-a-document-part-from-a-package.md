---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b0d3d890-431a-4838-89dc-1f0dccd5dcd0
title: 'How to: Get the contents of a document part from a package (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
localization_priority: Priority
---
# Get the contents of a document part from a package (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to retrieve the contents of a document part in a Wordprocessing
document programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System
    Imports System.IO
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
The code starts with opening a package file by passing a file name to
one of the overloaded [Open()](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx) methods (Visual Basic .NET Shared
method or C\# static method) of the [WordprocessingDocument](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class that takes a
string and a Boolean value that specifies whether the file should be
opened in read/write mode or not. In this case, the Boolean value is
**false** specifying that the file should be
opened in read-only mode to avoid accidental changes.

```csharp
    // Open a Wordprocessing document for editing.
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
    {
          // Insert other code here.
    }
```

```vb
    ' Open a Wordprocessing document for editing.
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, False)
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
long as you use using.


---------------------------------------------------------------------------------
## Basic Structure of a WordProcessingML Document
The basic document structure of a **WordProcessingML** document consists of the **document** and **body**
elements, followed by one or more block level elements such as **p**, which represents a paragraph. A paragraph
contains one or more **r** elements. The **r** stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The **t** element contains a range of text. The **WordprocessingML** markup for the document that the
sample code creates is shown in the following code example.

```xml
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Create text in body - CreateWordprocessingDocument</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

Using the Open XML SDK 2.5, you can create document structure and
content using strongly-typed classes that correspond to **WordprocessingML** elements. You can find these
classes in the [DocumentFormat.OpenXml.Wordprocessing](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.aspx)
namespace. The following table lists the class names of the classes that
correspond to the **document**, **body**, **p**, **r**, and **t** elements.

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | [Document](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.document.aspx) | The root element for the main document part. |
| body | [Body](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.body.aspx) | The container for the block level structures such as paragraphs, tables, annotations, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| p | [Paragraph](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) | A paragraph. |
| r | [Run](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.run.aspx) | A run. |
| t | [Text](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.text.aspx) | A range of text. |


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

```vb
    ' To get the contents of a document part.
    Public Shared Function GetCommentsFromDocument(ByVal document As String) As String
        Dim comments As String = Nothing

        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
            Dim WordprocessingCommentsPart As WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart
```

You can then use a **StreamReader** object to
read the contents of the **WordprocessingCommentsPart** part of the document
and return its contents.

```csharp
    using (StreamReader streamReader = new StreamReader(WordprocessingCommentsPart.GetStream()))
            {
                comments = streamReader.ReadToEnd();
            }
        }
        return comments;
```

```vb
    Using streamReader As New StreamReader(WordprocessingCommentsPart.GetStream())
    comments = streamReader.ReadToEnd()
    End Using
    Return comments
```

--------------------------------------------------------------------------------
## Sample Code
The following code retrieves the contents of a **WordprocessingCommentsPart** part contained in a
**WordProcessing** document package. You can
run the program by calling the **GetCommentsFromDocument** method as shown in the
following example.

```csharp
    string document = @"C:\Users\Public\Documents\MyPkg5.docx";
    GetCommentsFromDocument(document);
```

```vb
    Dim document As String = "C:\Users\Public\Documents\MyPkg5.docx"
    GetCommentsFromDocument(document)
```

Following is the complete code example in both C\# and Visual Basic.

```csharp
    // To get the contents of a document part.
    public static string GetCommentsFromDocument(string document)
    {
        string comments = null;

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            WordprocessingCommentsPart WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart;

            using (StreamReader streamReader = new StreamReader(WordprocessingCommentsPart.GetStream()))
            {
                comments = streamReader.ReadToEnd();
            }
        }
        return comments;
    }
```

```vb
    ' To get the contents of a document part.
    Public Function GetCommentsFromDocument(ByVal document As String) As String
        Dim comments As String = Nothing
        Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, False)
        Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
        Dim WordprocessingCommentsPart As WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart
        Dim streamReader As StreamReader = New StreamReader(WordprocessingCommentsPart.GetStream)
        comments = streamReader.ReadToEnd
        Return comments
    End Function
```

--------------------------------------------------------------------------------
## See also


[Open XML SDK 2.5 class library
reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
