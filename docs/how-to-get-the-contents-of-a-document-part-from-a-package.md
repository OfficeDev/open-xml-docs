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
---
# How to: Get the contents of a document part from a package (Open XML SDK)

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

An Open XML document is stored as a package, whose format is defined by
[ISO/IEC 29500-2](http://go.microsoft.com/fwlink/?LinkId=194337). The
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

The code starts with opening a package file by passing a file name to
one of the overloaded <span sdata="cer"
target="Overload:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open"><span
class="nolink">Open()</span></span> methods (Visual Basic .NET Shared
method or C\# static method) of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> class that takes a
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
when the closing brace is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case <span
class="keyword">wordDoc</span>. Because the <span
class="keyword">WordprocessingDocument</span> class in the Open XML SDK
automatically saves and closes the object as part of its <span
class="keyword">System.IDisposable</span> implementation, and because
the **Dispose** method is automatically called
when you exit the block; you do not have to explicitly call <span
class="keyword">Save</span> and **Close**─as
long as you use using.


---------------------------------------------------------------------------------

The basic document structure of a <span
class="keyword">WordProcessingML</span> document consists of the <span
class="keyword">document</span> and **body**
elements, followed by one or more block level elements such as <span
class="keyword">p</span>, which represents a paragraph. A paragraph
contains one or more **r** elements. The <span
class="keyword">r</span> stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The <span
class="keyword">t</span> element contains a range of text. The <span
class="keyword">WordprocessingML</span> markup for the document that the
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
content using strongly-typed classes that correspond to <span
class="keyword">WordprocessingML</span> elements. You can find these
classes in the <span sdata="cer"
target="N:DocumentFormat.OpenXml.Wordprocessing"><span
class="nolink">DocumentFormat.OpenXml.Wordprocessing</span></span>
namespace. The following table lists the class names of the classes that
correspond to the **document**, <span
class="keyword">body</span>, **p**, <span
class="keyword">r</span>, and **t** elements.

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | Document | The root element for the main document part. |
| body | Body | The container for the block level structures such as paragraphs, tables, annotations, and others specified in the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification. |
| p | Paragraph | A paragraph. |
| r | Run | A run. |
| t | Text | A range of text. |


--------------------------------------------------------------------------------

In this how-to, you are going to work with comments. Therefore, it is
useful to familiarize yourself with the structure of the \<<span
class="keyword">comments</span>\> element. The following information
from the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337)
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

After you have opened the source file for reading, you create a <span
class="keyword">mainPart</span> object by instantiating the <span
class="keyword">MainDocumentPart</span>. Then you can create a reference
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
read the contents of the <span
class="keyword">WordprocessingCommentsPart</span> part of the document
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

The following code retrieves the contents of a <span
class="keyword">WordprocessingCommentsPart</span> part contained in a
**WordProcessing** document package. You can
run the program by calling the <span
class="keyword">GetCommentsFromDocument</span> method as shown in the
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

#### Other resources

[Open XML SDK 2.5 class library
reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
