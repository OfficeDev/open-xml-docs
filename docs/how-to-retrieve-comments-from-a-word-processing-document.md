---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 70839c86-36ef-4b67-a682-abd5114b2bfe
title: 'How to: Retrieve comments from a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Retrieve comments from a word processing document (Open XML SDK)

This topic describes how to use the classes in the Open XML SDK 2.5 for
Office to programmatically retrieve the comments from the main document
part in a word processing document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports System
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

--------------------------------------------------------------------------------

To open an existing document, instantiate the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> class as shown in
the following **using** statement. In the same
statement, open the word processing file at the specified <span
class="term">fileName</span> by using the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)"><span
class="nolink">Open(String, Boolean)</span></span> method. To open the
file for editing the Boolean parameter is set to <span
class="keyword">true</span>. In this example you just need to read the
file; therefore, you can open the file for read-only access by setting
the Boolean parameter to **false**.

```csharp
    using (WordprocessingDocument wordDoc = 
           WordprocessingDocument.Open(fileName, false)) 
    { 
       // Insert other code here. 
    }
```

```vb
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(fileName, False)
        ' Insert other code here.
    End Using
```
The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case <span
class="term">wordDoc</span>.


--------------------------------------------------------------------------------

The **comments** and <span
class="keyword">comment</span> elements are crucial to working with
comments in a word processing file. It is important in this code example
to familiarize yourself with those elements.

The following information from the [ISO/IEC
29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification
introduces the comments element.

> **comments (Comments Collection)**

> This element specifies all of the comments defined in the current
> document. It is the root element of the comments part of a
> WordprocessingML document.

> Consider the following WordprocessingML fragment for the content of a
> comments part in a WordprocessingML document:

```xml
    <w:comments>
      <w:comment … >
        …
      </w:comment>
    </w:comments>
```

> © ISO/IEC29500: 2008.

The following XML schema segment defines the contents of the comments
element.

```xml
    <complexType name="CT_Comments">
       <sequence>
           <element name="comment" type="CT_Comment" minOccurs="0" maxOccurs="unbounded"/>
       </sequence>
    </complexType>
```

---------------------------------------------------------------------------------

The following information from the [ISO/IEC
29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification
introduces the comment element.

> **comment (Comment Content)**

> This element specifies the content of a single comment stored in the
> comments part of a WordprocessingML document.

> If a comment is not referenced by document content via a matching
> **id** attribute on a valid use of the **commentReference** element,
> then it may be ignored when loading the document. If more than one
> comment shares the same value for the **id** attribute, then only one
> comment shall be loaded and the others may be ignored.

> Consider a document with text with an annotated comment as follows:

![Document text with annotated comment](./media/w-comment01.gif)
> This comment is represented by the following WordprocessingML
> fragment.

```xml
    <w:comment w:id="1" w:initials="User">
      …
    </w:comment>
```
> The **comment** element specifies the presence of a single comment
> within the comments part.

> © ISO/IEC29500: 2008.

  
The following XML schema segment defines the contents of the comment
element.

```xml
    <complexType name="CT_Comment">
       <complexContent>
           <extension base="CT_TrackChange">
              <sequence>
                  <group ref="EG_BlockLevelElts" minOccurs="0" maxOccurs="unbounded"/>
              </sequence>
              <attribute name="initials" type="ST_String" use="optional"/>
           </extension>
       </complexContent>
    </complexType>
```

--------------------------------------------------------------------------------

After you have opened the file for read-only access, you instantiate the
**WordprocessingCommentsPart** class. You can
then display the inner text of the **Comment**
element.

```csharp
    foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
    {
        Console.WriteLine(comment.InnerText);
    }
```

```vb
    For Each comment As Comment In _
        commentsPart.Comments.Elements(Of Comment)()
        Console.WriteLine(comment.InnerText)
    Next
```

--------------------------------------------------------------------------------

The following code example shows how to retrieve comments that have been
inserted into a word processing document. To call the method <span
class="keyword">GetCommentsFromDocument</span> you can use the following
call, which retrieves comments from a file named "Word16.docx," as an
example.

```csharp
    string fileName = @"C:\Users\Public\Documents\Word16.docx";
    GetCommentsFromDocument(fileName);
```

```vb
    Dim fileName As String = "C:\Users\Public\Documents\Word16.docx"
    GetCommentsFromDocument(fileName)
```

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    public static void GetCommentsFromDocument(string fileName)
    {
        using (WordprocessingDocument wordDoc = 
            WordprocessingDocument.Open(fileName, false))
        {
            WordprocessingCommentsPart commentsPart = 
                wordDoc.MainDocumentPart.WordprocessingCommentsPart;

            if (commentsPart != null && commentsPart.Comments != null)
            {
                foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
                {
                    Console.WriteLine(comment.InnerText);
                }
            }
        }
    }
```

```vb
    Public Sub GetCommentsFromDocument(ByVal fileName As String)
        Using wordDoc As WordprocessingDocument = _
            WordprocessingDocument.Open(fileName, False)

            Dim commentsPart As WordprocessingCommentsPart = _
                wordDoc.MainDocumentPart.WordprocessingCommentsPart

            If commentsPart IsNot Nothing AndAlso _
                commentsPart.Comments IsNot Nothing Then
                For Each comment As Comment In _
                    commentsPart.Comments.Elements(Of Comment)()
                    Console.WriteLine(comment.InnerText)
                Next
            End If
        End Using
    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
