---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 70839c86-36ef-4b67-a682-abd5114b2bfe
title: 'How to: Retrieve comments from a word processing document (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Retrieve comments from a word processing document (Open XML SDK)

This topic describes how to use the classes in the Open XML SDK for
Office to programmatically retrieve the comments from the main document
part in a word processing document.



--------------------------------------------------------------------------------
## Open the Existing Document for Read-only Access
To open an existing document, instantiate the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class as shown in
the following **using** statement. In the same
statement, open the word processing file at the specified **fileName** by using the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562234.aspx) method. To open the
file for editing the Boolean parameter is set to **true**. In this example you just need to read the
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
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **wordDoc**.


--------------------------------------------------------------------------------
## Comments Element
The **comments** and **comment** elements are crucial to working with
comments in a word processing file. It is important in this code example
to familiarize yourself with those elements.

The following information from the [ISO/IEC
29500](https://www.iso.org/standard/71691.html) specification
introduces the comments element.

> **comments (Comments Collection)**
> 
> This element specifies all of the comments defined in the current
> document. It is the root element of the comments part of a
> WordprocessingML document.
> 
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
## Comment Element
The following information from the [ISO/IEC
29500](https://www.iso.org/standard/71691.html) specification
introduces the comment element.

> **comment (Comment Content)**
> 
> This element specifies the content of a single comment stored in the
> comments part of a WordprocessingML document.
> 
> If a comment is not referenced by document content via a matching
> **id** attribute on a valid use of the **commentReference** element,
> then it may be ignored when loading the document. If more than one
> comment shares the same value for the **id** attribute, then only one
> comment shall be loaded and the others may be ignored.
> 
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
> 
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
## How the Sample Code Works
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
## Sample Code
The following code example shows how to retrieve comments that have been
inserted into a word processing document. To call the method **GetCommentsFromDocument** you can use the following
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

### [C#](#tab/cs)
[!code-csharp[](../samples/word/retrieve_comments/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/retrieve_comments/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
