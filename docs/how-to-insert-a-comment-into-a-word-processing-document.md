---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 474f0a6c-62c8-4f04-b3f9-cd613a6e48d0
title: 'How to: Insert a comment into a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
localization_priority: Normal
---
# How to: Insert a comment into a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically add a comment to the first paragraph in a
word processing document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports System
    Imports System.Linq  
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

--------------------------------------------------------------------------------
## Open the Existing Document for Editing
To open an existing document, instantiate the [WordprocessingDocument](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class as shown in
the following **using** statement. In the same
statement, open the word processing file at the specified *filepath* by
using the [Open(String, Boolean)](https://msdn.microsoft.com/en-us/library/office/cc562234.aspx) method, with the
Boolean parameter set to **true** to enable
editing in the document.

```csharp
    using (WordprocessingDocument document =
           WordprocessingDocument.Open(filepath, true)) 
    { 
       // Insert other code here. 
    }
```

```vb
    Using document As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
       ' Insert other code here. 
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
that is used by the Open XML SDK to clean up resources) is automatically
called when the closing brace is reached. The block that follows the
**using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case *document*.


--------------------------------------------------------------------------------
## How the Sample Code Works
After you open the document, you can find the first paragraph to attach
a comment. The code finds the first paragraph by calling the
[First](http://msdn.microsoft.com/en-us/library/system.linq.enumerable.first.aspx)
extension method on all the descendant elements of the document element
that are of type [Paragraph](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.paragraph.aspx). The **First** method is a member
of the
[System.Linq.Enumerable](http://msdn.microsoft.com/en-us/library/system.linq.enumerable.aspx)
class. The **System.Linq.Enumerable** class
provides extension methods for objects that implement the
[System.Collections.Generic.IEnumerable](http://msdn.microsoft.com/en-us/library/9eekhta0.aspx)
interface.

```csharp
    Paragraph firstParagraph = document.MainDocumentPart.Document.Descendants<Paragraph>().First();
    Comments comments = null;
    string id = "0";
```

```vb
    Dim firstParagraph As Paragraph = document.MainDocumentPart.Document.Descendants(Of Paragraph)().First()
    Dim comments As Comments = Nothing
    Dim id As String = "0"
```

The code first determines whether a [WordprocessingCommentsPart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.wordprocessingcommentspart.aspx) part exists. To
do this, call the [MainDocumentPart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.maindocumentpart.aspx) generic method, **GetPartsCountOfType**, and specify a kind of **WordprocessingCommentsPart**.

If a **WordprocessingCommentsPart** part
exists, the code obtains a new **Id** value for
the [Comment](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.comment.aspx) object that it will add to the
existing **WordprocessingCommentsPart** [Comments](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.comments.aspx) collection object. It does this by
finding the highest **Id** attribute value
given to a **Comment** in the **Comments** collection object, incrementing the
value by one, and then storing that as the **Id** value.If no **WordprocessingCommentsPart** part exists, the code
creates one using the [AddNewPart\<T\>()](https://msdn.microsoft.com/en-us/library/office/cc562657.aspx) method of the [MainDocumentPart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.maindocumentpart.aspx) object and then adds a
**Comments** collection object to it.

```csharp
    if (document.MainDocumentPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
    {
        comments = 
            document.MainDocumentPart.WordprocessingCommentsPart.Comments;
        if (comments.HasChildren)
        {
            id = comments.Descendants<Comment>().Select(e => e.Id.Value).Max();
        }
    }
    else
    {
        WordprocessingCommentsPart commentPart = 
                    document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
       commentPart.Comments = new Comments();
       comments = commentPart.Comments;
    }
```

```vb
    If document.MainDocumentPart.GetPartsCountOfType(Of WordprocessingCommentsPart)() > 0 Then
        comments = document.MainDocumentPart.WordprocessingCommentsPart.Comments
        If comments.HasChildren Then
            id = comments.Descendants(Of Comment)().Select(Function(e) e.Id.Value).Max()
        End If
    Else
        Dim commentPart As WordprocessingCommentsPart = document.MainDocumentPart.AddNewPart(Of WordprocessingCommentsPart)()
        commentPart.Comments = New Comments()
        comments = commentPart.Comments
    End If
```

The **Comment** and **Comments** objects represent comment and comments
elements, respectively, in the Open XML Wordprocessing schema. A **Comment** must be added to a **Comments** object so the code first instantiates a
**Comments** object (using the string arguments
**author**, **initials**,
and **comments** that were passed in to the **AddCommentOnFirstParagraph** method).

The comment is represented by the following WordprocessingML code
example. .

```xml
    <w:comment w:id="1" w:initials="User">
      ...
    </w:comment>
```

The code then appends the **Comment** to the
**Comments** object and saves the changes. This
creates the required XML document object model (DOM) tree structure in
memory which consists of a **comments** parent
element with **comment** child elements under
it.

```csharp
    Paragraph p = new Paragraph(new Run(new Text(comment)));
    Comment cmt = new Comment() { Id = id, 
            Author = author, Initials = initials, Date = DateTime.Now };
    cmt.AppendChild(p);
    comments.AppendChild(cmt);
    comments.Save();
```

```vb
    Dim p As New Paragraph(New Run(New Text(comment)))
    Dim cmt As New Comment() With {.Id = id, .Author = author, .Initials = initials, .Date = Date.Now}
    cmt.AppendChild(p)
    comments.AppendChild(cmt)
    comments.Save()
```

The following WordprocessingML code example represents the content of a
comments part in a WordprocessingML document.

```xml
    <w:comments>
      <w:comment … >
        …
      </w:comment>
    </w:comments>
```

With the **Comment** object instantiated, the
code associates the **Comment** with a range in
the Wordprocessing document. [CommentRangeStart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.commentrangestart.aspx) and [CommentRangeEnd](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.commentrangeend.aspx) objects correspond to the
**commentRangeStart** and **commentRangeEnd** elements in the Open XML
Wordprocessing schema. A **CommentRangeStart**
object is given as the argument to the [InsertBefore\<T\>(T, OpenXmlElement)](https://msdn.microsoft.com/en-us/library/office/cc863927.aspx) method
of the [Paragraph](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) object and a **CommentRangeEnd** object is passed to the [InsertAfter\<T\>(T, OpenXmlElement)](https://msdn.microsoft.com/en-us/library/office/cc865786.aspx) method.
This creates a comment range that extends from immediately before the
first character of the first paragraph in the Wordprocessing document to
immediately after the last character of the first paragraph.

A [CommentReference](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.commentreference.aspx) object represents a
**commentReference** element in the Open XML Wordprocessing schema. A
commentReference links a specific comment in the **WordprocessingCommentsPart** part (the Comments.xml
file in the Wordprocessing package) to a specific location in the
document body (the **MainDocumentPart** part
contained in the Document.xml file in the Wordprocessing package). The
**id** attribute of the comment,
commentRangeStart, commentRangeEnd, and commentReference is the same for
a given comment, so the commentReference **id**
attribute must match the comment **id** attribute
value that it links to. In the sample, the code adds a **commentReference** element by using the API, and
instantiates a **CommentReference** object,
specifying the **Id** value, and then adds it to a [Run](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.run.aspx) object.

```csharp
    firstParagraph.InsertBefore(new CommentRangeStart() 
                { Id = id }, firstParagraph.GetFirstChild<Run>());

            var cmtEnd = firstParagraph.InsertAfter(new CommentRangeEnd() 
                { Id = id }, firstParagraph.Elements<Run>().Last());

            firstParagraph.InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
```

```vb
    firstParagraph.InsertBefore(New CommentRangeStart() With {.Id = id}, firstParagraph.GetFirstChild(Of Run)())

            Dim cmtEnd = firstParagraph.InsertAfter(New CommentRangeEnd() With {.Id = id}, firstParagraph.Elements(Of Run)().Last())

            firstParagraph.InsertAfter(New Run(New CommentReference() With {.Id = id}), cmtEnd)
```

--------------------------------------------------------------------------------
## Sample Code
The following code example shows how to create a comment and associate
it with a range in a word processing document. To call the method **AddCommentOnFirstParagraph** pass in the path of
the document, your name, your initials, and the comment text. For
example, the following call to the **AddCommentOnFirstParagraph** method writes the
comment "This is my comment." in the file "Word8.docx."

```csharp
    AddCommentOnFirstParagraph(@"C:\Users\Public\Documents\Word8.docx",
     author, initials, "This is my comment.");
```

```vb
    AddCommentOnFirstParagraph("C:\Users\Public\Documents\Word8.docx", _
     author, initials, comment)
```

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Insert a comment on the first paragraph.
    public static void AddCommentOnFirstParagraph(string fileName, 
        string author, string initials, string comment)
    {
        // Use the file name and path passed in as an 
        // argument to open an existing Wordprocessing document. 
        using (WordprocessingDocument document = 
            WordprocessingDocument.Open(fileName, true))
        {
            // Locate the first paragraph in the document.
            Paragraph firstParagraph = 
                document.MainDocumentPart.Document.Descendants<Paragraph>().First();
            Comments comments = null;
            string id = "0";

            // Verify that the document contains a 
            // WordProcessingCommentsPart part; if not, add a new one.
            if (document.MainDocumentPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
            {
                comments = 
                    document.MainDocumentPart.WordprocessingCommentsPart.Comments;
                if (comments.HasChildren)
                {
                    // Obtain an unused ID.
                    id = comments.Descendants<Comment>().Select(e => e.Id.Value).Max();
                }
            }
            else
            {
                // No WordprocessingCommentsPart part exists, so add one to the package.
                WordprocessingCommentsPart commentPart = 
                    document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                commentPart.Comments = new Comments();
                comments = commentPart.Comments;
            }

            // Compose a new Comment and add it to the Comments part.
            Paragraph p = new Paragraph(new Run(new Text(comment)));
            Comment cmt = 
                new Comment() { Id = id, 
                    Author = author, Initials = initials, Date = DateTime.Now };
            cmt.AppendChild(p);
            comments.AppendChild(cmt);
            comments.Save();

            // Specify the text range for the Comment. 
            // Insert the new CommentRangeStart before the first run of paragraph.
            firstParagraph.InsertBefore(new CommentRangeStart() 
                { Id = id }, firstParagraph.GetFirstChild<Run>());

            // Insert the new CommentRangeEnd after last run of paragraph.
            var cmtEnd = firstParagraph.InsertAfter(new CommentRangeEnd() 
                { Id = id }, firstParagraph.Elements<Run>().Last());

            // Compose a run with CommentReference and insert it.
            firstParagraph.InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
        }
    }
```

```vb
    ' Insert a comment on the first paragraph.
    Public Sub AddCommentOnFirstParagraph(ByVal fileName As String, ByVal author As String, ByVal initials As String, ByVal comment As String)
        ' Use the file name and path passed in as an 
        ' argument to open an existing Wordprocessing document. 
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            ' Locate the first paragraph in the document.
            Dim firstParagraph As Paragraph = document.MainDocumentPart.Document.Descendants(Of Paragraph)().First()
            Dim comments As Comments = Nothing
            Dim id As String = "0"

            ' Verify that the document contains a 
            ' WordProcessingCommentsPart part; if not, add a new one.
            If document.MainDocumentPart.GetPartsCountOfType(Of WordprocessingCommentsPart)() > 0 Then
                comments = document.MainDocumentPart.WordprocessingCommentsPart.Comments
                If comments.HasChildren Then
                    ' Obtain an unused ID.
                    id = comments.Descendants(Of Comment)().[Select](Function(e) e.Id.Value).Max()
                End If
            Else
                ' No WordprocessingCommentsPart part exists, so add one to the package.
                Dim commentPart As WordprocessingCommentsPart = document.MainDocumentPart.AddNewPart(Of WordprocessingCommentsPart)()
                commentPart.Comments = New Comments()
                comments = commentPart.Comments
            End If

            ' Compose a new Comment and add it to the Comments part.
            Dim p As New Paragraph(New Run(New Text(comment)))
            Dim cmt As New Comment() With {.Id = id, .Author = author, .Initials = initials, .Date = DateTime.Now}
            cmt.AppendChild(p)
            comments.AppendChild(cmt)
            comments.Save()

            ' Specify the text range for the Comment. 
            ' Insert the new CommentRangeStart before the first run of paragraph.
            firstParagraph.InsertBefore(New CommentRangeStart() With {.Id = id}, firstParagraph.GetFirstChild(Of Run)())

            ' Insert the new CommentRangeEnd after last run of paragraph.
            Dim cmtEnd = firstParagraph.InsertAfter(New CommentRangeEnd() With {.Id = id}, firstParagraph.Elements(Of Run)().Last())

            ' Compose a run with CommentReference and insert it.
            firstParagraph.InsertAfter(New Run(New CommentReference() With {.Id = id}), cmtEnd)
        End Using
    End Sub
```

--------------------------------------------------------------------------------
## See also
#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)

[Language-Integrated Query (LINQ)](http://msdn.microsoft.com/en-us/library/bb397926.aspx)

[Extension Methods (C\# Programming Guide)](http://msdn.microsoft.com/en-us/library/bb383977.aspx)

[Extension Methods (Visual Basic)](http://msdn.microsoft.com/en-us/library/bb384936.aspx)
