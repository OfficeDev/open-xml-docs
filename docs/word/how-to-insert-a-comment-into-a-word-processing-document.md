---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 474f0a6c-62c8-4f04-b3f9-cd613a6e48d0
title: 'How to: Insert a comment into a word processing document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Insert a comment into a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically add a comment to the first paragraph in a
word processing document.



--------------------------------------------------------------------------------
## Open the Existing Document for Editing
To open an existing document, instantiate the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class as shown in
the following **using** statement. In the same
statement, open the word processing file at the specified *filepath* by
using the [Open(String, Boolean)](/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument.open) method, with the
Boolean parameter set to **true** to enable
editing in the document.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument document =
           WordprocessingDocument.Open(filepath, true)) 
    { 
       // Insert other code here. 
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    Using document As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
       ' Insert other code here. 
    End Using
```
***

[!include[Using Statement](../includes/using-statement.md)]


--------------------------------------------------------------------------------
## How the Sample Code Works
After you open the document, you can find the first paragraph to attach
a comment. The code finds the first paragraph by calling the
[First](/dotnet/api/system.linq.enumerable.first)
extension method on all the descendant elements of the document element
that are of type [Paragraph](/dotnet/api/documentformat.openxml.wordprocessing.paragraph). The **First** method is a member
of the
[System.Linq.Enumerable](/dotnet/api/system.linq.enumerable)
class. The **System.Linq.Enumerable** class
provides extension methods for objects that implement the
[System.Collections.Generic.IEnumerable](/dotnet/api/system.collections.generic.ienumerable-1)
interface.

### [C#](#tab/cs-1)
```csharp
    Paragraph firstParagraph = document.MainDocumentPart.Document.Descendants<Paragraph>().First();
    Comments comments = null;
    string id = "0";
```

### [Visual Basic](#tab/vb-1)
```vb
    Dim firstParagraph As Paragraph = document.MainDocumentPart.Document.Descendants(Of Paragraph)().First()
    Dim comments As Comments = Nothing
    Dim id As String = "0"
```
***


The code first determines whether a [WordprocessingCommentsPart](/dotnet/api/documentformat.openxml.packaging.wordprocessingcommentspart) part exists. To
do this, call the [MainDocumentPart](/dotnet/api/documentformat.openxml.packaging.maindocumentpart) generic method, **GetPartsCountOfType**, and specify a kind of **WordprocessingCommentsPart**.

If a **WordprocessingCommentsPart** part
exists, the code obtains a new **Id** value for
the [Comment](/dotnet/api/documentformat.openxml.wordprocessing.comment) object that it will add to the
existing **WordprocessingCommentsPart** [Comments](/dotnet/api/documentformat.openxml.wordprocessing.comments) collection object. It does this by
finding the highest **Id** attribute value
given to a **Comment** in the **Comments** collection object, incrementing the
value by one, and then storing that as the **Id** value.If no **WordprocessingCommentsPart** part exists, the code
creates one using the [AddNewPart\<T\>()](/dotnet/api/documentformat.openxml.packaging.openxmlpartcontainer.addnewpart) method of the [MainDocumentPart](/dotnet/api/documentformat.openxml.packaging.maindocumentpart) object and then adds a
**Comments** collection object to it.

### [C#](#tab/cs-2)
```csharp
    if (document.MainDocumentPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
    {
        comments = 
            document.MainDocumentPart.WordprocessingCommentsPart.Comments;
        if (comments.HasChildren)
        {
            id = (comments.Descendants<Comment>().Select(e => int.Parse(e.Id.Value)).Max() + 1).ToString();
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

### [Visual Basic](#tab/vb-2)
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
***


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

### [C#](#tab/cs-3)
```csharp
    Paragraph p = new Paragraph(new Run(new Text(comment)));
    Comment cmt = new Comment() { Id = id, 
            Author = author, Initials = initials, Date = DateTime.Now };
    cmt.AppendChild(p);
    comments.AppendChild(cmt);
    comments.Save();
```

### [Visual Basic](#tab/vb-3)
```vb
    Dim p As New Paragraph(New Run(New Text(comment)))
    Dim cmt As New Comment() With {.Id = id, .Author = author, .Initials = initials, .Date = Date.Now}
    cmt.AppendChild(p)
    comments.AppendChild(cmt)
    comments.Save()
```
***


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
the Wordprocessing document. [CommentRangeStart](/dotnet/api/documentformat.openxml.wordprocessing.commentrangestart) and [CommentRangeEnd](/dotnet/api/documentformat.openxml.wordprocessing.commentrangeend) objects correspond to the
**commentRangeStart** and **commentRangeEnd** elements in the Open XML
Wordprocessing schema. A **CommentRangeStart**
object is given as the argument to the [InsertBefore\<T\>(T, OpenXmlElement)](/dotnet/api/documentformat.openxml.openxmlcompositeelement.insertbefore) method
of the [Paragraph](/dotnet/api/documentformat.openxml.wordprocessing.paragraph) object and a **CommentRangeEnd** object is passed to the [InsertAfter\<T\>(T, OpenXmlElement)](/dotnet/api/documentformat.openxml.openxmlcompositeelement.insertafter) method.
This creates a comment range that extends from immediately before the
first character of the first paragraph in the Wordprocessing document to
immediately after the last character of the first paragraph.

A [CommentReference](/dotnet/api/documentformat.openxml.wordprocessing.commentreference) object represents a
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
specifying the **Id** value, and then adds it to a [Run](/dotnet/api/documentformat.openxml.wordprocessing.run) object.

### [C#](#tab/cs-4)
```csharp
    firstParagraph.InsertBefore(new CommentRangeStart() 
                { Id = id }, firstParagraph.GetFirstChild<Run>());

            var cmtEnd = firstParagraph.InsertAfter(new CommentRangeEnd() 
                { Id = id }, firstParagraph.Elements<Run>().Last());

            firstParagraph.InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
```

### [Visual Basic](#tab/vb-4)
```vb
    firstParagraph.InsertBefore(New CommentRangeStart() With {.Id = id}, firstParagraph.GetFirstChild(Of Run)())

            Dim cmtEnd = firstParagraph.InsertAfter(New CommentRangeEnd() With {.Id = id}, firstParagraph.Elements(Of Run)().Last())

            firstParagraph.InsertAfter(New Run(New CommentReference() With {.Id = id}), cmtEnd)
```
***


--------------------------------------------------------------------------------
## Sample Code
The following code example shows how to create a comment and associate
it with a range in a word processing document. To call the method **AddCommentOnFirstParagraph** pass in the path of
the document, your name, your initials, and the comment text. For
example, the following call to the **AddCommentOnFirstParagraph** method writes the
comment "This is my comment." in the file "Word8.docx."

### [C#](#tab/cs-5)
```csharp
    AddCommentOnFirstParagraph(@"C:\Users\Public\Documents\Word8.docx",
     author, initials, "This is my comment.");
```

### [Visual Basic](#tab/vb-5)
```vb
    AddCommentOnFirstParagraph("C:\Users\Public\Documents\Word8.docx", _
     author, initials, comment)
```
***


Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/insert_a_comment/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/insert_a_comment/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

[Language-Integrated Query (LINQ)](/previous-versions/bb397926(v=vs.140))

[Extension Methods (C\# Programming Guide)](/dotnet/csharp/programming-guide/classes-and-structs/extension-methods)

[Extension Methods (Visual Basic)](/dotnet/visual-basic/programming-guide/language-features/procedures/extension-methods)
