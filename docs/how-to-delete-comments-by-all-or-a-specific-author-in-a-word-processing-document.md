---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: c66a64ca-cb0d-4acc-9d05-535b5bbb8c96
title: 'How to: Delete comments by all or a specific author in a word processing document (Open XML SDK)'
description: 'Learn how to delete comments by all or a specific author in a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: medium
---
# Delete comments by all or a specific author in a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically delete comments by all or a specific author
in a word processing document, without having to load the document into
Microsoft Word. It contains an example **DeleteComments** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml). You
must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

--------------------------------------------------------------------------------

## DeleteComments Method

You can use the **DeleteComments** method to
delete all of the comments from a word processing document, or only
those written by a specific author. As shown in the following code, the
method accepts two parameters that indicate the name of the document to
modify (string) and, optionally, the name of the author whose comments
you want to delete (string). If you supply an author name, the code
deletes comments written by the specified author. If you do not supply
an author name, the code deletes all comments.

```csharp
    // Delete comments by a specific author. Pass an empty string for the 
    // author to delete all comments, by all authors.
    public static void DeleteComments(string fileName, 
        string author = "")
```

```vb
    ' Delete comments by a specific author. Pass an empty string for the author 
    ' to delete all comments, by all authors.
    Public Sub DeleteComments(ByVal fileName As String,
        Optional ByVal author As String = "")
```

--------------------------------------------------------------------------------

## Calling the DeleteComments Method

To call the **DeleteComments** method, provide
the required parameters as shown in the following code.

```csharp
    DeleteComments(@"C:\Users\Public\Documents\DeleteComments.docx",
    "David Jones");
```

```vb
    DeleteComments("C:\Users\Public\Documents\DeleteComments.docx",
    "David Jones")
```

--------------------------------------------------------------------------------
## How the Code Works
The following code starts by opening the document, using the [WordprocessingDocument.Open](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx) method and
indicating that the document should be open for read/write access (the
final **true** parameter value). Next, the code
retrieves a reference to the comments part, using the [WordprocessingCommentsPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.maindocumentpart.wordprocessingcommentspart.aspx) property of the
main document part, after having retrieved a reference to the main
document part from the [MainDocumentPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.maindocumentpart.aspx) property of the word
processing document. If the comments part is missing, there is no point
in proceeding, as there cannot be any comments to delete.

```csharp
    // Get an existing Wordprocessing document.
    using (WordprocessingDocument document =
      WordprocessingDocument.Open(fileName, true))
    {
        // Set commentPart to the document WordprocessingCommentsPart, 
        // if it exists.
        WordprocessingCommentsPart commentPart =
          document.MainDocumentPart.WordprocessingCommentsPart;

        // If no WordprocessingCommentsPart exists, there can be no 
        // comments. Stop execution and return from the method.
        if (commentPart == null)
        {
            return;
        }
        // Code removed here…
    }
```

```vb
    ' Get an existing Wordprocessing document.
    Using document As WordprocessingDocument =
        WordprocessingDocument.Open(fileName, True)
        ' Set commentPart to the document 
        ' WordprocessingCommentsPart, if it exists.
        Dim commentPart As WordprocessingCommentsPart =
          document.MainDocumentPart.WordprocessingCommentsPart

        ' If no WordprocessingCommentsPart exists, there can be no
        ' comments. Stop execution and return from the method.
        If (commentPart Is Nothing) Then
            Return
        End If
        ' Code removed here…
    End Using
```

--------------------------------------------------------------------------------

## Creating the List of Comments

The code next performs two tasks: creating a list of all the comments to
delete, and creating a list of comment IDs that correspond to the
comments to delete. Given these lists, the code can both delete the
comments from the comments part that contains the comments, and delete
the references to the comments from the document part.The following code
starts by retrieving a list of [Comment](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.comment.aspx) elements. To retrieve the list, it
converts the [Elements](https://msdn.microsoft.com/library/office/documentformat.openxml.openxmlelement.elements.aspx) collection exposed by the **commentPart** variable into a list of **Comment** objects.

```csharp
    List<Comment> commentsToDelete =
        commentPart.Comments.Elements<Comment>().ToList();
```

```vb
    Dim commentsToDelete As List(Of Comment) = _
        commentPart.Comments.Elements(Of Comment)().ToList()
```

So far, the list of comments contains all of the comments. If the author
parameter is not an empty string, the following code limits the list to
only those comments where the [Author](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.comment.author.aspx) property matches the parameter you
supplied.

```csharp
    if (!String.IsNullOrEmpty(author))
    {
        commentsToDelete = commentsToDelete.
        Where(c => c.Author == author).ToList();
    }
```

```vb
    If Not String.IsNullOrEmpty(author) Then
        commentsToDelete = commentsToDelete.
        Where(Function(c) c.Author = author).ToList()
    End If
```

Before deleting any comments, the code retrieves a list of comments ID
values, so that it can later delete matching elements from the document
part. The call to the [Select](https://msdn2.microsoft.com/library/bb357126)
method effectively projects the list of comments, retrieving an [IEnumerable\<T\>](https://msdn2.microsoft.com/library/9eekhta0)
of strings that contain all the comment ID values.

```csharp
    IEnumerable<string> commentIds = 
        commentsToDelete.Select(r => r.Id.Value);
```

```vb
    Dim commentIds As IEnumerable(Of String) =
        commentsToDelete.Select(Function(r) r.Id.Value)
```

--------------------------------------------------------------------------------

## Deleting Comments and Saving the Part

Given the **commentsToDelete** collection, to
the following code loops through all the comments that require deleting
and performs the deletion. The code then saves the comments part.

```csharp
    // Delete each comment in commentToDelete from the 
    // Comments collection.
    foreach (Comment c in commentsToDelete)
    {
        c.Remove();
    }

    // Save the comment part changes.
    commentPart.Comments.Save();
```

```vb
    ' Delete each comment in commentToDelete from the Comments 
    ' collection.
    For Each c As Comment In commentsToDelete
        c.Remove()
    Next

    ' Save the comment part changes.
    commentPart.Comments.Save()
```

--------------------------------------------------------------------------------

## Deleting Comment References in the Document

Although the code has successfully removed all the comments by this
point, that is not enough. The code must also remove references to the
comments from the document part. This action requires three steps
because the comment reference includes the [CommentRangeStart](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.commentrangestart.aspx), [CommentRangeEnd](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.commentrangeend.aspx), and [CommentReference](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.commentreference.aspx) elements, and the code
must remove all three for each comment. Before performing any deletions,
the code first retrieves a reference to the root element of the main
document part, as shown in the following code.

```csharp
    Document doc = document.MainDocumentPart.Document;
```

```vb
    Dim doc As Document = document.MainDocumentPart.Document
```

Given a reference to the document element, the following code performs
its deletion loop three times, once for each of the different elements
it must delete. In each case, the code looks for all descendants of the
correct type (**CommentRangeStart**, **CommentRangeEnd**, or **CommentReference**) and limits the list to those
whose [Id](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.markuprangetype.id.aspx) property value is contained in the list
of comment IDs to be deleted. Given the list of elements to be deleted,
the code removes each element in turn. Finally, the code completes by
saving the document.

```csharp
    // Delete CommentRangeStart for each
    // deleted comment in the main document.
    List<CommentRangeStart> commentRangeStartToDelete =
        doc.Descendants<CommentRangeStart>().
        Where(c => commentIds.Contains(c.Id.Value)).ToList();
    foreach (CommentRangeStart c in commentRangeStartToDelete)
    {
        c.Remove();
    }

    // Delete CommentRangeEnd for each deleted comment in the main document.
    List<CommentRangeEnd> commentRangeEndToDelete =
        doc.Descendants<CommentRangeEnd>().
        Where(c => commentIds.Contains(c.Id.Value)).ToList();
    foreach (CommentRangeEnd c in commentRangeEndToDelete)
    {
        c.Remove();
    }

    // Delete CommentReference for each deleted comment in the main document.
    List<CommentReference> commentRangeReferenceToDelete =
        doc.Descendants<CommentReference>().
        Where(c => commentIds.Contains(c.Id.Value)).ToList();
    foreach (CommentReference c in commentRangeReferenceToDelete)
    {
        c.Remove();
    }

    // Save changes back to the MainDocumentPart part.
    doc.Save();
```

```vb
    ' Delete CommentRangeStart for each 
    ' deleted comment in the main document.
    Dim commentRangeStartToDelete As List(Of CommentRangeStart) = _
        doc.Descendants(Of CommentRangeStart). _
        Where(Function(c) commentIds.Contains(c.Id.Value)).ToList()
    For Each c As CommentRangeStart In commentRangeStartToDelete
        c.Remove()
    Next

    ' Delete CommentRangeEnd for each deleted comment in the main document.
    Dim commentRangeEndToDelete As List(Of CommentRangeEnd) = _
        doc.Descendants(Of CommentRangeEnd). _
        Where(Function(c) commentIds.Contains(c.Id.Value)).ToList()
    For Each c As CommentRangeEnd In commentRangeEndToDelete
        c.Remove()
    Next

    ' Delete CommentReference for each deleted comment in the main document.
    Dim commentRangeReferenceToDelete As List(Of CommentReference) = _
        doc.Descendants(Of CommentReference). _
        Where(Function(c) commentIds.Contains(c.Id.Value)).ToList
    For Each c As CommentReference In commentRangeReferenceToDelete
        c.Remove()
    Next

    ' Save changes back to the MainDocumentPart part.
    doc.Save()
```

--------------------------------------------------------------------------------

## Sample Code

The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/word/delete_comments_by_all_or_a_specific_author/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/delete_comments_by_all_or_a_specific_author/vb/Program.vb)]

--------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
