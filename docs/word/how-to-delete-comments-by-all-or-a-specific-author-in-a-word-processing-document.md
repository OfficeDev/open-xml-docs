---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: c66a64ca-cb0d-4acc-9d05-535b5bbb8c96
title: 'How to: Delete comments by all or a specific author in a word processing document'
description: 'Learn how to delete comments by all or a specific author in a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 02/09/2024
ms.localizationpriority: medium
---
# Delete comments by all or a specific author in a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically delete comments by all or a specific author
in a word processing document, without having to load the document into
Microsoft Word. It contains an example `DeleteComments` method to illustrate this task.



--------------------------------------------------------------------------------

## DeleteComments Method

You can use the `DeleteComments` method to
delete all of the comments from a word processing document, or only
those written by a specific author. As shown in the following code, the
method accepts two parameters that indicate the name of the document to
modify (string) and, optionally, the name of the author whose comments
you want to delete (string). If you supply an author name, the code
deletes comments written by the specified author. If you do not supply
an author name, the code deletes all comments.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/delete_comments_by_all_or_a_specific_author/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/delete_comments_by_all_or_a_specific_author/vb/Program.vb#snippet1)]
***


--------------------------------------------------------------------------------

## Calling the DeleteComments Method

To call the `DeleteComments` method, provide
the required parameters as shown in the following code.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/delete_comments_by_all_or_a_specific_author/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/delete_comments_by_all_or_a_specific_author/vb/Program.vb#snippet2)]
***


--------------------------------------------------------------------------------
## How the Code Works

The following code starts by opening the document, using the
<xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A?displayProperty=nameWithType>
method and indicating that the document should be open for read/write access (the
final `true` parameter value). Next, the code retrieves a reference to the comments
part, using the <xref:DocumentFormat.OpenXml.Packaging.MainDocumentPart.WordprocessingCommentsPart>
property of the main document part, after having retrieved a reference to the main
document part from the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart>
property of the word processing document. If the comments part is missing, there is no point
in proceeding, as there cannot be any comments to delete.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/delete_comments_by_all_or_a_specific_author/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/delete_comments_by_all_or_a_specific_author/vb/Program.vb#snippet3)]
***


--------------------------------------------------------------------------------

## Creating the List of Comments

The code next performs two tasks: creating a list of all the comments to
delete, and creating a list of comment IDs that correspond to the
comments to delete. Given these lists, the code can both delete the
comments from the comments part that contains the comments, and delete
the references to the comments from the document part.The following code
starts by retrieving a list of <xref:DocumentFormat.OpenXml.Wordprocessing.Comment>
elements. To retrieve the list, it converts the <xref:DocumentFormat.OpenXml.OpenXmlElement.Elements>
collection exposed by the `commentPart` variable into a list of `Comment` objects.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/delete_comments_by_all_or_a_specific_author/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/delete_comments_by_all_or_a_specific_author/vb/Program.vb#snippet4)]
***


So far, the list of comments contains all of the comments. If the author
parameter is not an empty string, the following code limits the list to
only those comments where the <xref:DocumentFormat.OpenXml.Wordprocessing.Comment.Author>
property matches the parameter you supplied.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/delete_comments_by_all_or_a_specific_author/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/delete_comments_by_all_or_a_specific_author/vb/Program.vb#snippet5)]
***


Before deleting any comments, the code retrieves a list of comments ID
values, so that it can later delete matching elements from the document
part. The call to the <xref:System.Linq.Enumerable.Select%2A>
method effectively projects the list of comments, retrieving an 
<xref:System.Collections.Generic.IEnumerable%601> of strings that
contain all the comment ID values.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/word/delete_comments_by_all_or_a_specific_author/cs/Program.cs#snippet6)]
### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/word/delete_comments_by_all_or_a_specific_author/vb/Program.vb#snippet6)]
***


--------------------------------------------------------------------------------

## Deleting Comments and Saving the Part

Given the `commentsToDelete` collection, to
the following code loops through all the comments that require deleting
and performs the deletion.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/word/delete_comments_by_all_or_a_specific_author/cs/Program.cs#snippet7)]
### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/word/delete_comments_by_all_or_a_specific_author/vb/Program.vb#snippet7)]
***


--------------------------------------------------------------------------------

## Deleting Comment References in the Document

Although the code has successfully removed all the comments by this
point, that is not enough. The code must also remove references to the
comments from the document part. This action requires three steps
because the comment reference includes the
<xref:DocumentFormat.OpenXml.Wordprocessing.CommentRangeStart>,
<xref:DocumentFormat.OpenXml.Wordprocessing.CommentRangeEnd>,
and <xref:DocumentFormat.OpenXml.Wordprocessing.CommentReference>
elements, and the code must remove all three for each comment.
Before performing any deletions, the code first retrieves a reference
to the root element of the main document part, as shown in the following code.

### [C#](#tab/cs-7)
```csharp
    Document doc = document.MainDocumentPart.Document;
```

### [Visual Basic](#tab/vb-7)
```vb
    Dim doc As Document = document.MainDocumentPart.Document
```
***


Given a reference to the document element, the following code performs
its deletion loop three times, once for each of the different elements
it must delete. In each case, the code looks for all descendants of the
correct type (`CommentRangeStart`, `CommentRangeEnd`, or `CommentReference`)
and limits the list to those whose <xref:DocumentFormat.OpenXml.Wordprocessing.MarkupRangeType.Id>
property value is contained in the list of comment IDs to be deleted.
Given the list of elements to be deleted, the code removes each element in turn.
Finally, the code completes by saving the document.

### [C#](#tab/cs-8)
[!code-csharp[](../../samples/word/delete_comments_by_all_or_a_specific_author/cs/Program.cs#snippet9)]
### [Visual Basic](#tab/vb-8)
[!code-vb[](../../samples/word/delete_comments_by_all_or_a_specific_author/vb/Program.vb#snippet9)]
***


--------------------------------------------------------------------------------

## Sample Code

The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/delete_comments_by_all_or_a_specific_author/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/delete_comments_by_all_or_a_specific_author/vb/Program.vb#snippet)]

--------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
