---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 9774217d-1f71-494a-9ab9-a711661f8e83
title: 'How to: Reply to a comment in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 08/12/2025
ms.localizationpriority: medium
---

# Reply to a comment in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to reply to existing comments in a presentation
programmatically.


[!include[Structure](../includes/presentation/structure.md)]

[!include[description of a comment](../includes/presentation/modern-comment-description.md)]

## How the Sample Code Works

The sample code opens the presentation document in the using statement. Then it gets or creates the CommentAuthorsPart, and verifies that there is an existing comment authors part. If there is not, it adds one.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/reply_to_comment/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/reply_to_comment/vb/Program.vb#snippet1)]
***

Next the code determines if the author that is passed in is on the list of existing authors; if so, it assigns the existing author ID. If not, it adds a new author to the list of authors and assigns an author ID and the parameter values.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/reply_to_comment/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/reply_to_comment/vb/Program.vb#snippet2)]
***

Next the code gets the first slide part and verifies that it exists, then checks if there are any comment parts associated with the slide.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/reply_to_comment/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/reply_to_comment/vb/Program.vb#snippet3)]
***

The code then retrieves the comment list and then iterates through each comment in the comment list, displays the comment text to the user, and prompts whether they want to reply to each comment.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/presentation/reply_to_comment/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/presentation/reply_to_comment/vb/Program.vb#snippet4)]
***

When the user chooses to reply to a comment, the code prompts for the reply text, then gets or creates a `CommentReplyList` for the comment and adds the new reply with the appropriate author information and timestamp.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/presentation/reply_to_comment/cs/Program.cs#snippet5)]

### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/presentation/reply_to_comment/vb/Program.vb#snippet5)]
***

## Sample Code

Following is the complete code sample showing how to reply to existing comments
in a presentation slide with modern PowerPoint comments.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/reply_to_comment/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/reply_to_comment/vb/Program.vb#snippet0)]
***

## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
