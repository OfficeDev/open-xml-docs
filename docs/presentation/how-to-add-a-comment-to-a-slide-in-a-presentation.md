---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 403abe97-7ab2-40ba-92c0-d6312a6d10c8
title: 'How to: Add a comment to a slide in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/02/2025
ms.localizationpriority: medium
---

# Add a comment to a slide in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to add a comment to the first slide in a presentation
programmatically.

> [!NOTE]
> This sample is for PowerPoint modern comments. For classic comments view
> the [archived sample on GitHub](https://github.com/OfficeDev/open-xml-docs/blob/7002d692ab4abc629d617ef6a0214fc2bf2910c8/docs/how-to-add-a-comment-to-a-slide-in-a-presentation.md).


[!include[Structure](../includes/presentation/structure.md)]

[!include[description of a comment](../includes/presentation/modern-comment-description.md)]

## How the Sample Code Works

The sample code opens the presentation document in the using statement. Then it instantiates the CommentAuthorsPart, and verifies that there is an existing comment authors part. If there is not, it adds one.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/add_comment/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/add_comment/vb/Program.vb#snippet1)]
***

The code determines whether there is an existing PowerPoint author part in the presentation part; if not, it adds one, then checks if there is an authors list 
and adds one if it is missing. It also verifies that the author that is passed in is on the list of existing authors; if so, it assigns the existing author ID. If not, it adds a new author to the list of authors and assigns an author ID and the parameter values.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/add_comment/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/add_comment/vb/Program.vb#snippet2)]
***

Next the code determines if there is a slide id and returns if one does not exist

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/add_comment/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/add_comment/vb/Program.vb#snippet3)]
***

In the segment below, the code gets the relationship ID. If it exists, it is used to find the slide part
otherwise the first slide in the slide parts enumerable is taken. Then it verifies that there is 
a PowerPoint comments part for the slide and if not adds one.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/presentation/add_comment/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/presentation/add_comment/vb/Program.vb#snippet4)]
***

Below the code creates a new modern comment then adds a comment list to the PowerPoint comment part
if one does not exist and adds the comment to that comment list.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/presentation/add_comment/cs/Program.cs#snippet5)]

### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/presentation/add_comment/vb/Program.vb#snippet5)]
***

With modern comments the slide needs to have the correct extension list and extension.
The following code determines if the slide already has a SlideExtensionList and
SlideExtension and adds them to the slide if they are not present.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/presentation/add_comment/cs/Program.cs#snippet6)]

### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/presentation/add_comment/vb/Program.vb#snippet6)]
***

## Sample Code

Following is the complete code sample showing how to add a new comment with
a new or existing author to a slide with or without existing comments.

> [!NOTE]
> To get the exact author name and initials, open the presentation file and click the **File** menu item, and then click **Options**. The **PowerPointOptions** window opens and the content of the **General** tab is displayed. The author name and initials must match the **User name** and **Initials** in this tab.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/add_comment/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/add_comment/vb/Program.vb#snippet0)]
***

## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

