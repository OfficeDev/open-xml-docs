---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 3b892a6a-2972-461e-94a9-0a1ede854bda
title: 'Delete all the comments by an author from all the slides in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 12/30/2024
ms.localizationpriority: medium
---
# Delete all the comments by an author from all the slides in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to delete all of the comments by a specific author in a
presentation programmatically.

> [!NOTE]
> This sample is for PowerPoint modern comments. For classic comments view
> the [archived sample on GitHub](https://github.com/OfficeDev/open-xml-docs/blob/7002d692ab4abc629d617ef6a0214fc2bf2910c8/docs/how-to-delete-all-the-comments-by-an-author-from-all-the-slides-in-a-presentatio.md).





## Getting a PresentationDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the
<xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)> method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read/write,
specify the value `true` for this parameter
as shown in the following `using` statement.
In this code, the *fileName* parameter is a string that represents the
path for the file from which you want to open the document, and the
author is the user name displayed in the General tab of the PowerPoint
Options.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/delete_all_the_comments_by_an_author_from_all_the_slides_a_presentatio/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/delete_all_the_comments_by_an_author_from_all_the_slides_a_presentatio/vb/Program.vb#snippet1)]
***


[!include[Using Statement](../includes/presentation/using-statement.md)] `doc`.

[!include[Structure](../includes/presentation/structure.md)]

## The Structure of the Comment Element

The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces comments in a presentation package.

> A comment is a text note attached to a slide, with the primary purpose
> of allowing readers of a presentation to provide feedback to the
> presentation author. Each comment contains an unformatted text string
> and information about its author, and is attached to a particular
> location on a slide. Comments can be visible while editing the
> presentation, but do not appear when a slide show is given. The
> displaying application decides when to display comments and determines
> their visual appearance.
> 
> &copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

[!include[description of a comment](../includes/presentation/modern-comment-description.md)]

## How the Sample Code Works

After opening the presentation document for read/write access and
instantiating the `PresentationDocument`
class, the code gets the specified comment author from the list of
comment authors.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/delete_all_the_comments_by_an_author_from_all_the_slides_a_presentatio/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/delete_all_the_comments_by_an_author_from_all_the_slides_a_presentatio/vb/Program.vb#snippet2)]
***


By iterating through the matching authors and all the slides in the
presentation the code gets all the slide parts, and the comments part of
each slide part. It then gets the list of comments by the specified
author and deletes each one. It also verifies that the comment part has
no existing comment, in which case it deletes that part. It also deletes
the comment author from the comment authors part.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/delete_all_the_comments_by_an_author_from_all_the_slides_a_presentatio/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/delete_all_the_comments_by_an_author_from_all_the_slides_a_presentatio/vb/Program.vb#snippet3)]
***


## Sample Code

The following method takes as parameters the source presentation file
name and path and the name of the comment author whose comments are to
be deleted. It finds all the comments by the specified author in the
presentation and deletes them. It then deletes the comment author from
the list of comment authors.

> [!NOTE]
> To get the exact author's name, open the presentation file and click the **File** menu item, and then click **Options**. The **PowerPoint Options** window opens and the content of the **General** tab is displayed. The author's name must match the **User name** in this tab.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/delete_all_the_comments_by_an_author_from_all_the_slides_a_presentatio/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/delete_all_the_comments_by_an_author_from_all_the_slides_a_presentatio/vb/Program.vb#snippet0)]
***

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
