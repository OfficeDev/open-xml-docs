---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: d7f0f1d3-bcf9-40b5-aaa4-4a08d862ac8e
title: Working with comments
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/26/2024
ms.localizationpriority: medium
---
# Working with comments

This topic discusses the Open XML SDK for Office <xref:DocumentFormat.OpenXml.Presentation.Comment> class and how it relates to the
Open XML File Format PresentationML schema. For more information about the overall structure of the parts and elements that make up a PresentationML document, see [Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

## Comments in PresentationML

The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification describes the Comments section of the Open XML PresentationML framework as follows:

A comment is a text note attached to a slide, with the primary purpose
of allowing readers of a presentation to provide feedback to the
presentation author. Each comment contains an unformatted text string
and information about its author, and is attached to a particular
location on a slide. Comments can be visible while editing the
presentation, but do not appear when a slide show is given. The
displaying application decides when to display comments and determines
their visual appearance.

The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification describes the Open XML PresentationML `<cm/>` element used to represent comments in a PresentationML document as follows:

This element specifies a single comment attached to a slide. It contains
the text of the comment, its position on the slide, and attributes
referring to its author and date.

Example:

```xml
<p:cm authorId="0" dt="2006-08-28T17:26:44.129" idx="1">  
   <p:pos x="10" y="10"/>  
   <p:text\>Add diagram to clarify.</p:text>  
</p:cm>
```

The following table lists the child elements of the `<cm/>` element used when working with comments and the Open XML SDK classes that correspond to them.

| **PresentationML Element** |                                                               **Open XML SDK Class**                                                 |
|----------------------------|------------------------------------------------------------------------------------------------------------------------------------------|
| `<extLst/>`        | <xref:DocumentFormat.OpenXml.Presentation.ExtensionListWithModification> |
| `<pos/>`           | <xref:DocumentFormat.OpenXml.Presentation.Position>                  |
| `<text/>`          | <xref:DocumentFormat.OpenXml.Presentation.Text>                          |

The following table from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the attributes of the `<cm/>` element.

| **Attributes** |  **Description**   |
|----------------|--------------------|
|    authorId    | This attribute specifies the author of the comment.It refers to the ID of an author in the comment author list for the document.<br/>The possible values for this attribute are defined by the W3C XML Schema `unsignedInt` datatype.  |
|       dt       | This attribute specifies the date and time this comment was last modified.<br/>The possible values for this attribute are defined by the W3C XML Schema `datetime` datatype.           |
|      idx       | This attribute specifies an identifier for this comment that is unique within a list of all comments by this author in this document. An author's first comment in a document has index 1.<br/>Note: Because the index is unique only for the comment author, a document can contain multiple comments with the same index created by different authors.<br/>The possible values for this attribute are defined by the ST_Index simple type (§19.7.3). |

## Open XML SDK Comment Class

The OXML SDK `Comment` class represents the `<cm/>` element defined in the Open XML File Format schema for PresentationML documents. Use the `Comment`
class to manipulate individual `<cm/>` elements in a PresentationML document.

Classes that represent child elements of the `<cm/>` element and that are
therefore commonly associated with the `Comment` class are shown in the following list.

### ExtensionListWithModification Class

The `ExtensionListWithModification` class corresponds to the `<extLst/>`element. The following information from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification introduces the `<extLst/>` element:

This element specifies the extension list with modification ability within which all future extensions of element type `<ext/>` are defined. The extension list along with corresponding future extensions is used to extend the storage capabilities of the PresentationML framework. This allows for various new kinds of data to be stored natively within the framework.

> [!NOTE]
> Using this `extLst` element allows the generating application to store whether this extension property has been modified. end note

### Position Class

The `Position` class corresponds to the
`<pos/>`element. The following information from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification introduces the `<pos/>` element:

This element specifies the positioning information for the placement of
a comment on a slide surface. In LTR versions of the generating
application, this position information should refer to the upper left
point of the comment shape. In RTL versions of the generating
application, this position information should refer to the upper right
point of the comment shape.

[Note: The anchoring point on the slide surface is unaffected by a
right-to-left or left-to-right layout change. That is the anchoring
point remains the same for all language versions. end note]

[Note: Because there is no specified size or formatting for comments,
this UI widget used to display a comment can be any size and thus the
lower right point of the comment shape is determined by how the viewing
application chooses to display comments. end note]

[Example: \<p:pos x="1426" y="660"/\> end example]

### Text class

The `Text` class corresponds to the
`<text/>` element. The following information from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification introduces the `<text/>` element:

This element specifies the content of a comment. This is the text with
which the author has annotated the slide.

The possible values for this element are defined by the W3C XML Schema
`string` datatype.

## Working with the Comment Class

A comment is a text note attached to a slide, with the primary purpose
of enabling readers of a presentation to provide feedback to the
presentation author. Each comment contains an unformatted text string
and information about its author and is attached to a particular
location on a slide. Comments can be visible while editing the
presentation, but do not appear when a slide show is given. The
displaying application decides when to display comments and determines
their visual appearance.

As shown in the Open XML SDK code sample that follows, every instance of
the `Comment` class is associated with an
instance of the <xref:DocumentFormat.OpenXml.Packaging.SlideCommentsPart> class, which represents a
slide comments part, one of the parts of a PresentationML presentation
file package, and a part that is required for each slide in a
presentation file that contains comments. Each `Comment` class instance is also associated with an
instance of the <xref:DocumentFormat.OpenXml.Presentation.CommentAuthor> class, which is in turn
associated with a similarly named presentation part, represented by the
<xref:DocumentFormat.OpenXml.Packaging.CommentAuthorsPart> class. Comment authors
for a presentation are specified in a comment author list, represented
by the <xref:DocumentFormat.OpenXml.Presentation.CommentAuthorList> class, while comments for
each slide are listed in a comments list for that slide, represented by
the <xref:DocumentFormat.OpenXml.Presentation.CommentList> class.

The `Comment` class, which represents the `<cm/>` element, is therefore also associated with other classes that represent the child elements of the `<cm/>` element. Among these classes, as shown in the following code sample, are the `Position` class, which specifies the position of the comment relative to the slide, and the `Text` class, which specifies the text content of the comment.

## Open XML SDK Code Example

The following code segment from the article [How to: Add a comment to a slide in a presentation](how-to-add-a-comment-to-a-slide-in-a-presentation.md) adds a new comments part to an existing slide in a presentation (if the slide does not already contain comments) and creates an instance of an Open XML SDK `Comment` class in the slide comments part. It also adds a comment list to the comments part by creating an instance of the `CommentList` class, if one does not already exist; assigns an ID to the comment; and then adds a comment to the comment list by creating an instance of the `Comment` class, assigning the required attribute values. In addition, it creates instances of the `Position` and `Text` classes associated with the new `Comment` class instance. For the complete code sample, see the aforementioned article.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/add_comment/cs/Program.cs#extsnippet1)]
### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/add_comment/vb/Program.vb#extsnippet1)]

## Generated PresentationML

When the Open XML SDK code in [How to: Add a comment to a slide in a presentation](how-to-add-a-comment-to-a-slide-in-a-presentation.md) is run, including
the segment shown in this article, the following XML is written to a new CommentAuthors.xml part in the existing PresentationML document referenced in the code, assuming that the document contained no comments
or comment authors before the code was run.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cmAuthor id="1"
                  name="userName"
                  initials="userInitials"
                  lastIdx="1"
                  clrIdx="0" />
    </p:cmAuthorLst>
```

In addition, the following XML is written to a new Comments.xml part in
the existing PresentationML document referenced in the code in the
article.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cm authorId="1"
            dt="2010-09-07T16:01:18.5351166-07:00"
            idx="1">
        <p:pos x="100"
               y="200" />
        <p:text>commentText</p:text>
      </p:cm>
    </p:cmLst>
```

## See also

[About the Open XML SDK for Office](../about-the-open-xml-sdk.md)
[How to: Create a Presentation by Providing a File Name](how-to-create-a-presentation-document-by-providing-a-file-name.md)
[How to: Add a comment to a slide in a presentation](how-to-add-a-comment-to-a-slide-in-a-presentation.md)
[How to: Delete all the comments by an author from all the slides in a presentation](how-to-delete-all-the-comments-by-an-author-from-all-the-slides-in-a-presentation.md)  
