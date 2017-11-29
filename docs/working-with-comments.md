---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: d7f0f1d3-bcf9-40b5-aaa4-4a08d862ac8e
title: Working with comments (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Working with comments (Open XML SDK)

This topic discusses the Open XML SDK 2.5 for Office [Comment](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.comment.aspx) class and how it relates to the
Open XML File Format PresentationML schema. For more information about
the overall structure of the parts and elements that make up a
PresentationML document, see [Structure of a PresentationML
Document](structure-of-a-presentationml-document.md).


---------------------------------------------------------------------------------
## Comments in PresentationML 
The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Comments section of the Open XML
PresentationML framework as follows:

A comment is a text note attached to a slide, with the primary purpose
of allowing readers of a presentation to provide feedback to the
presentation author. Each comment contains an unformatted text string
and information about its author, and is attached to a particular
location on a slide. Comments can be visible while editing the
presentation, but do not appear when a slide show is given. The
displaying application decides when to display comments and determines
their visual appearance.

© ISO/IEC29500: 2008.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<cm\> element used
to represent comments in a PresentationML document as follows:

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

© ISO/IEC29500: 2008.

The following table lists the child elements of the \<cm \> element used
when working with comments and the Open XML SDK 2.5 classes that
correspond to them.

**PresentationML Element**|**Open XML SDK 2.5 Class**
---|---
\<extLst\>|[ExtensionListWithModification](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.extensionlistwithmodification.aspx)
\<pos\>|[Position](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.position.aspx)
\<text\>|[Text](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.text.aspx)

The following table from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the attributes of the \<cm\> element.

**Attributes**|**Description**
---|---
authorId|This attribute specifies the author of the comment.It refers to the ID of an author in the comment author list for the document.<br/>The possible values for this attribute are defined by the W3C XML Schema **unsignedInt** datatype.
dt|This attribute specifies the date and time this comment was last modified.<br/>The possible values for this attribute are defined by the W3C XML Schema **datetime** datatype.
idx|This attribute specifies an identifier for this comment that is unique within a list of all comments by this author in this document. An author's first comment in a document has index 1.<br/>[Note: Because the index is unique only for the comment author, a document can contain multiple comments with the same index created by different authors. end note]<br/>The possible values for this attribute are defined by the ST_Index simple type (§19.7.3).

© ISO/IEC29500: 2008.


--------------------------------------------------------------------------------
## Open XML SDK 2.5 Comment Class 
The OXML SDK **Comment** class represents the
\<cm\> element defined in the Open XML File Format schema for
PresentationML documents. Use the **Comment**
class to manipulate individual \<cm\> elements in a PresentationML
document.

Classes that represent child elements of the \<cm\> element and that are
therefore commonly associated with the **Comment** class are shown in the following list.

### ExtensionListWithModification Class

The **ExtensionListWithModification** class
corresponds to the \<extLst\>element. The following information from the
[ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<extLst\> element:

This element specifies the extension list with modification ability
within which all future extensions of element type \<ext\> are defined.
The extension list along with corresponding future extensions is used to
extend the storage capabilities of the PresentationML framework. This
allows for various new kinds of data to be stored natively within the
framework.

> [!NOTE]
> Using this extLst element allows the generating application to
store whether this extension property has been modified. end note

© ISO/IEC29500: 2008.

### Position Class

The **Position** class corresponds to the
\<pos\>element. The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<pos\> element:

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

© ISO/IEC29500: 2008.

### Text class

The **Text** class corresponds to the
\<text\>element. The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<text\> element:

This element specifies the content of a comment. This is the text with
which the author has annotated the slide.

[Example: \<p:text\>Add diagram to clarify.\</p:text\> end example]

The possible values for this element are defined by the W3C XML Schema
**string** datatype.

© ISO/IEC29500: 2008.


--------------------------------------------------------------------------------
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
the **Comment** class is associated with an
instance of the [SlideCommentsPart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.slidecommentspart.aspx) class, which represents a
slide comments part, one of the parts of a PresentationML presentation
file package, and a part that is required for each slide in a
presentation file that contains comments. Each **Comment** class instance is also associated with an
instance of the [CommentAuthor](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.commentauthor.aspx) class, which is in turn
associated with a similarly named presentation part, represented by the
[CommentAuthorsPart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.commentauthorspart.aspx) class. Comment authors
for a presentation are specified in a comment author list, represented
by the [CommentAuthorList](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.commentauthorlist.aspx) class, while comments for
each slide are listed in a comments list for that slide, represented by
the [CommentList](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.commentlist.aspx) class.

The **Comment** class, which represents the
\<cm\> element, is therefore also associated with other classes that
represent the child elements of the \<cm\> element. Among these classes,
as shown in the following code sample, are the **Position** class, which specifies the position of
the comment relative to the slide, and the **Text** class, which specifies the text content of
the comment.


--------------------------------------------------------------------------------
## Open XML SDK Code Example 
The following code segment from the article <span sdata="link">[How to:
Add a comment to a slide in a presentation (Open XML
SDK)](how-to-add-a-comment-to-a-slide-in-a-presentation.md) adds a new
comments part to an existing slide in a presentation (if the slide does
not already contain comments) and creates an instance of an Open XML SDK
2.0 **Comment** class in the slide comments
part. It also adds a comment list to the comments part by creating an
instance of the **CommentList** class, if one
does not already exist; assigns an ID to the comment; and then adds a
comment to the comment list by creating an instance of the **Comment** class, assigning the required attribute
values. In addition, it creates instances of the **Position** and **Text**
classes associated with the new **Comment**
class instance. For the complete code sample, see the aforementioned
article.

```csharp
       // Declare a comments part.
            SlideCommentsPart commentsPart;

            // Verify that there is a comments part in the first slide part.
            if (slidePart1.GetPartsOfType<SlideCommentsPart>().Count() == 0)
            {
                // If not, add a new comments part.
                commentsPart = slidePart1.AddNewPart<SlideCommentsPart>();
            }
            else
            {
                // Else, use the first comments part in the slide part.
                commentsPart = slidePart1.GetPartsOfType<SlideCommentsPart>().First();
            }

            // If the comment list does not exist.
            if (commentsPart.CommentList == null)
            {
                // Add a new comments list.
                commentsPart.CommentList = new CommentList();
            }

            // Get the new comment ID.
            uint commentIdx = author.LastIndex == null ? 1 : author.LastIndex + 1;
            author.LastIndex = commentIdx;

            // Add a new comment.
            Comment comment = commentsPart.CommentList.AppendChild<Comment>(
                new Comment()
                {
                    AuthorId = authorId,
                    Index = commentIdx,
                    DateTime = DateTime.Now
                });

            // Add the position child node to the comment element.
            comment.Append(
                new Position() { X = 100, Y = 200 },
                new Text() { Text = text });
```

```vb
     ' Declare a comments part.
          Dim commentsPart As SlideCommentsPart

          ' Verify that there is a comments part in the first slide part.
          If slidePart1.GetPartsOfType(Of SlideCommentsPart)().Count() = 0 Then

             ' If not, add a new comments part.
             commentsPart = slidePart1.AddNewPart(Of SlideCommentsPart)()
          Else

             ' Else, use the first comments part in the slide part.
             commentsPart = _
              slidePart1.GetPartsOfType(Of SlideCommentsPart)().First()
          End If

          ' If the comment list does not exist.
          If (commentsPart.CommentList Is Nothing) Then

             ' Add a new comments list.
             commentsPart.CommentList = New CommentList()
          End If

          ' Get the new comment ID.
          Dim commentIdx As UInteger
          If author.LastIndex Is Nothing Then
             commentIdx = 1
          Else
             commentIdx = CType(author.LastIndex, UInteger) + 1
          End If

          author.LastIndex = commentIdx

          ' Add a new comment.
          Dim comment As Comment = _
           (commentsPart.CommentList.AppendChild(Of Comment)(New Comment() _
             With {.AuthorId = authorId, .Index = commentIdx, .DateTime = DateTime.Now}))

          ' Add the position child node to the comment element.
          comment.Append(New Position() With _
               {.X = 100, .Y = 200}, New Text() With {.Text = text})

```

---------------------------------------------------------------------------------
## Generated PresentationML 
When the Open XML SDK 2.5 code in <span sdata="link">[How to: Add a
comment to a slide in a presentation (Open XML
SDK)](how-to-add-a-comment-to-a-slide-in-a-presentation.md) is run, including
the segment shown in this article, the following XML is written to a new
CommentAuthors.xml part in the existing PresentationML document
referenced in the code, assuming that the document contained no comments
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

--------------------------------------------------------------------------------
## See also 
#### Concepts

[About the Open XML SDK 2.5 for Office](about-the-open-xml-sdk-2-5.md)  

[How to: Create a Presentation by Providing a File Name](how-to-create-a-presentation-document-by-providing-a-file-name.md)

[How to: Add a comment to a slide in a presentation (Open XML SDK)](how-to-add-a-comment-to-a-slide-in-a-presentation.md)  

[How to: Delete all the comments by an author from all the slides in a presentation (Open XML SDK)](how-to-delete-all-the-comments-by-an-author-from-all-the-slides-in-a-presentatio.md)  
