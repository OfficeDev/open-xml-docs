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
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Delete all the comments by an author from all the slides in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to delete all of the comments by a specific author in a
presentation programmatically.



## Getting a PresentationDocument Object

In the Open XML SDK, the [PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx) class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the
[Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562287.aspx) method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read/write,
specify the value **true** for this parameter
as shown in the following **using** statement.
In this code, the *fileName* parameter is a string that represents the
path for the file from which you want to open the document, and the
author is the user name displayed in the General tab of the PowerPoint
Options.

```csharp
    public static void DeleteCommentsByAuthorInPresentation(string fileName, string author)
    {
        using (PresentationDocument doc = PresentationDocument.Open(fileName, true))
        {
            // Insert other code here.
        }
```

```vb
    Public Shared Sub DeleteCommentsByAuthorInPresentation(ByVal fileName As String, ByVal author As String)

        Using doc As PresentationDocument = PresentationDocument.Open(fileName, True)
            ' Insert other code here.
        End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case *doc*.

[!include[Structure](../includes/presentation/structure.md)]

## The Structure of the Comment Element

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
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
> © ISO/IEC29500: 2008.

The following XML element specifies a single comment attached to a
slide. It contains the text of the comment (**text**), its position on the slide (**pos**), and attributes referring to its author
(**authorId**), date and time (**dt**), and comment index (**idx**).

```xml
    <p:cm authorId="0" dt="2006-08-28T17:26:44.129" idx="1">
        <p:pos x="10" y="10"/>
        <p:text>Add diagram to clarify.</p:text>
    </p:cm>
```

The following table lists the definitions of the members and attributes
of the **cm** (comment) element.

| Member/Attribute | Definition |
|---|---|
| authorId | Refers to the ID of an author in the comment author list for the document. |
| dt | The date and time this comment was last modified. |
| idx | An identifier for this comment that is unique within a list of all comments by this author in this document. An author's first comment in a document has index 1. |
| pos | The positioning information for the placement of a comment on a slide surface. |
| text | Comment's text content. |
| extLst | Specifies the extension list with modification ability within which all future extensions of element type ext are defined. The extension list along with corresponding future extensions is used to extend the storage capabilities of the PresentationML framework. This allows for various new kinds of data to be stored natively within the framework. |

The following XML schema example defines the members of the **cm** element in addition to the required and
optional attributes.

```xml
    <complexType name="CT_Comment">
       <sequence>
           <element name="pos" type="a:CT_Point2D" minOccurs="1" maxOccurs="1"/>
           <element name="text" type="xsd:string" minOccurs="1" maxOccurs="1"/>
           <element name="extLst" type="CT_ExtensionListModify" minOccurs="0" maxOccurs="1"/>
       </sequence>
       <attribute name="authorId" type="xsd:unsignedInt" use="required"/>
       <attribute name="dt" type="xsd:dateTime" use="optional"/>
       <attribute name="idx" type="ST_Index" use="required"/>
    </complexType>
```

## How the Sample Code Works

After opening the presentation document for read/write access and
instantiating the **PresentationDocument**
class, the code gets the specified comment author from the list of
comment authors.

```csharp
    // Get the specifed comment author.
    IEnumerable<CommentAuthor> commentAuthors = 
        doc.PresentationPart.CommentAuthorsPart.CommentAuthorList.Elements<CommentAuthor>()
        .Where(e => e.Name.Value.Equals(author));
```

```vb
    ' Get the specifed comment author.
    Dim commentAuthors As IEnumerable(Of CommentAuthor) = _
        doc.PresentationPart.CommentAuthorsPart.CommentAuthorList.Elements _
       (Of CommentAuthor)().Where(Function(e) e.Name.Value.Equals(author))
```

By iterating through the matching authors and all the slides in the
presentation the code gets all the slide parts, and the comments part of
each slide part. It then gets the list of comments by the specified
author and deletes each one. It also verifies that the comment part has
no existing comment, in which case it deletes that part. It also deletes
the comment author from the comment authors part.

```csharp
    // Iterate through all the matching authors.
    foreach (CommentAuthor commentAuthor in commentAuthors)
    {
        UInt32Value authorId = commentAuthor.Id;

        // Iterate through all the slides and get the slide parts.
        foreach (SlidePart slide in doc.PresentationPart.SlideParts)
        {
            SlideCommentsPart slideCommentsPart = slide.SlideCommentsPart;
            // Get the list of comments.
            if (slideCommentsPart != null && slide.SlideCommentsPart.CommentList != null)
            {
                IEnumerable<Comment> commentList = 
                    slideCommentsPart.CommentList.Elements<Comment>().Where(e => e.AuthorId == authorId.Value);
                List<Comment> comments = new List<Comment>();
                comments = commentList.ToList<Comment>();

                foreach (Comment comm in comments)
                {
                    // Delete all the comments by the specified author.
                    slideCommentsPart.CommentList.RemoveChild<Comment>(comm);
                }

                // If the commentPart has no existing comment.
                if (slideCommentsPart.CommentList.ChildElements.Count == 0)
                    // Delete this part.
                    slide.DeletePart(slideCommentsPart);
            }
        }
        // Delete the comment author from the comment authors part.
        doc.PresentationPart.CommentAuthorsPart.CommentAuthorList.RemoveChild<CommentAuthor>(commentAuthor);
    }
```

```vb
    'Iterate through all the matching authors
    For Each commentAuthor In commentAuthors

        Dim authorId = commentAuthor.Id

        ' Iterate through all the slides and get the slide parts.
        For Each slide In doc.PresentationPart.GetPartsOfType(Of SlidePart)()

            ' Get the slide comments part of each slide.
            For Each slideCommentsPart In slide.GetPartsOfType(Of SlideCommentsPart)()

                ' Delete all the comments by the specified author.
                Dim commentList = slideCommentsPart.CommentList.Elements(Of Comment)(). _
                    Where(Function(e) e.AuthorId.Value.Equals(authorId.Value))

                Dim comments As List(Of Comment) = commentList.ToList()

                For Each comm As Comment In comments
                    slideCommentsPart.CommentList.RemoveChild(Of Comment)(comm)
                Next

            Next

        Next

        ' Delete the comment author from the comment authors part.
        doc.PresentationPart.CommentAuthorsPart.CommentAuthorList.RemoveChild(Of CommentAuthor)(commentAuthor)

    Next
```

## Sample Code

The following method takes as parameters the source presentation file
name and path and the name of the comment author whose comments are to
be deleted. It finds all the comments by the specified author in the
presentation and deletes them. It then deletes the comment author from
the list of comment authors.

You can use the following example to call the
*DeleteCommentsByAuthorInPresentation* method to remove the comments of
the specified author from the presentation file, *myppt5.pptx*.

```csharp
    string fileName = @"C:\Users\Public\Documents\myppt5.pptx";
    string author = "Katie Jordan";
    DeleteCommentsByAuthorInPresentation(fileName, author);
```

```vb
    Dim fileName As String = "C:\Users\Public\Documents\myppt5.pptx"
    Dim author As String = "Katie Jordan"
    DeleteCommentsByAuthorInPresentation(fileName, author)
```
> [!NOTE]
> To get the exact author's name, open the presentation file and click the **File** menu item, and then click **Options**. The **PowerPoint Options** window opens and the content of the **General** tab is displayed. The author's name must match the **User name** in this tab.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/delete_all_the_comments_by_an_author_from_all_the_slides_a_presentatio/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/delete_all_the_comments_by_an_author_from_all_the_slides_a_presentatio/vb/Program.vb)]

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
