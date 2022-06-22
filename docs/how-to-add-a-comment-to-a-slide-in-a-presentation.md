---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 403abe97-7ab2-40ba-92c0-d6312a6d10c8
title: 'How to: Add a comment to a slide in a presentation (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---

# Add a comment to a slide in a presentation (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to add a comment to the first slide in a presentation
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Presentation;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System
    Imports System.Linq
    Imports DocumentFormat.OpenXml.Presentation
    Imports DocumentFormat.OpenXml.Packaging
```

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
In this code, the *file* parameter is a string that represents the path
for the file from which you want to open the document.

```csharp
    using (PresentationDocument doc = PresentationDocument.Open(file, true))
    {
        // Insert other code here.
    }
```

```vb
    Using doc As PresentationDocument = PresentationDocument.Open(file, True)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case *doc*.


## Basic Presentation Document Structure 

The basic document structure of a **PresentationML** document consists of a number of
parts, among which is the main part that contains the presentation
definition. The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the overall form of a **PresentationML** package.

> The main part of a **PresentationML** package
> starts with a presentation root element. That element contains a
> presentation, which, in turn, refers to a **slide** list, a *slide master* list, a *notes
> master* list, and a *handout master* list. The slide list refers to
> all of the slides in the presentation; the slide master list refers to
> the entire slide masters used in the presentation; the notes master
> contains information about the formatting of notes pages; and the
> handout master describes how a handout looks.
> 
> A *handout* is a printed set of slides that can be provided to an
> *audience*.
> 
> As well as text and graphics, each slide can contain *comments* and
> *notes*, can have a *layout*, and can be part of one or more *custom
> presentations*. A comment is an annotation intended for the person
> maintaining the presentation slide deck. A note is a reminder or piece
> of text intended for the presenter or the audience.
> 
> Other features that a **PresentationML**
> document can include the following: *animation*, *audio*, *video*, and
> *transitions* between slides.
> 
> A **PresentationML** document is not stored
> as one large body in a single part. Instead, the elements that
> implement certain groupings of functionality are stored in separate
> parts. For example, all comments in a document are stored in one
> comment part while each slide has its own part.
> 
> © ISO/IEC29500: 2008.

The following XML code example represents a presentation that contains
two slides denoted by the IDs 267 and 256.

```xml
    <p:presentation xmlns:p="…" … > 
       <p:sldMasterIdLst>
          <p:sldMasterId
             xmlns:rel="https://…/relationships" rel:id="rId1"/>
       </p:sldMasterIdLst>
       <p:notesMasterIdLst>
          <p:notesMasterId
             xmlns:rel="https://…/relationships" rel:id="rId4"/>
       </p:notesMasterIdLst>
       <p:handoutMasterIdLst>
          <p:handoutMasterId
             xmlns:rel="https://…/relationships" rel:id="rId5"/>
       </p:handoutMasterIdLst>
       <p:sldIdLst>
          <p:sldId id="267"
             xmlns:rel="https://…/relationships" rel:id="rId2"/>
          <p:sldId id="256"
             xmlns:rel="https://…/relationships" rel:id="rId3"/>
       </p:sldIdLst>
           <p:sldSz cx="9144000" cy="6858000"/>
       <p:notesSz cx="6858000" cy="9144000"/>
    </p:presentation>
```

Using the Open XML SDK 2.5, you can create document structure and
content using strongly-typed classes that correspond to PresentationML
elements. You can find these classes in the [DocumentFormat.OpenXml.Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.aspx)
namespace. The following table lists the class names of the classes that
correspond to the **sld**, **sldLayout**, **sldMaster**, and **notesMaster** elements.

| PresentationML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| sld | [Slide](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slide.aspx) | Presentation Slide. It is the root element of SlidePart. |
| sldLayout | [SlideLayout](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slidelayout.aspx) | Slide Layout. It is the root element of SlideLayoutPart. |
| sldMaster | [SlideMaster](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slidemaster.aspx) | Slide Master. It is the root element of SlideMasterPart. |
| notesMaster | [NotesMaster](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.notesmaster.aspx) | Notes Master (or handoutMaster). It is the root element of NotesMasterPart. |


## The Structure of the Comment Element

A comment is a text note attached to a slide, with the primary purpose
of enabling readers of a presentation to provide feedback to the
presentation author. Each comment contains an unformatted text string
and information about its author, and is attached to a particular
location on a slide. Comments can be visible while editing the
presentation, but do not appear when a slide show is given. The
displaying application decides when to display comments and determines
their visual appearance.

The following XML element specifies a single comment attached to a
slide. It contains the text of the comment (**text**), its position on the slide (**pos**), and attributes referring to its author
(**authorId**), date and time (**dt**), and comment index (**idx**).

```xml
    <p:cm authorId="0" dt="2006-08-28T17:26:44.129" idx="1">
        <p:pos x="10" y="10"/>
        <p:text>Add diagram to clarify.</p:text>
    </p:cm>
```

The following table contains the definitions of the members and
attributes of the **cm** (comment) element.

| Member/Attribute | Definition |
|---|---|
| authorId | Refers to the ID of an author in the comment author list for the document. |
| dt | The date and time this comment was last modified. |
| idx | An identifier for this comment that is unique within a list of all comments by this author in this document. An author's first comment in a document has index 1. |
| pos | The positioning information for the placement of a comment on a slide surface. |
| text | Comment text. |
| extLst | Specifies the extension list with modification ability within which all future extensions of element type ext are defined. The extension list along with corresponding future extensions is used to extend the storage capabilities of the PresentationML framework. This allows for various new kinds of data to be stored natively within the framework. |


The following XML schema code example defines the members of the **cm** element in addition to the required and
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

The sample code opens the presentation document in the **using** statement. Then it instantiates the **CommentAuthorsPart**, and verifies that there is an
existing comment authors part. If there is not, it adds one.

```csharp
    // Declare a CommentAuthorsPart object.
    CommentAuthorsPart authorsPart;

    // Verify that there is an existing comment authors part. 
    if (doc.PresentationPart.CommentAuthorsPart == null)
    {
        // If not, add a new one.
        authorsPart = doc.PresentationPart.AddNewPart<CommentAuthorsPart>();
    }
    else
    {
        authorsPart = doc.PresentationPart.CommentAuthorsPart;
    }
```

```vb
    ' Declare a CommentAuthorsPart object.
    Dim authorsPart As CommentAuthorsPart

    ' Verify that there is an existing comment authors part. 
    If doc.PresentationPart.CommentAuthorsPart Is Nothing Then
        ' If not, add a new one.
        authorsPart = doc.PresentationPart.AddNewPart(Of CommentAuthorsPart)()
    Else
        authorsPart = doc.PresentationPart.CommentAuthorsPart
    End If
```

The code determines whether there is an existing comment author list in
the comment-authors part; if not, it adds one. It also verifies that the
author that is passed in is on the list of existing comment authors; if
so, it assigns the existing author ID. If not, it adds a new author to
the list of comment authors and assigns an author ID and the parameter
values.

```csharp
    // Verify that there is a comment author list in the comment authors part.
    if (authorsPart.CommentAuthorList == null)
    {
        // If not, add a new one.
        authorsPart.CommentAuthorList = new CommentAuthorList();
    }

    // Declare a new author ID.
    uint authorId = 0;
    CommentAuthor author = null;

    // If there are existing child elements in the comment authors list...
    if (authorsPart.CommentAuthorList.HasChildren)
    {
        // Verify that the author passed in is on the list.
        var authors = authorsPart.CommentAuthorList.Elements<CommentAuthor>().Where(a => a.Name == name && a.Initials == initials);

        // If so...
        if (authors.Any())
        {
            // Assign the new comment author the existing author ID.
            author = authors.First();
            authorId = author.Id;
        }

        // If not...
        if (author == null)
        {
            // Assign the author passed in a new ID                        
            authorId = authorsPart.CommentAuthorList.Elements<CommentAuthor>().Select(a => a.Id.Value).Max();
        }
    }

    // If there are no existing child elements in the comment authors list.
    if (author == null)
    {
        authorId++;

        // Add a new child element(comment author) to the comment author list.
        author = authorsPart.CommentAuthorList.AppendChild<CommentAuthor>
            (new CommentAuthor()
            {
                Id = authorId,
                Name = name,
                Initials = initials,
                ColorIndex = 0
            });
    }
```

```vb
    ' Verify that there is a comment author list in the comment authors part.
    If authorsPart.CommentAuthorList Is Nothing Then
        ' If not, add a new one.
        authorsPart.CommentAuthorList = New CommentAuthorList()
    End If

    ' Declare a new author ID.
    Dim authorId As UInteger = 0
    Dim author As CommentAuthor = Nothing

    ' If there are existing child elements in the comment authors list...
    If authorsPart.CommentAuthorList.HasChildren Then
        ' Verify that the author passed in is on the list.
        Dim authors = authorsPart.CommentAuthorList.Elements(Of CommentAuthor)().Where(Function(a) a.Name = name AndAlso a.Initials = initials)

        ' If so...
        If authors.Any() Then
            ' Assign the new comment author the existing author ID.
            author = authors.First()
            authorId = author.Id
        End If

        ' If not...
        If author Is Nothing Then
            ' Assign the author passed in a new ID                        
            authorId = authorsPart.CommentAuthorList.Elements(Of CommentAuthor)().Select(Function(a) a.Id.Value).Max()
        End If
    End If

    ' If there are no existing child elements in the comment authors list.
    If author Is Nothing Then
        authorId += 1

        ' Add a new child element(comment author) to the comment author list.
        author = authorsPart.CommentAuthorList.AppendChild(Of CommentAuthor) (New CommentAuthor() With {.Id = authorId, .Name = name, .Initials = initials, .ColorIndex = 0})
    End If
```

In the following code segment, the code gets the first slide in the
presentation by calling the *GetFirstSlide* method. Then it verifies
that there is a comments part in the slide; if not, it adds one. It also
verifies that a comments list exists in the comments part; if not, it
creates one.

```csharp
    // Get the first slide, using the GetFirstSlide method.
    SlidePart slidePart1 = GetFirstSlide(doc);

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
        commentsPart = slidePart1.SlideCommentsPart;
    }

    // If the comment list does not exist.
    if (commentsPart.CommentList == null)
    {
        // Add a new comments list.
        commentsPart.CommentList = new CommentList();
    }
```

```vb
    ' Get the first slide, using the GetFirstSlide method.
    Dim slidePart1 As SlidePart = GetFirstSlide(doc)

    ' Declare a comments part.
    Dim commentsPart As SlideCommentsPart

    ' Verify that there is a comments part in the first slide part.
    If slidePart1.GetPartsOfType(Of SlideCommentsPart)().Count() = 0 Then
        ' If not, add a new comments part.
        commentsPart = slidePart1.AddNewPart(Of SlideCommentsPart)()
    Else
        ' Else, use the first comments part in the slide part.
        commentsPart = slidePart1.SlideCommentsPart
    End If

    ' If the comment list does not exist.
    If commentsPart.CommentList Is Nothing Then
        ' Add a new comments list.
        commentsPart.CommentList = New CommentList()
    End If
```

The code then gets the ID of the new comment, and adds the specified
comment, containing the specified text, at the specified position. Then
it saves the comment authors part and the comments part.

```csharp
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

    // Save the comment authors part.
    authorsPart.CommentAuthorList.Save();

    // Save the comments part.
    commentsPart.CommentList.Save();
```

```vb
    ' Get the new comment ID.
    Dim commentIdx As UInteger = If(author.LastIndex Is Nothing, 1, author.LastIndex + 1)
    author.LastIndex = commentIdx

    ' Add a new comment.
    Dim comment As Comment = commentsPart.CommentList.AppendChild(Of Comment)(New Comment() With {.AuthorId = authorId, .Index = commentIdx, .DateTime = Date.Now})

    ' Add the position child node to the comment element.
    comment.Append(New Position() With {.X = 100, .Y = 200}, New Text() With {.Text = text})

    ' Save the comment authors part.
    authorsPart.CommentAuthorList.Save()

    ' Save the comments part.
    commentsPart.CommentList.Save()
```

## Sample Code

The **AddCommentToPresentation** method can be
used to add a comment to a slide. The method takes as parameters the
source presentation file name and path, the initials and name of the
comment author, and the text of the comment to be added. It adds an
author to the list of comment authors and then adds the specified
comment text at the specified coordinates in the first slide in the
presentation.

The second method, **GetFirstSlide**, is used
to get the first slide in the presentation. It takes the **PresentationDocument** object passed in, gets its
presentation part, and then gets the ID of the first slide in its slide
list. It then gets the relationship ID of the slide, gets the slide part
from the relationship ID, and returns the slide part to the calling
method.

The following code example shows a call to the **AddCommentToPresentation** which adds the specified
comment string to the first slide in the presentation file Myppt1.pptx.

```csharp
    AddCommentToPresentation(@"C:\Users\Public\Documents\Myppt1.pptx", 
    "Katie Jordan", "KJ", 
    "This is my programmatically added comment.");
```

```vb
    AddCommentToPresentation("C:\Users\Public\Documents\Myppt1.pptx", _
    "Katie Jordan", "KJ", _
    "This is my programmatically added comment.")
```

> [!NOTE]
> To get the exact author name and initials, open the presentation file and click the **File** menu item, and then click **Options**. The **PowerPointOptions** window opens and the content of the **General** tab is displayed. The author name and initials must match the **User name** and **Initials** in this tab.


```csharp
    // Adds a comment to the first slide of the presentation document.
    // The presentation document must contain at least one slide.
    public static void AddCommentToPresentation(string file, string initials, string name, string text)
    {
        using (PresentationDocument doc = PresentationDocument.Open(file, true))
        {

            // Declare a CommentAuthorsPart object.
            CommentAuthorsPart authorsPart;

            // Verify that there is an existing comment authors part. 
            if (doc.PresentationPart.CommentAuthorsPart == null)
            {
                // If not, add a new one.
                authorsPart = doc.PresentationPart.AddNewPart<CommentAuthorsPart>();
            }
            else
            {
                authorsPart = doc.PresentationPart.CommentAuthorsPart;
            }

            // Verify that there is a comment author list in the comment authors part.
            if (authorsPart.CommentAuthorList == null)
            {
                // If not, add a new one.
                authorsPart.CommentAuthorList = new CommentAuthorList();
            }

            // Declare a new author ID.
            uint authorId = 0;
            CommentAuthor author = null;

            // If there are existing child elements in the comment authors list...
            if (authorsPart.CommentAuthorList.HasChildren)
            {
                // Verify that the author passed in is on the list.
                var authors = authorsPart.CommentAuthorList.Elements<CommentAuthor>().Where(a => a.Name == name && a.Initials == initials);

                // If so...
                if (authors.Any())
                {
                    // Assign the new comment author the existing author ID.
                    author = authors.First();
                    authorId = author.Id;
                }

                // If not...
                if (author == null)
                {
                    // Assign the author passed in a new ID                        
                    authorId = authorsPart.CommentAuthorList.Elements<CommentAuthor>().Select(a => a.Id.Value).Max();
                }
            }

            // If there are no existing child elements in the comment authors list.
            if (author == null)
            {

                authorId++;

                // Add a new child element(comment author) to the comment author list.
                author = authorsPart.CommentAuthorList.AppendChild<CommentAuthor>
                    (new CommentAuthor()
                    {
                        Id = authorId,
                        Name = name,
                        Initials = initials,
                        ColorIndex = 0
                    });
            }

            // Get the first slide, using the GetFirstSlide method.
            SlidePart slidePart1 = GetFirstSlide(doc);

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

            // Save the comment authors part.
            authorsPart.CommentAuthorList.Save();

            // Save the comments part.
            commentsPart.CommentList.Save();
        }
     }
    // Get the slide part of the first slide in the presentation document.
    public static SlidePart GetFirstSlide(PresentationDocument presentationDocument)
    {
        // Get relationship ID of the first slide
        PresentationPart part = presentationDocument.PresentationPart;
        SlideId slideId = part.Presentation.SlideIdList.GetFirstChild<SlideId>();
        string relId = slideId.RelationshipId;

        // Get the slide part by the relationship ID.
        SlidePart slidePart = (SlidePart)part.GetPartById(relId);

        return slidePart;
    }
```

```vb
    ' Adds a comment to the first slide of the presentation document.
    ' The presentation document must contain at least one slide.
    Public Sub AddCommentToPresentation(ByVal file As String, _
                ByVal initials As String, _
                ByVal name As String, _
                ByVal text As String)

       Dim doc As PresentationDocument = _
          PresentationDocument.Open(file, True)

       Using (doc)

          ' Declare a CommentAuthorsPart object.
          Dim authorsPart As CommentAuthorsPart

          ' Verify that there is an existing comment authors part.
          If (doc.PresentationPart.CommentAuthorsPart Is Nothing) Then

             ' If not, add a new one.
             authorsPart = doc.PresentationPart.AddNewPart(Of CommentAuthorsPart)()
          Else
             authorsPart = doc.PresentationPart.CommentAuthorsPart
          End If

          ' Verify that there is a comment author list in the comment authors part.
          If (authorsPart.CommentAuthorList Is Nothing) Then

             ' If not, add a new one.
             authorsPart.CommentAuthorList = New CommentAuthorList()
          End If

          ' Declare a new author ID.
          Dim authorId As UInteger = 0
          Dim author As CommentAuthor = Nothing

          ' If there are existing child elements in the comment authors list.
          If authorsPart.CommentAuthorList.HasChildren = True Then

             ' Verify that the author passed in is on the list.
             Dim authors = authorsPart.CommentAuthorList.Elements(Of CommentAuthor)().Where _
              (Function(a) a.Name = name AndAlso a.Initials = initials)

             ' If so...
             If (authors.Any()) Then

                ' Assign the new comment author the existing ID.
                author = authors.First()
                authorId = author.Id
             End If

             ' If not...
             If (author Is Nothing) Then

                ' Assign the author passed in a new ID.
                authorId = _
                authorsPart.CommentAuthorList.Elements(Of CommentAuthor)().Select(Function(a) a.Id.Value).Max()
             End If

          End If

          ' If there are no existing child elements in the comment authors list.
          If (author Is Nothing) Then

             authorId = authorId + 1

             ' Add a new child element (comment author) to the comment author list.
             author = (authorsPart.CommentAuthorList.AppendChild(Of CommentAuthor) _
              (New CommentAuthor() With {.Id = authorId, _
               .Name = name, _
               .Initials = initials, _
               .ColorIndex = 0}))
          End If

          ' Get the first slide, using the GetFirstSlide() method.
          Dim slidePart1 As SlidePart
          slidePart1 = GetFirstSlide(doc)

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


          ' Save comment authors part.
          authorsPart.CommentAuthorList.Save()

          ' Save comments part.
          commentsPart.CommentList.Save()

       End Using

    End Sub

    ' Get the slide part of the first slide in the presentation document.
    Public Function GetFirstSlide(ByVal presentationDocument As PresentationDocument) As SlidePart
       ' Get relationship ID of the first slide
       Dim part As PresentationPart = presentationDocument.PresentationPart
       Dim slideId As SlideId = part.Presentation.SlideIdList.GetFirstChild(Of SlideId)()
       Dim relId As String = slideId.RelationshipId

       ' Get the slide part by the relationship ID.
       Dim slidePart As SlidePart = DirectCast(part.GetPartById(relId), SlidePart)

       Return slidePart
    End Function
    End Module
```

## See also



- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)


