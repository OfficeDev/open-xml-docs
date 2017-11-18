---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b7d5d1fd-dcdf-4f88-9d57-884562c8144f
title: 'How to: Get the titles of all the slides in a presentation (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Get the titles of all the slides in a presentation (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to get the titles of all slides in a presentation
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;
    using D = DocumentFormat.OpenXml.Drawing;
```

```vb
    Imports System
    Imports System.Collections.Generic
    Imports System.Linq
    Imports System.Text
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Presentation
    Imports D = DocumentFormat.OpenXml.Drawing
```

---------------------------------------------------------------------------------

In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.PresentationDocument"><span
class="nolink">PresentationDocument</span></span> class represents a
presentation document package. To work with a presentation document,
first create an instance of the <span
class="keyword">PresentationDocument</span> class, and then work with
that instance. To create the class instance from the document call the
<span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(System.String,System.Boolean)"><span
class="nolink">PresentationDocument.Open(String, Boolean)</span></span>
method that uses a file path, and a Boolean value as the second
parameter to specify whether a document is editable. To open a document
for read-only, specify the value **false** for
this parameter as shown in the following <span
class="keyword">using</span> statement. In this code, the <span
class="keyword">presentationFile</span> parameter is a string that
represents the path for the file from which you want to open the
document.

```csharp
    // Open the presentation as read-only.
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
    {
        // Insert other code here.
    }
```

```vb
    ' Open the presentation as read-only.
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case
*presentationDocument*.


---------------------------------------------------------------------------------

The basic document structure of a <span
class="keyword">PresentationML</span> document consists of the main part
that contains the presentation definition. The following text from the
[ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337)
specification introduces the overall form of a <span
class="keyword">PresentationML</span> package.

> A **PresentationML** package's main part
> starts with a presentation root element. That element contains a
> presentation, which, in turn, refers to a <span
> class="keyword">slide</span> list, a *slide master* list, a *notes
> master* list, and a *handout master* list. The slide list refers to
> all of the slides in the presentation; the slide master list refers to
> the entire slide masters used in the presentation; the notes master
> contains information about the formatting of notes pages; and the
> handout master describes how a handout looks.

> A *handout* is a printed set of slides that can be provided to an
> *audience* for future reference.

> As well as text and graphics, each slide can contain *comments* and
> *notes*, can have a *layout*, and can be part of one or more *custom
> presentations*. (A comment is an annotation intended for the person
> maintaining the presentation slide deck. A note is a reminder or piece
> of text intended for the presenter or the audience.)

> Other features that a **PresentationML**
> document can include the following: *animation*, *audio*, *video*, and
> *transitions* between slides.

> A **PresentationML** document is not stored
> as one large body in a single part. Instead, the elements that
> implement certain groupings of functionality are stored in separate
> parts. For example, all comments in a document are stored in one
> comment part while each slide has its own part.

> © ISO/IEC29500: 2008.

This following XML code segment represents a presentation that contains
two slides denoted by the Id's 267 and 256.

```xml
    <p:presentation xmlns:p="…" … > 
       <p:sldMasterIdLst>
          <p:sldMasterId
             xmlns:rel="http://…/relationships" rel:id="rId1"/>
       </p:sldMasterIdLst>
       <p:notesMasterIdLst>
          <p:notesMasterId
             xmlns:rel="http://…/relationships" rel:id="rId4"/>
       </p:notesMasterIdLst>
       <p:handoutMasterIdLst>
          <p:handoutMasterId
             xmlns:rel="http://…/relationships" rel:id="rId5"/>
       </p:handoutMasterIdLst>
       <p:sldIdLst>
          <p:sldId id="267"
             xmlns:rel="http://…/relationships" rel:id="rId2"/>
          <p:sldId id="256"
             xmlns:rel="http://…/relationships" rel:id="rId3"/>
       </p:sldIdLst>
           <p:sldSz cx="9144000" cy="6858000"/>
       <p:notesSz cx="6858000" cy="9144000"/>
    </p:presentation>
```

Using the Open XML SDK 2.5, you can create document structure and
content using strongly-typed classes that correspond to PresentationML
elements. You can find these classes in the <span sdata="cer"
target="N:DocumentFormat.OpenXml.Presentation"><span
class="nolink">DocumentFormat.OpenXml.Presentation</span></span>
namespace. The following table lists the class names of the classes that
correspond to the **sld**, <span
class="keyword">sldLayout</span>, <span
class="keyword">sldMaster</span>, and <span
class="keyword">notesMaster</span> elements:

| PresentationML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| sld | Slide | Presentation Slide. It is the root element of SlidePart. |
| sldLayout | SlideLayout | Slide Layout. It is the root element of SlideLayoutPart. |
| sldMaster | SlideMaster | Slide Master. It is the root element of SlideMasterPart. |
| notesMaster | NotesMaster | Notes Master (or handoutMaster). It is the root element of NotesMasterPart. |


--------------------------------------------------------------------------------

The sample code consists of two overloads of the method <span
class="keyword">GetSlideTitles</span>. In the first overloaded method,
the presentation file is opened in the <span
class="keyword">using</span> statement. Then it passes the <span
class="keyword">PresentationDocument</span> object to the second
overloaded **GetSlideTitles** method, which
returns a list that represents the titles of all the slides in the
presentation.

```csharp
    // Get a list of the titles of all the slides in the presentation.
    public static IList<string> GetSlideTitles(string presentationFile)
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument =
            PresentationDocument.Open(presentationFile, false))
        {
            return GetSlideTitles(presentationDocument);
        }
    }
```

```vb
    ' Get a list of the titles of all the slides in the presentation.
    Public Shared Function GetSlideTitles(ByVal presentationFile As String) As IList(Of String)
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            Return GetSlideTitles(presentationDocument)
        End Using
    End Function
```

The second overloaded **GetSlideTitles** method
is used to get a list of slide titles. It takes the <span
class="keyword">PresentationDocument</span> object passed in, iterates
through its slides, and gets the slide IDs of all the slides in the
presentation. For each slide ID, it gets a slide part to pass to the
**GetSlideTitle** method. It returns to the
first **GetSlideTitles** method a list of
strings that it assembles from the titles, each of which represents a
slide title.

```csharp
    // Get a list of the titles of all the slides in the presentation.
    public static IList<string> GetSlideTitles(PresentationDocument presentationDocument)
    {
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Get a PresentationPart object from the PresentationDocument object.
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        if (presentationPart != null &&
            presentationPart.Presentation != null)
        {
            // Get a Presentation object from the PresentationPart object.
            Presentation presentation = presentationPart.Presentation;

            if (presentation.SlideIdList != null)
            {
                List<string> titlesList = new List<string>();

                // Get the title of each slide in the slide order.
                foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                {
                    SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;

                    // Get the slide title.
                    string title = GetSlideTitle(slidePart);

                    // An empty title can also be added.
                    titlesList.Add(title);
                }

                return titlesList;
            }

        }

        return null;
    }
```

```vb
    ' Get a list of the titles of all the slides in the presentation.
    Public Shared Function GetSlideTitles(ByVal presentationDocument As PresentationDocument) As IList(Of String)
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Get a PresentationPart object from the PresentationDocument object.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        If presentationPart IsNot Nothing AndAlso presentationPart.Presentation IsNot Nothing Then
            ' Get a Presentation object from the PresentationPart object.
            Dim presentation As Presentation = presentationPart.Presentation

            If presentation.SlideIdList IsNot Nothing Then
                Dim titlesList As New List(Of String)()

                ' Get the title of each slide in the slide order.
                For Each slideId In presentation.SlideIdList.Elements(Of SlideId)()
                    Dim slidePart As SlidePart = TryCast(presentationPart.GetPartById(slideId.RelationshipId), SlidePart)

                    ' Get the slide title.
                    Dim title As String = GetSlideTitle(slidePart)

                    ' An empty title can also be added.
                    titlesList.Add(title)
                Next slideId

                Return titlesList
            End If

        End If

        Return Nothing
    End Function
```

The method **GetSlideTitle** is used to get the
title of each slide. It takes the slide part passed in and returns to
the second overloaded GetSlideTitles method a string that represents the
title of the slide.

```csharp
    // Get the title string of the slide.
    public static string GetSlideTitle(SlidePart slidePart)
    {
        if (slidePart == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Declare a paragraph separator.
        string paragraphSeparator = null;

        if (slidePart.Slide != null)
        {
            // Find all the title shapes.
            var shapes = from shape in slidePart.Slide.Descendants<Shape>()
                         where IsTitleShape(shape)
                         select shape;

            StringBuilder paragraphText = new StringBuilder();

            foreach (var shape in shapes)
            {
                // Get the text in each paragraph in this shape.
                foreach (var paragraph in shape.TextBody.Descendants<D.Paragraph>())
                {
                    // Add a line break.
                    paragraphText.Append(paragraphSeparator);

                    foreach (var text in paragraph.Descendants<D.Text>())
                    {
                        paragraphText.Append(text.Text);
                    }

                    paragraphSeparator = "\n";
                }
            }

            return paragraphText.ToString();
        }

        return string.Empty;
    }
```

```vb
    ' Get the title string of the slide.
    Public Shared Function GetSlideTitle(ByVal slidePart As SlidePart) As String
        If slidePart Is Nothing Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Declare a paragraph separator.
        Dim paragraphSeparator As String = Nothing

        If slidePart.Slide IsNot Nothing Then
            ' Find all the title shapes.
            Dim shapes = From shape In slidePart.Slide.Descendants(Of Shape)()
                         Where IsTitleShape(shape)
                         Select shape

            Dim paragraphText As New StringBuilder()

            For Each shape In shapes
                ' Get the text in each paragraph in this shape.
                For Each paragraph In shape.TextBody.Descendants(Of D.Paragraph)()
                    ' Add a line break.
                    paragraphText.Append(paragraphSeparator)

                    For Each text In paragraph.Descendants(Of D.Text)()
                        paragraphText.Append(text.Text)
                    Next text

                    paragraphSeparator = vbLf
                Next paragraph
            Next shape

            Return paragraphText.ToString()
        End If

        Return String.Empty
    End Function
```

The Boolean method **IsTitleShape** is called
from within the method **GetSlideTitle** to
determine whether the shape is a title shape. It takes the slide part
passed in and returns **true** if the shape is
a title shape; otherwise, it returns **false**.

```csharp
    // Determines whether the shape is a title shape.
    private static bool IsTitleShape(Shape shape)
    {
        var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
        if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
        {
            switch ((PlaceholderValues)placeholderShape.Type)
            {
                // Any title shape.
                case PlaceholderValues.Title:

                // A centered title.
                case PlaceholderValues.CenteredTitle:
                    return true;

                default:
                    return false;
            }
        }
        return false;
    }
```

```vb
    ' Determines whether the shape is a title shape.
    Private Shared Function IsTitleShape(ByVal shape As Shape) As Boolean
        Dim placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild(Of PlaceholderShape)()
        If placeholderShape IsNot Nothing AndAlso placeholderShape.Type IsNot Nothing AndAlso placeholderShape.Type.HasValue Then
            Select Case CType(placeholderShape.Type, PlaceholderValues)
                ' Any title shape.
                Case PlaceholderValues.Title, PlaceholderValues.CenteredTitle

                ' A centered title.
                    Return True

                Case Else
                    Return False
            End Select
        End If
        Return False
    End Function
```

--------------------------------------------------------------------------------

The following is the complete sample code that you can use to get the
titles of the slides in a presentation file. For example you can use the
following **foreach** statement in your program
to return all the titles in the presentation file, "Myppt9.pptx."

```csharp
    foreach (string s in GetSlideTitles(@"C:\Users\Public\Documents\Myppt9.pptx"))
       Console.WriteLine(s);
```

```vb
    For Each s As String In GetSlideTitles("C:\Users\Public\Documents\Myppt9.pptx")
       Console.WriteLine(s)
    Next
```

The result would be a list of the strings that represent the titles in
the presentation, each on a separate line.

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Get a list of the titles of all the slides in the presentation.
    public static IList<string> GetSlideTitles(string presentationFile)
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument =
            PresentationDocument.Open(presentationFile, false))
        {
            return GetSlideTitles(presentationDocument);
        }
    }

    // Get a list of the titles of all the slides in the presentation.
    public static IList<string> GetSlideTitles(PresentationDocument presentationDocument)
    {
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Get a PresentationPart object from the PresentationDocument object.
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        if (presentationPart != null &&
            presentationPart.Presentation != null)
        {
            // Get a Presentation object from the PresentationPart object.
            Presentation presentation = presentationPart.Presentation;

            if (presentation.SlideIdList != null)
            {
                List<string> titlesList = new List<string>();

                // Get the title of each slide in the slide order.
                foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                {
                    SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;

                    // Get the slide title.
                    string title = GetSlideTitle(slidePart);

                    // An empty title can also be added.
                    titlesList.Add(title);
                }

                return titlesList;
            }

        }

        return null;
    }

    // Get the title string of the slide.
    public static string GetSlideTitle(SlidePart slidePart)
    {
        if (slidePart == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Declare a paragraph separator.
        string paragraphSeparator = null;

        if (slidePart.Slide != null)
        {
            // Find all the title shapes.
            var shapes = from shape in slidePart.Slide.Descendants<Shape>()
                         where IsTitleShape(shape)
                         select shape;

            StringBuilder paragraphText = new StringBuilder();

            foreach (var shape in shapes)
            {
                // Get the text in each paragraph in this shape.
                foreach (var paragraph in shape.TextBody.Descendants<D.Paragraph>())
                {
                    // Add a line break.
                    paragraphText.Append(paragraphSeparator);

                    foreach (var text in paragraph.Descendants<D.Text>())
                    {
                        paragraphText.Append(text.Text);
                    }

                    paragraphSeparator = "\n";
                }
            }

            return paragraphText.ToString();
        }

        return string.Empty;
    }

    // Determines whether the shape is a title shape.
    private static bool IsTitleShape(Shape shape)
    {
        var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
        if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
        {
            switch ((PlaceholderValues)placeholderShape.Type)
            {
                // Any title shape.
                case PlaceholderValues.Title:

                // A centered title.
                case PlaceholderValues.CenteredTitle:
                    return true;

                default:
                    return false;
            }
        }
        return false;
    }
```

```vb
    ' Get a list of the titles of all the slides in the presentation.
    Public Function GetSlideTitles(ByVal presentationFile As String) As IList(Of String)

        ' Open the presentation as read-only.
        Dim presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, False)
        Using (presentationDocument)
            Return GetSlideTitles(presentationDocument)
        End Using

    End Function
    ' Get a list of the titles of all the slides in the presentation.
    Public Function GetSlideTitles(ByVal presentationDocument As PresentationDocument) As IList(Of String)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Get a PresentationPart object from the PresentationDocument object.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart
        If ((Not (presentationPart) Is Nothing) _
           AndAlso (Not (presentationPart.Presentation) Is Nothing)) Then

            ' Get a Presentation object from the PresentationPart object.
            Dim presentation As Presentation = presentationPart.Presentation
            If (Not (presentation.SlideIdList) Is Nothing) Then

                Dim titlesList As List(Of String) = New List(Of String)

                ' Get the title of each slide in the slide order.
                For Each slideId As Object In presentation.SlideIdList.Elements(Of SlideId)()

                    Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slideId.RelationshipId.ToString()), SlidePart)

                    ' Get the slide title.
                    Dim title As String = GetSlideTitle(slidePart)

                    ' An empty title can also be added.
                    titlesList.Add(title)
                Next
                Return titlesList
            End If
        End If
        Return Nothing
    End Function
    ' Get the title string of the slide.
    Public Function GetSlideTitle(ByVal slidePart As SlidePart) As String
        If (slidePart Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Declare a paragraph separator.
        Dim paragraphSeparator As String = Nothing
        If (Not (slidePart.Slide) Is Nothing) Then

            ' Find all the title shapes.
            Dim shapes = From shape In slidePart.Slide.Descendants(Of Shape)() _
             Where (IsTitleShape(shape)) _
             Select shape

            Dim paragraphText As StringBuilder = New StringBuilder

            For Each shape As Object In shapes

                ' Get the text in each paragraph in this shape.
                For Each paragraph As Object In shape.TextBody.Descendants(Of D.Paragraph)()

                    ' Add a line break.
                    paragraphText.Append(paragraphSeparator)

                    For Each text As Object In paragraph.Descendants(Of D.Text)()
                        paragraphText.Append(text.Text)
                    Next

                    paragraphSeparator = "" & vbLf
                Next
            Next
            Return paragraphText.ToString
        End If
        Return String.Empty
    End Function
    ' Determines whether the shape is a title shape.
    Private Function IsTitleShape(ByVal shape As Shape) As Boolean
        Dim placeholderShape As Object = _
         shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild(Of PlaceholderShape)()
        If ((Not (placeholderShape) Is Nothing) _
           AndAlso ((Not (placeholderShape.Type) Is Nothing) _
           AndAlso placeholderShape.Type.HasValue)) Then
            Select Case placeholderShape.Type.Value

                ' Any title shape
                Case PlaceholderValues.Title
                    Return True

                    ' A centered title.
                Case PlaceholderValues.CenteredTitle
                    Return True
                Case Else
                    Return False
            End Select
        End If
        Return False
    End Function
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library
reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
