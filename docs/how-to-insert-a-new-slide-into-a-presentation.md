---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 6079a1ae-4567-4d99-b350-b819fd06fe5c
title: 'How to: Insert a new slide into a presentation (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Insert a new slide into a presentation (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 to
insert a new slide into a presentation programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Presentation;
    using DocumentFormat.OpenXml.Packaging;
    using Drawing = DocumentFormat.OpenXml.Drawing;
```

```vb
    Imports System
    Imports DocumentFormat.OpenXml.Presentation
    Imports DocumentFormat.OpenXml.Packaging
    Imports Drawing = DocumentFormat.OpenXml.Drawing
```

## Getting a PresentationDocument Object

In the Open XML SDK, the [PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx) class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562287.aspx) method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read/write,
specify the value **true** for this parameter
as shown in the following **using** statement.
In this code segment, the *presentationFile* parameter is a string that
represents the full path for the file from which you want to open the
document.

```csharp
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
    {
        // Insert other code here.
    }
```

```vb
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case
*presentationDocument*.


## Basic Presentation Document Structure

The basic document structure of a **PresentationML** document consists of the main part
that contains the presentation definition. The following text from the
[ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification introduces the overall form of a **PresentationML** package.

> A **PresentationML** package's main part
> starts with a presentation root element. That element contains a
> presentation, which, in turn, refers to a **slide** list, a *slide master* list, a *notes
> master* list, and a *handout master* list. The slide list refers to
> all of the slides in the presentation; the slide master list refers to
> the entire slide masters used in the presentation; the notes master
> contains information about the formatting of notes pages; and the
> handout master describes how a handout looks.
> 
> A *handout* is a printed set of slides that can be provided to an
> *audience* for future reference.
> 
> As well as text and graphics, each slide can contain *comments* and
> *notes*, can have a *layout*, and can be part of one or more *custom
> presentations*. (A comment is an annotation intended for the person
> maintaining the presentation slide deck. A note is a reminder or piece
> of text intended for the presenter or the audience.)
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

The following XML code segment represents a presentation that contains
two slides denoted by the Id's 267 and 256.

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
correspond to the **sld**, **sldLayout**, **sldMaster**, and **notesMaster** elements:

| PresentationML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| sld | [Slide](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slide.aspx) | Presentation Slide. It is the root element of SlidePart. |
| sldLayout | [SlideLayout](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slidelayout.aspx) | Slide Layout. It is the root element of SlideLayoutPart. |
| sldMaster | [SlideMaster](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slidemaster.aspx) | Slide Master. It is the root element of SlideMasterPart. |
| notesMaster | [NotesMaster](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.notesmaster.aspx) | Notes Master (or handoutMaster). It is the root element of NotesMasterPart. |


## How the Sample Code Works 

The sample code consists of two overloads of the **InsertNewSlide** method. The first overloaded
method takes three parameters: the full path to the presentation file to
which to add a slide, an integer that represents the zero-based slide
index position in the presentation where to add the slide, and the
string that represents the title of the new slide. It opens the
presentation file as read/write, gets a **PresentationDocument** object, and then passes that
object to the second overloaded **InsertNewSlide** method, which performs the
insertion.

```csharp
    // Insert a slide into the specified presentation.
     public static void InsertNewSlide(string presentationFile, int position, string slideTitle)
    {
        // Open the source document as read/write. 
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            // Pass the source document and the position and title of the slide to be inserted to the next method.
            InsertNewSlide(presentationDocument, position, slideTitle);
        }
    }
```

```vb
    ' Insert a slide into the specified presentation.
    Public Shared Sub InsertNewSlide(ByVal presentationFile As String, ByVal position As Integer, ByVal slideTitle As String)
        ' Open the source document as read/write. 
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
            ' Pass the source document and the position and title of the slide to be inserted to the next method.
            InsertNewSlide(presentationDocument, position, slideTitle)
        End Using
    End Sub
```

The second overloaded **InsertNewSlide** method
creates a new **Slide** object, sets its
properties, and then inserts it into the slide order in the
presentation. The first section of the method creates the slide and sets
its properties.

```csharp
    // Insert the specified slide into the presentation at the specified position.
    public static void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
    {
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        if (slideTitle == null)
        {
            throw new ArgumentNullException("slideTitle");
        }

        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Verify that the presentation is not empty.
        if (presentationPart == null)
        {
            throw new InvalidOperationException("The presentation document is empty.");
        }

        // Declare and instantiate a new slide.
        Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
        uint drawingObjectId = 1;

        // Construct the slide content.            
        // Specify the non-visual properties of the new slide.
        NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
        nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
        nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
        nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

        // Specify the group shape properties of the new slide.
        slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());
        
    }
```

```vb
    ' Insert the specified slide into the presentation at the specified position.
    Public Shared Sub InsertNewSlide(ByVal presentationDocument As PresentationDocument, ByVal position As Integer, ByVal slideTitle As String)
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        If slideTitle Is Nothing Then
            Throw New ArgumentNullException("slideTitle")
        End If

        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Verify that the presentation is not empty.
        If presentationPart Is Nothing Then
            Throw New InvalidOperationException("The presentation document is empty.")
        End If

        ' Declare and instantiate a new slide.
        Dim slide As New Slide(New CommonSlideData(New ShapeTree()))
        Dim drawingObjectId As UInteger = 1

        ' Construct the slide content.            
        ' Specify the non-visual properties of the new slide.
        Dim nonVisualProperties As NonVisualGroupShapeProperties = slide.CommonSlideData.ShapeTree.AppendChild(New NonVisualGroupShapeProperties())
        nonVisualProperties.NonVisualDrawingProperties = New NonVisualDrawingProperties() With {.Id = 1, .Name = ""}
        nonVisualProperties.NonVisualGroupShapeDrawingProperties = New NonVisualGroupShapeDrawingProperties()
        nonVisualProperties.ApplicationNonVisualDrawingProperties = New ApplicationNonVisualDrawingProperties()

        ' Specify the group shape properties of the new slide.
        slide.CommonSlideData.ShapeTree.AppendChild(New GroupShapeProperties())
```

The next section of the second overloaded **InsertNewSlide** method adds a title shape to the
slide and sets its properties, including its text

```csharp
    // Declare and instantiate the title shape of the new slide.
    Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());

    drawingObjectId++;

    // Specify the required shape properties for the title shape. 
    titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
        (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
        new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
        new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
    titleShape.ShapeProperties = new ShapeProperties();

    // Specify the text of the title shape.
    titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = slideTitle })));
```

```vb
    ' Declare and instantiate the title shape of the new slide.
    Dim titleShape As DocumentFormat.OpenXml.Presentation.Shape = slide.CommonSlideData.ShapeTree.AppendChild _
        (New DocumentFormat.OpenXml.Presentation.Shape())
    drawingObjectId = (drawingObjectId + 1)

    ' Specify the required shape properties for the title shape. 
    titleShape.NonVisualShapeProperties = New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(New _
        DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = drawingObjectId, .Name = "Title"}, _
        New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties _
        (New Drawing.ShapeLocks() With {.NoGrouping = True}), _
        New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Type = PlaceholderValues.Title}))

    titleShape.ShapeProperties = New DocumentFormat.OpenXml.Presentation.ShapeProperties()

    ' Specify the text of the title shape.
    titleShape.TextBody = New DocumentFormat.OpenXml.Presentation.TextBody(New Drawing.BodyProperties, _
         New Drawing.ListStyle, New Drawing.Paragraph _
         (New Drawing.Run(New Drawing.Text() With {.Text = slideTitle})))
```

The next section of the second overloaded **InsertNewSlide** method adds a body shape to the
slide and sets its properties, including its text.

```csharp
    // Declare and instantiate the body shape of the new slide.
    Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
    drawingObjectId++;

    // Specify the required shape properties for the body shape.
    bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
            new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
    bodyShape.ShapeProperties = new ShapeProperties();

    // Specify the text of the body shape.
    bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph());
```

```vb
    ' Declare and instantiate the body shape of the new slide.
    Dim bodyShape As Shape = slide.CommonSlideData.ShapeTree.AppendChild(New Shape())
        drawingObjectId += 1

        ' Specify the required shape properties for the body shape.
        bodyShape.NonVisualShapeProperties = New NonVisualShapeProperties(New NonVisualDrawingProperties() With {.Id = drawingObjectId, .Name = "Content Placeholder"}, New NonVisualShapeDrawingProperties(New Drawing.ShapeLocks() With {.NoGrouping = True}), New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Index = 1}))
        bodyShape.ShapeProperties = New ShapeProperties()

        ' Specify the text of the body shape.
        bodyShape.TextBody = New TextBody(New Drawing.BodyProperties(), New Drawing.ListStyle(), New Drawing.Paragraph())
```

The final section of the second overloaded **InsertNewSlide** method creates a new slide part,
finds the specified index position where to insert the slide, and then
inserts it and saves the modified presentation.

```csharp
    // Create the slide part for the new slide.
    SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

    // Save the new slide part.
    slide.Save(slidePart);

    // Modify the slide ID list in the presentation part.
    // The slide ID list should not be null.
    SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

    // Find the highest slide ID in the current list.
    uint maxSlideId = 1;
    SlideId prevSlideId = null;

    foreach (SlideId slideId in slideIdList.ChildElements)
    {
        if (slideId.Id > maxSlideId)
        {
            maxSlideId = slideId.Id;
        }

    position--;
    if (position == 0)
    {
        prevSlideId = slideId;
    }

    }

    maxSlideId++;

    // Get the ID of the previous slide.
    SlidePart lastSlidePart;

    if (prevSlideId != null)
    {
        lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
    }
    else
    {
        lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
    }

    // Use the same slide layout as that of the previous slide.
    if (null != lastSlidePart.SlideLayoutPart)
    {
        slidePart.AddPart(lastSlidePart.SlideLayoutPart);
    }

    // Insert the new slide into the slide list after the previous slide.
    SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
    newSlideId.Id = maxSlideId;
    newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

    // Save the modified presentation.
    presentationPart.Presentation.Save();
    }
```

```vb
    ' Create the slide part for the new slide.
    Dim slidePart As SlidePart = presentationPart.AddNewPart(Of SlidePart)()

        ' Save the new slide part.
        slide.Save(slidePart)

        ' Modify the slide ID list in the presentation part.
        ' The slide ID list should not be null.
        Dim slideIdList As SlideIdList = presentationPart.Presentation.SlideIdList

        ' Find the highest slide ID in the current list.
        Dim maxSlideId As UInteger = 1
        Dim prevSlideId As SlideId = Nothing

        For Each slideId As SlideId In slideIdList.ChildElements
            If slideId.Id > maxSlideId Then
                maxSlideId = slideId.Id
            End If

            position -= 1
            If position = 0 Then
                prevSlideId = slideId
            End If

        Next slideId

        maxSlideId += 1

        ' Get the ID of the previous slide.
        Dim lastSlidePart As SlidePart

        If prevSlideId IsNot Nothing Then
            lastSlidePart = CType(presentationPart.GetPartById(prevSlideId.RelationshipId), SlidePart)
        Else
            lastSlidePart = CType(presentationPart.GetPartById((CType(slideIdList.ChildElements(0), SlideId)).RelationshipId), SlidePart)
        End If

        ' Use the same slide layout as that of the previous slide.
        If Nothing IsNot lastSlidePart.SlideLayoutPart Then
            slidePart.AddPart(lastSlidePart.SlideLayoutPart)
        End If

        ' Insert the new slide into the slide list after the previous slide.
        Dim newSlideId As SlideId = slideIdList.InsertAfter(New SlideId(), prevSlideId)
        newSlideId.Id = maxSlideId
        newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart)

        ' Save the modified presentation.
        presentationPart.Presentation.Save()
    End Sub
```

## Sample Code

By using the sample code you can add a new slide to an existing
presentation. In your program, you can use the following call to the
**InsertNewSlide** method to add a new slide to
a presentation file named "Myppt10.pptx," with the title "My new slide,"
at position 1.

```csharp
    InsertNewSlide(@"C:\Users\Public\Documents\Myppt10.pptx", 1, "My new slide");
```

```vb
    InsertNewSlide("C:\Users\Public\Documents\Myppt10.pptx", 1, "My new slide")
```

After you have run the program, the new slide would show up as the
second slide in the presentation.

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Insert a slide into the specified presentation.
    public static void InsertNewSlide(string presentationFile, int position, string slideTitle)
    {
        // Open the source document as read/write. 
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            // Pass the source document and the position and title of the slide to be inserted to the next method.
            InsertNewSlide(presentationDocument, position, slideTitle);
        }
    }

    // Insert the specified slide into the presentation at the specified position.
    public static void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
    {

        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        if (slideTitle == null)
        {
            throw new ArgumentNullException("slideTitle");
        }

        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Verify that the presentation is not empty.
        if (presentationPart == null)
        {
            throw new InvalidOperationException("The presentation document is empty.");
        }

        // Declare and instantiate a new slide.
        Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
        uint drawingObjectId = 1;

        // Construct the slide content.            
        // Specify the non-visual properties of the new slide.
        NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
        nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
        nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
        nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

        // Specify the group shape properties of the new slide.
        slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

        // Declare and instantiate the title shape of the new slide.
        Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());

        drawingObjectId++;

        // Specify the required shape properties for the title shape. 
        titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
            (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
            new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
        titleShape.ShapeProperties = new ShapeProperties();

        // Specify the text of the title shape.
        titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = slideTitle })));

        // Declare and instantiate the body shape of the new slide.
        Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
        drawingObjectId++;

        // Specify the required shape properties for the body shape.
        bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
        bodyShape.ShapeProperties = new ShapeProperties();

        // Specify the text of the body shape.
        bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph());

        // Create the slide part for the new slide.
        SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

        // Save the new slide part.
        slide.Save(slidePart);

        // Modify the slide ID list in the presentation part.
        // The slide ID list should not be null.
        SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

        // Find the highest slide ID in the current list.
        uint maxSlideId = 1;
        SlideId prevSlideId = null;

        foreach (SlideId slideId in slideIdList.ChildElements)
        {
            if (slideId.Id > maxSlideId)
            {
                maxSlideId = slideId.Id;
            }

        position--;
        if (position == 0)
        {
            prevSlideId = slideId;
        }

    }

        maxSlideId++;

        // Get the ID of the previous slide.
        SlidePart lastSlidePart;

        if (prevSlideId != null)
        {
            lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
        }
        else
        {
            lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
        }

        // Use the same slide layout as that of the previous slide.
        if (null != lastSlidePart.SlideLayoutPart)
        {
            slidePart.AddPart(lastSlidePart.SlideLayoutPart);
        }

        // Insert the new slide into the slide list after the previous slide.
        SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
        newSlideId.Id = maxSlideId;
        newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

        // Save the modified presentation.
        presentationPart.Presentation.Save();
    }
```

```vb
    ' Insert a slide into the specified presentation.
    Public Sub InsertNewSlide(ByVal presentationFile As String, ByVal position As Integer, ByVal slideTitle As String)

        ' Open the source document as read/write. 
        Dim presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, True)

        Using (presentationDocument)

            'Pass the source document and the position and title of the slide to be inserted to the next method.
            InsertNewSlide(presentationDocument, position, slideTitle)

        End Using

    End Sub
    ' Insert a slide into the specified presentation.
    Public Sub InsertNewSlide(ByVal presentationDocument As PresentationDocument, ByVal position As Integer, ByVal slideTitle As String)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If
        If (slideTitle Is Nothing) Then
            Throw New ArgumentNullException("slideTitle")
        End If

        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Verify that the presentation is not empty.
        If (presentationPart Is Nothing) Then
            Throw New InvalidOperationException("The presentation document is empty.")
        End If

        ' Declare and instantiate a new slide.
        Dim slide As Slide = New Slide(New CommonSlideData(New ShapeTree))
        Dim drawingObjectId As UInteger = 1

        ' Construct the slide content.
        ' Specify the non-visual properties of the new slide.
        Dim nonVisualProperties As DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties = slide.CommonSlideData.ShapeTree.AppendChild(New _
            DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties())
        nonVisualProperties.NonVisualDrawingProperties = New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = 1, .Name = ""}
        nonVisualProperties.NonVisualGroupShapeDrawingProperties = New DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties()
        nonVisualProperties.ApplicationNonVisualDrawingProperties = New ApplicationNonVisualDrawingProperties()

        ' Specify the group shape properties of the new slide.
        slide.CommonSlideData.ShapeTree.AppendChild(New DocumentFormat.OpenXml.Presentation.GroupShapeProperties())
        ' Declare and instantiate the title shape of the new slide.
        Dim titleShape As DocumentFormat.OpenXml.Presentation.Shape = slide.CommonSlideData.ShapeTree.AppendChild _
            (New DocumentFormat.OpenXml.Presentation.Shape())
        drawingObjectId = (drawingObjectId + 1)

        ' Specify the required shape properties for the title shape. 
        titleShape.NonVisualShapeProperties = New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(New _
            DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = drawingObjectId, .Name = "Title"}, _
            New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties _
            (New Drawing.ShapeLocks() With {.NoGrouping = True}), _
            New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Type = PlaceholderValues.Title}))

        titleShape.ShapeProperties = New DocumentFormat.OpenXml.Presentation.ShapeProperties()

        ' Specify the text of the title shape.
        titleShape.TextBody = New DocumentFormat.OpenXml.Presentation.TextBody(New Drawing.BodyProperties, _
             New Drawing.ListStyle, New Drawing.Paragraph _
             (New Drawing.Run(New Drawing.Text() With {.Text = slideTitle})))
        ' Declare and instantiate the body shape of the new slide.
        Dim bodyShape As DocumentFormat.OpenXml.Presentation.Shape = slide.CommonSlideData.ShapeTree.AppendChild _
            (New DocumentFormat.OpenXml.Presentation.Shape())
        drawingObjectId = (drawingObjectId + 1)

        ' Specify the required shape properties for the body shape.
        bodyShape.NonVisualShapeProperties = New NonVisualShapeProperties(New NonVisualDrawingProperties() With {.Id = drawingObjectId, .Name = "ContentPlaceholder"}, _
             New NonVisualShapeDrawingProperties(New Drawing.ShapeLocks() With {.NoGrouping = True}), _
             New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Index = 1}))

        bodyShape.ShapeProperties = New ShapeProperties()

        ' Specify the text of the body shape.
        bodyShape.TextBody = New TextBody(New Drawing.BodyProperties, New Drawing.ListStyle, New Drawing.Paragraph)
        ' Create the slide part for the new slide.
        Dim slidePart As SlidePart = presentationPart.AddNewPart(Of SlidePart)()

        ' Save the new slide part.
        slide.Save(slidePart)

        ' Modify the slide ID list in the presentation part.
        ' The slide ID list should not be null.
        Dim slideIdList As SlideIdList = presentationPart.Presentation.SlideIdList

        ' Find the highest slide ID in the current list.
        Dim maxSlideId As UInteger = 1
        Dim prevSlideId As SlideId = Nothing

        For Each slideId As SlideId In slideIdList.ChildElements
            If (CType(slideId.Id, UInteger) > maxSlideId) Then
                maxSlideId = slideId.Id
            End If
            position = (position - 1)
            If (position = 0) Then
                prevSlideId = slideId
            End If
        Next

        maxSlideId = (maxSlideId + 1)

        ' Get the ID of the previous slide.
        Dim lastSlidePart As SlidePart = Nothing

        If (Not prevSlideId Is Nothing) Then
            lastSlidePart = CType(presentationPart.GetPartById(prevSlideId.RelationshipId), SlidePart)
        Else
            lastSlidePart = CType(presentationPart.GetPartById(CType(slideIdList.ChildElements(0), SlideId).RelationshipId), SlidePart)
        End If


        ' Use the same slide layout as that of the previous slide.
        If (Not (lastSlidePart.SlideLayoutPart) Is Nothing) Then
            slidePart.AddPart(lastSlidePart.SlideLayoutPart)
        End If

        ' Insert the new slide into the slide list after the previous slide.
        Dim newSlideId As SlideId = slideIdList.InsertAfter(New SlideId, prevSlideId)
        newSlideId.Id = maxSlideId
        newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart)

        ' Save the modified presentation.
        presentationPart.Presentation.Save()

    End Sub
```

## See also



- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk.md)
