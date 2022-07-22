---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 95fd9dcd-41e9-4e83-9191-2f3110ae73d5
title: 'How to: Move a slide to a new position in a presentation (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Move a slide to a new position in a presentation (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to move a slide to a new position in a presentation
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

--------------------------------------------------------------------------------
## Getting a Presentation Object
In the Open XML SDK, the [PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx) class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the
[Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562287.aspx) method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. In order to count the number of slides
in a presentation, it is best to open the file for read-only access in
order to avoid accidental writing to the file. To do that, specify the
value **false** for the Boolean parameter as
shown in the following **using** statement. In
this code, the **presentationFile** parameter
is a string that represents the path for the file from which you want to
open the document.

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
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **presentationDocument**.


--------------------------------------------------------------------------------
## Basic Presentation Document Structure
The basic document structure of a **PresentationML** document consists of a number of
parts, among which is the main part that contains the presentation
definition. The following text from the [ISO/IEC
29500](https://www.iso.org/standard/71691.html) specification
introduces the overall form of a **PresentationML** package.

> A **PresentationML** package's main part
> starts with a presentation root element. That element contains a
> presentation, which, in turn, refers to a <span
> class="keyword">slide** list, a *slide master* list, a *notes
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

This following XML code segment represents a presentation that contains
two slides denoted by the ID 267 and 256.

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


--------------------------------------------------------------------------------
## How the Sample Code Works
In order to move a specific slide in a presentation file to a new
position, you need to know first the number of slides in the
presentation. Therefore, the code in this topic is divided into two
parts. The first is counting the number of slides, and the second is
moving a slide to a new position.


--------------------------------------------------------------------------------
## Counting the Number of Slides
The sample code for counting the number of slides consists of two
overloads of the method **CountSlides**. The
first overload uses a **string** parameter and
the second overload uses a **PresentationDocument** parameter. In the first
**CountSlides** method, the sample code opens
the presentation document in the **using**
statement. Then it passes the **PresentationDocument** object to the second **CountSlides** method, which returns an integer
number that represents the number of slides in the presentation.

```csharp
    // Pass the presentation to the next CountSlides method
    // and return the slide count.
    return CountSlides(presentationDocument);
```

```vb
    ' Pass the presentation to the next CountSlides method
    ' and return the slide count.
    Return CountSlides(presentationDocument)
```

In the second **CountSlides** method, the code
verifies that the **PresentationDocument**
object passed in is not **null**, and if it is
not, it gets a **PresentationPart** object from
the **PresentationDocument** object. By using
the **SlideParts** the code gets the **slideCount** and returns it.

```csharp
    // Check for a null document object.
    if (presentationDocument == null)
    {
        throw new ArgumentNullException("presentationDocument");
    }

    int slidesCount = 0;

    // Get the presentation part of document.
    PresentationPart presentationPart = presentationDocument.PresentationPart;

    // Get the slide count from the SlideParts.
    if (presentationPart != null)
    {
        slidesCount = presentationPart.SlideParts.Count();
    }
    // Return the slide count to the previous method.
    return slidesCount;
```

```vb
    ' Check for a null document object.
    If presentationDocument Is Nothing Then
        Throw New ArgumentNullException("presentationDocument")
    End If

    Dim slidesCount As Integer = 0

    ' Get the presentation part of document.
    Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

    ' Get the slide count from the SlideParts.
    If presentationPart IsNot Nothing Then
        slidesCount = presentationPart.SlideParts.Count()
    End If
    ' Return the slide count to the previous method.
    Return slidesCount
```

--------------------------------------------------------------------------------
## Moving a Slide from one Position to Another
Moving a slide to a new position requires opening the file for
read/write access by specifying the value **true** to the Boolean parameter as shown in the
following **using** statement. The code for
moving a slide consists of two overloads of the **MoveSlide** method. The first overloaded **MoveSlide** method takes three parameters: a string
that represents the presentation file name and path and two integers
that represent the current index position of the slide and the index
position to which to move the slide respectively. It opens the
presentation file, gets a **PresentationDocument** object, and then passes that
object and the two integers, *from* and *to*, to the second overloaded
**MoveSlide** method, which performs the actual
move.

```csharp
    // Move a slide to a different position in the slide order in the presentation.
    public static void MoveSlide(string presentationFile, int from, int to)
    {
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            MoveSlide(presentationDocument, from, to);
        }
    }
```

```vb
    ' Move a slide to a different position in the slide order in the presentation.
    Public Shared Sub MoveSlide(ByVal presentationFile As String, ByVal [from] As Integer, ByVal [to] As Integer)
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
            MoveSlide(presentationDocument, From, [to])
        End Using
    End Sub
```

In the second overloaded **MoveSlide** method,
the **CountSlides** method is called to get the
number of slides in the presentation. The code then checks if the
zero-based indexes, *from* and *to*, are within the range and different
from one another.

```csharp
    public static void MoveSlide(PresentationDocument presentationDocument, int from, int to)
    {
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Call the CountSlides method to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        // Verify that both from and to positions are within range and different from one another.
        if (from < 0 || from >= slidesCount)
        {
            throw new ArgumentOutOfRangeException("from");
        }

        if (to < 0 || from >= slidesCount || to == from)
        {
            throw new ArgumentOutOfRangeException("to");
        }
```

```vb
    Public Shared Sub MoveSlide(ByVal presentationDocument As PresentationDocument, ByVal [from] As Integer, ByVal [to] As Integer)
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Call the CountSlides method to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)

        ' Verify that both from and to positions are within range and different from one another.
        If
            From < 0 OrElse
            From >= slidesCount Then
            Throw New ArgumentOutOfRangeException("from")
        End If

        If [to] < 0 OrElse
            From >= slidesCount OrElse [to] =
            From Then
            Throw New ArgumentOutOfRangeException("to")
        End If
```

A **PresentationPart** object is declared and
set equal to the presentation part of the **PresentationDocument** object passed in. The **PresentationPart** object is used to create a **Presentation** object, and then create a **SlideIdList** object that represents the list of
slides in the presentation from the **Presentation** object. A slide ID of the source
slide (the slide to move) is obtained, and then the position of the
target slide (the slide after which in the slide order to move the
source slide) is identified.

```csharp
    // Get the presentation part from the presentation document.
    PresentationPart presentationPart = presentationDocument.PresentationPart;

    // The slide count is not zero, so the presentation must contain slides.            
    Presentation presentation = presentationPart.Presentation;
    SlideIdList slideIdList = presentation.SlideIdList;

    // Get the slide ID of the source slide.
    SlideId sourceSlide = slideIdList.ChildElements[from] as SlideId;

    SlideId targetSlide = null;

    // Identify the position of the target slide after which to move the source slide.
    if (to == 0)
    {
        targetSlide = null;
    }
    else if (from < to)
    {
        targetSlide = slideIdList.ChildElements[to] as SlideId;
    }
    else
    {
        targetSlide = slideIdList.ChildElements[to - 1] as SlideId;
    }
```

```vb
    ' Get the presentation part from the presentation document.
    Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

    ' The slide count is not zero, so the presentation must contain slides.            
    Dim presentation As Presentation = presentationPart.Presentation
    Dim slideIdList As SlideIdList = presentation.SlideIdList

    ' Get the slide ID of the source slide.
    Dim sourceSlide As SlideId = TryCast(slideIdList.ChildElements(From), SlideId)

    Dim targetSlide As SlideId = Nothing

    ' Identify the position of the target slide after which to move the source slide.
    If to = 0 Then
        targetSlide = Nothing
    ElseIf From < to Then
        targetSlide = TryCast(slideIdList.ChildElements(to), SlideId)
    Else
        targetSlide = TryCast(slideIdList.ChildElements(to - 1), SlideId)
    End If
```

The **Remove** method of the **SlideID** object is used to remove the source slide
from its current position, and then the **InsertAfter** method of the **SlideIdList** object is used to insert the source
slide in the index position after the target slide. Finally, the
modified presentation is saved.

```csharp
    // Remove the source slide from its current position.
    sourceSlide.Remove();

    // Insert the source slide at its new position after the target slide.
    slideIdList.InsertAfter(sourceSlide, targetSlide);

    // Save the modified presentation.
    presentation.Save();
```

```vb
    ' Remove the source slide from its current position.
    sourceSlide.Remove()

    ' Insert the source slide at its new position after the target slide.
    slideIdList.InsertAfter(sourceSlide, targetSlide)

    ' Save the modified presentation.
    presentation.Save()
```

--------------------------------------------------------------------------------
## Sample Code
Following is the complete sample code that you can use to move a slide
from one position to another in the same presentation file. For
instance, you can use the following call in your program to move a slide
from position 0 to position 1 in a presentation file named
"Myppt11.pptx".

```csharp
    MoveSlide(@"C:\Users\Public\Documents\Myppt11.pptx", 0, 1);
```

```vb
    MoveSlide("C:\Users\Public\Documents\Myppt11.pptx", 0, 1)
```

After you run the program, check your presentation file to see the new
positions of the slides.

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Counting the slides in the presentation.
     public static int CountSlides(string presentationFile)
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
        {
            // Pass the presentation to the next CountSlides method
            // and return the slide count.
            return CountSlides(presentationDocument);
        }
    }

    // Count the slides in the presentation.
    public static int CountSlides(PresentationDocument presentationDocument)
    {
        // Check for a null document object.
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        int slidesCount = 0;

        // Get the presentation part of document.
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Get the slide count from the SlideParts.
        if (presentationPart != null)
        {
            slidesCount = presentationPart.SlideParts.Count();
        }

        // Return the slide count to the previous method.
        return slidesCount;
    }

    // Move a slide to a different position in the slide order in the presentation.
    public static void MoveSlide(string presentationFile, int from, int to)
    {
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            MoveSlide(presentationDocument, from, to);
        }
    }
    // Move a slide to a different position in the slide order in the presentation.
    public static void MoveSlide(PresentationDocument presentationDocument, int from, int to)
    {
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Call the CountSlides method to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        // Verify that both from and to positions are within range and different from one another.
        if (from < 0 || from >= slidesCount)
        {
            throw new ArgumentOutOfRangeException("from");
        }

        if (to < 0 || from >= slidesCount || to == from)
        {
            throw new ArgumentOutOfRangeException("to");
        }

        // Get the presentation part from the presentation document.
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // The slide count is not zero, so the presentation must contain slides.            
        Presentation presentation = presentationPart.Presentation;
        SlideIdList slideIdList = presentation.SlideIdList;

        // Get the slide ID of the source slide.
        SlideId sourceSlide = slideIdList.ChildElements[from] as SlideId;

        SlideId targetSlide = null;

        // Identify the position of the target slide after which to move the source slide.
        if (to == 0)
        {
            targetSlide = null;
        }
        else if (from < to)
        {
            targetSlide = slideIdList.ChildElements[to] as SlideId;
        }
        else
        {
            targetSlide = slideIdList.ChildElements[to - 1] as SlideId;
        }

        // Remove the source slide from its current position.
        sourceSlide.Remove();

        // Insert the source slide at its new position after the target slide.
        slideIdList.InsertAfter(sourceSlide, targetSlide);

        // Save the modified presentation.
        presentation.Save();
    }
```

```vb
    ' Count the slides in the presentation.
    Public Function CountSlides(ByVal presentationFile As String) As Integer

        ' Open the presentation as read-only.
        Dim presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, False)
        Using (presentationDocument)

            ' Pass the presentation to the next CountSlide method
            ' and return the slide count.
            Return CountSlides(presentationDocument)
        End Using
    End Function
    ' Count the slides in the presentation.
    Public Function CountSlides(ByVal presentationDocument As PresentationDocument) As Integer

        ' Check for a null document object.
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If
        Dim slidesCount As Integer = 0

        ' Get the presentation part of the document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart
        If ((Not (presentationPart) Is Nothing) AndAlso (Not (presentationPart.Presentation) Is Nothing)) Then

            ' Get the Presentation object from the presentation part.
            Dim presentation As Presentation = presentationPart.Presentation
            If (Not (presentation.SlideIdList) Is Nothing) Then

                ' Get the slide count from the slide ID list.
                slidesCount = presentation.SlideIdList.Elements.Count()

            End If
        End If

        ' Return the slide count to the previous function.
        Return slidesCount
    End Function
    ' Move a slide to a different position in the slide order in the presentation.
    Public Sub MoveSlide(ByVal presentationFile As String, ByVal from As Integer, ByVal moveTo As Integer)
        Dim presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, True)

        Using (presentationDocument)
            MoveSlide(presentationDocument, from, moveTo)
        End Using

    End Sub
    ' Move a slide to a different position in the slide order in the presentation.
    Public Sub MoveSlide(ByVal presentationDocument As PresentationDocument, ByVal from As Integer, ByVal moveTo As Integer)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Use the CountSlides sample to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)

        ' Verify that both from and to positions are within range and different from one another.
        If ((from < 0) OrElse (from >= slidesCount)) Then
            Throw New ArgumentOutOfRangeException("from")
        End If

        If ((moveTo < 0) _
                    OrElse ((from >= slidesCount) _
                    OrElse (moveTo = from))) Then
            Throw New ArgumentOutOfRangeException("moveTo")
        End If

        ' Get the presentation part from the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' The slide count is not zero, so the presentation must contain slides. 
        Dim presentation As Presentation = presentationPart.Presentation
        Dim slideIdList As SlideIdList = presentation.SlideIdList

        ' Get the slide ID of the source slide.
        Dim sourceSlide As SlideId = CType(slideIdList.ChildElements(from), SlideId)
        Dim targetSlide As SlideId = Nothing

        ' Identify the position of the target slide after which to move the source slide.
        If (moveTo = 0) Then
            targetSlide = Nothing
        ElseIf (from < moveTo) Then
            targetSlide = CType(slideIdList.ChildElements(moveTo), SlideId)
        Else
            targetSlide = CType(slideIdList.ChildElements((moveTo - 1)), SlideId)
        End If

        ' Remove the source slide from its current position.
        sourceSlide.Remove()

        ' Insert the source slide at its new position after the target slide.
        slideIdList.InsertAfter(sourceSlide, targetSlide)

        ' Save the modified presentation.
        presentation.Save()

    End Sub
```

--------------------------------------------------------------------------------
## See also


- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)
