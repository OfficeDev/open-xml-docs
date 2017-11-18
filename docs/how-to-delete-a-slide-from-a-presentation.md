---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 98781b17-8de4-46e9-b29a-5b4033665491
title: 'How to: Delete a slide from a presentation (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Delete a slide from a presentation (Open XML SDK)

This topic shows how to use the Open XML SDK 2.5 for Office to delete a
slide from a presentation programmatically. It also shows how to delete
all references to the slide from any custom shows that may exist. To
delete a specific slide in a presentation file you need to know first
the number of slides in the presentation. Therefore the code in this
how-to is divided into two parts. The first is counting the number of
slides, and the second is deleting a slide at a specific index.

> [!NOTE]
> Deleting a slide from more complex presentations, such as those that contain outline view settings, for example, may require additional steps.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Presentation;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System
    Imports System.Collections.Generic
    Imports System.Linq
    Imports DocumentFormat.OpenXml.Presentation
    Imports DocumentFormat.OpenXml.Packaging
```

--------------------------------------------------------------------------------

In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.PresentationDocument"><span
class="nolink">PresentationDocument</span></span> class represents a
presentation document package. To work with a presentation document,
first create an instance of the <span
class="keyword">PresentationDocument</span> class, and then work with
that instance. To create the class instance from the document call one
of the **Open** method overloads. The code in
this topic uses the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(System.String,System.Boolean)"><span
class="nolink">PresentationDocument.Open(String, Boolean)</span></span>
method, which takes a file path as the first parameter to specify the
file to open, and a Boolean value as the second parameter to specify
whether a document is editable. Set this second parameter to <span
class="keyword">false</span> to open the file for read-only access, or
**true** if you want to open the file for
read/write access. The code in this topic opens the file twice, once to
count the number of slides and once to delete a specific slide. When you
count the number of slides in a presentation, it is best to open the
file for read-only access to protect the file against accidental
writing. The following **using** statement
opens the file for read-only access. In this code example, the <span
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

To delete a slide from the presentation file, open it for read/write
access as shown in the following **using**
statement.

```csharp
    // Open the source document as read/write.
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
    {
        // Place other code here.
    }
```

```vb
    ' Open the source document as read/write.
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
        ' Place other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case <span
class="keyword">presentationDocument</span>.


--------------------------------------------------------------------------------

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
two slides denoted by the IDs 267 and 256.

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
class="keyword">notesMaster</span> elements.

| PresentationML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| sld | Slide | Presentation Slide. It is the root element of SlidePart. |
| sldLayout | SlideLayout | Slide Layout. It is the root element of SlideLayoutPart. |
| sldMaster | SlideMaster | Slide Master. It is the root element of SlideMasterPart. |
| notesMaster | NotesMaster | Notes Master (or handoutMaster). It is the root element of NotesMasterPart. |

--------------------------------------------------------------------------------

The sample code consists of two overloads of the <span
class="keyword">CountSlides</span> method. The first overload uses a
**string** parameter and the second overload
uses a **PresentationDocument** parameter. In
the first **CountSlides** method, the sample
code opens the presentation document in the <span
class="keyword">using</span> statement. Then it passes the <span
class="keyword">PresentationDocument</span> object to the second <span
class="keyword">CountSlides</span> method, which returns an integer
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
the **SlideParts** the code gets the slideCount
and returns it.

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

The code for deleting a slide uses two overloads of the <span
class="keyword">DeleteSlide</span> method. The first overloaded <span
class="keyword">DeleteSlide</span> method takes two parameters: a string
that represents the presentation file name and path, and an integer that
represents the zero-based index position of the slide to delete. It
opens the presentation file for read/write access, gets a <span
class="keyword">PresentationDocument</span> object, and then passes that
object and the index number to the next overloaded <span
class="keyword">DeleteSlide</span> method, which performs the deletion.

```csharp
    // Get the presentation object and pass it to the next DeleteSlide method.
    public static void DeleteSlide(string presentationFile, int slideIndex)
    {
        // Open the source document as read/write.

        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
          // Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
          DeleteSlide(presentationDocument, slideIndex);
        }
    }  
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

The first section of the second overloaded <span
class="keyword">DeleteSlide</span> method uses the <span
class="keyword">CountSlides</span> method to get the number of slides in
the presentation. Then, it gets the list of slide IDs in the
presentation, identifies the specified slide in the slide list, and
removes the slide from the slide list.

```csharp
    // Delete the specified slide from the presentation.
    public static void DeleteSlide(PresentationDocument presentationDocument, int slideIndex)
    {
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Use the CountSlides sample to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        if (slideIndex < 0 || slideIndex >= slidesCount)
        {
            throw new ArgumentOutOfRangeException("slideIndex");
        }

        // Get the presentation part from the presentation document. 
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Get the presentation from the presentation part.
        Presentation presentation = presentationPart.Presentation;

        // Get the list of slide IDs in the presentation.
        SlideIdList slideIdList = presentation.SlideIdList;

        // Get the slide ID of the specified slide
        SlideId slideId = slideIdList.ChildElements[slideIndex] as SlideId;

        // Get the relationship ID of the slide.
        string slideRelId = slideId.RelationshipId;

        // Remove the slide from the slide list.
        slideIdList.RemoveChild(slideId);
```

```vb
    ' Delete the specified slide from the presentation.
    Public Shared Sub DeleteSlide(ByVal presentationDocument As 
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Use the CountSlides sample to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)

        If slideIndex < 0 OrElse slideIndex >= slidesCount Then
            Throw New ArgumentOutOfRangeException("slideIndex")
        End If

        ' Get the presentation part from the presentation document. 
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the presentation from the presentation part.
        Dim presentation As Presentation = presentationPart.Presentation

        ' Get the list of slide IDs in the presentation.
        Dim slideIdList As SlideIdList = presentation.SlideIdList

        ' Get the slide ID of the specified slide
        Dim slideId As SlideId = TryCast(slideIdList.ChildElements(slideIndex), SlideId)

        ' Get the relationship ID of the slide.
        Dim slideRelId As String = slideId.RelationshipId

        ' Remove the slide from the slide list.
        slideIdList.RemoveChild(slideId)
```

The next section of the second overloaded <span
class="keyword">DeleteSlide</span> method removes all references to the
deleted slide from custom shows. It does that by iterating through the
list of custom shows and through the list of slides in each custom show.
It then declares and instantiates a linked list of slide list entries,
and finds references to the deleted slide by using the relationship ID
of that slide. It adds those references to the list of slide list
entries, and then removes each such reference from the slide list of its
respective custom show.

```csharp
    // Remove references to the slide from all custom shows.
    if (presentation.CustomShowList != null)
    {
        // Iterate through the list of custom shows.
        foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
        {
            if (customShow.SlideList != null)
            {
                // Declare a link list of slide list entries.
                LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                {
                    // Find the slide reference to remove from the custom show.
                    if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                    {
                        slideListEntries.AddLast(slideListEntry);
                    }
                }

                // Remove all references to the slide from the custom show.
                foreach (SlideListEntry slideListEntry in slideListEntries)
                {
                    customShow.SlideList.RemoveChild(slideListEntry);
                }
            }
        }
    }
```

```vb
    ' Remove references to the slide from all custom shows.
    If presentation.CustomShowList IsNot Nothing Then
        ' Iterate through the list of custom shows.
        For Each customShow In presentation.CustomShowList.Elements(Of CustomShow)()
            If customShow.SlideList IsNot Nothing Then
                ' Declare a link list of slide list entries.
                Dim slideListEntries As New LinkedList(Of SlideListEntry)()
                For Each slideListEntry As SlideListEntry In customShow.SlideList.Elements()
                    ' Find the slide reference to remove from the custom show.
                    If slideListEntry.Id IsNot Nothing AndAlso slideListEntry.Id = slideRelId Then
                        slideListEntries.AddLast(slideListEntry)
                    End If
                Next slideListEntry

                ' Remove all references to the slide from the custom show.
                For Each slideListEntry As SlideListEntry In slideListEntries
                    customShow.SlideList.RemoveChild(slideListEntry)
                Next slideListEntry
            End If
        Next customShow
    End If
```

Finally, the code saves the modified presentation, and deletes the slide
part for the deleted slide.

```csharp
    // Save the modified presentation.
    presentation.Save();

    // Get the slide part for the specified slide.
    SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;

    // Remove the slide part.
    presentationPart.DeletePart(slidePart);
    }
```

```vb
    ' Save the modified presentation.
    presentation.Save()

    ' Get the slide part for the specified slide.
    Dim slidePart As SlidePart = TryCast(presentationPart.GetPartById(slideRelId), SlidePart)

    ' Remove the slide part.
    presentationPart.DeletePart(slidePart)
    End Sub
```

--------------------------------------------------------------------------------

The following is the complete sample code for the two overloaded
methods, **CountSlides** and <span
class="keyword">DeleteSlide</span>. To use the code, you can use the
following call as an example to delete the slide at index 2 in the
presentation file "Myppt6.pptx."

```csharp
    DeleteSlide(@"C:\Users\Public\Documents\Myppt6.pptx", 2);
```

```vb
    DeleteSlide("C:\Users\Public\Documents\Myppt6.pptx", 0)
```

You can also use the following call to count the number of slides in the
presentation.

```csharp
    Console.WriteLine("Number of slides = {0}",
    CountSlides(@"C:\Users\Public\Documents\Myppt6.pptx"));
```

```vb
    Console.WriteLine("Number of slides = {0}", _
    CountSlides("C:\Users\Public\Documents\Myppt6.pptx"))
```

It might be a good idea to count the number of slides before and after
performing the deletion.

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Get the presentation object and pass it to the next CountSlides method.
    public static int CountSlides(string presentationFile)
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
        {
            // Pass the presentation to the next CountSlide method
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
    //
    // Get the presentation object and pass it to the next DeleteSlide method.
    public static void DeleteSlide(string presentationFile, int slideIndex)
    {
        // Open the source document as read/write.

        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
          // Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
          DeleteSlide(presentationDocument, slideIndex);
        }
    }  

    // Delete the specified slide from the presentation.
    public static void DeleteSlide(PresentationDocument presentationDocument, int slideIndex)
    {
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Use the CountSlides sample to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        if (slideIndex < 0 || slideIndex >= slidesCount)
        {
            throw new ArgumentOutOfRangeException("slideIndex");
        }

        // Get the presentation part from the presentation document. 
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Get the presentation from the presentation part.
        Presentation presentation = presentationPart.Presentation;

        // Get the list of slide IDs in the presentation.
        SlideIdList slideIdList = presentation.SlideIdList;

        // Get the slide ID of the specified slide
        SlideId slideId = slideIdList.ChildElements[slideIndex] as SlideId;

        // Get the relationship ID of the slide.
        string slideRelId = slideId.RelationshipId;

        // Remove the slide from the slide list.
        slideIdList.RemoveChild(slideId);

    //
        // Remove references to the slide from all custom shows.
        if (presentation.CustomShowList != null)
        {
            // Iterate through the list of custom shows.
            foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
            {
                if (customShow.SlideList != null)
                {
                    // Declare a link list of slide list entries.
                    LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                    foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                    {
                        // Find the slide reference to remove from the custom show.
                        if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                        {
                            slideListEntries.AddLast(slideListEntry);
                        }
                    }

                    // Remove all references to the slide from the custom show.
                    foreach (SlideListEntry slideListEntry in slideListEntries)
                    {
                        customShow.SlideList.RemoveChild(slideListEntry);
                    }
                }
            }
        }

        // Save the modified presentation.
        presentation.Save();

        // Get the slide part for the specified slide.
        SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;

        // Remove the slide part.
        presentationPart.DeletePart(slidePart);
    }
```

```vb
    ' Count the number of slides in the presentation.
    Public Function CountSlides(ByVal presentationFile As String) As Integer
        ' Open the presentation as read-only.
        Using presentationDocument__1 As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            ' Pass the presentation to the next CountSlides method
            ' and return the slide count.
            Return CountSlides(presentationDocument__1)
        End Using
    End Function
    ' Count the slides in the presentation.
    Public Function CountSlides(ByVal presentationDocument As PresentationDocument) As Integer
        ' Check for a null document object.
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        Dim slidesCount As Integer = 0

        ' Get the presentation part of document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        If presentationPart IsNot Nothing AndAlso presentationPart.Presentation IsNot Nothing Then
            ' Get the Presentation object from the presentation part.
            Dim presentation As Presentation = presentationPart.Presentation

            ' Verify that the presentation contains slides. 
            If presentation.SlideIdList IsNot Nothing Then

                ' Get the slide count from the slide ID list. 
                slidesCount = presentation.SlideIdList.Elements(Of SlideId)().Count()
            End If
        End If

        ' Return the slide count to the previous method.
        Return slidesCount
    End Function
    ' Delete the specified slide from the presentation.
    Public Sub DeleteSlide(ByVal presentationFile As String, ByVal slideIndex As Integer)

        ' Open the source document as read/write.
        Dim presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, True)

        Using (presentationDocument)

            ' Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
            DeleteSlide2(presentationDocument, slideIndex)

        End Using

    End Sub
    ' Delete the specified slide in the presentation.
    Public Sub DeleteSlide2(ByVal presentationDocument As PresentationDocument, ByVal slideIndex As Integer)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Use the CountSlides code example to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)
        If ((slideIndex < 0) OrElse (slideIndex >= slidesCount)) Then
            Throw New ArgumentOutOfRangeException("slideIndex")
        End If

        ' Get the presentation part from the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the presentation from the presentation part. 
        Dim presentation As Presentation = presentationPart.Presentation

        ' Get the list of slide IDs in the presentation.
        Dim slideIdList As SlideIdList = presentation.SlideIdList

        ' Get the slide ID of the specified slide.
        Dim slideId As SlideId = CType(slideIdList.ChildElements(slideIndex), SlideId)

        ' Get the relationship ID of the specified slide.
        Dim slideRelId As String = slideId.RelationshipId

        ' Remove the slide from the slide list.
        slideIdList.RemoveChild(slideId)
        ' Remove references to the slide from all custom shows.
        If (Not (presentation.CustomShowList) Is Nothing) Then

            ' Iterate through the list of custom shows.
            For Each customShow As System.Object In presentation.CustomShowList.Elements(Of  _
                                   DocumentFormat.OpenXml.Presentation.CustomShow)()

                If (Not (customShow.SlideList) Is Nothing) Then

                    ' Declare a linked list.
                    Dim slideListEntries As LinkedList(Of SlideListEntry) = New LinkedList(Of SlideListEntry)

                    ' Iterate through all the slides in the custom show.
                    For Each slideListEntry As SlideListEntry In customShow.SlideList.Elements

                        ' Find the slide reference to be removed from the custom show.
                        If ((Not (slideListEntry.Id) Is Nothing) _
                                    AndAlso (slideListEntry.Id = slideRelId)) Then

                            ' Add that slide reference to the end of the linked list.
                            slideListEntries.AddLast(slideListEntry)
                        End If
                    Next

                    ' Remove references to the slide from the custom show.
                    For Each slideListEntry As SlideListEntry In slideListEntries
                        customShow.SlideList.RemoveChild(slideListEntry)
                    Next
                End If
            Next
        End If

        ' Save the change to the presentation part.
        presentation.Save()

        ' Get the slide part for the specified slide.
        Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slideRelId), SlidePart)

        ' Remove the slide part.
        presentationPart.DeletePart(slidePart)

    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
