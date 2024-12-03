---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 98781b17-8de4-46e9-b29a-5b4033665491
title: 'How to: Delete a slide from a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Delete a slide from a presentation

This topic shows how to use the Open XML SDK for Office to delete a
slide from a presentation programmatically. It also shows how to delete
all references to the slide from any custom shows that may exist. To
delete a specific slide in a presentation file you need to know first
the number of slides in the presentation. Therefore the code in this
how-to is divided into two parts. The first is counting the number of
slides, and the second is deleting a slide at a specific index.

> [!NOTE]
> Deleting a slide from more complex presentations, such as those that contain outline view settings, for example, may require additional steps.



--------------------------------------------------------------------------------
## Getting a Presentation Object 

In the Open XML SDK, the **<xref:DocumentFormat.OpenXml.Packaging.PresentationDocument>** class represents a presentation document package. To work with a presentation document, first create an instance of the **PresentationDocument** class, and then work with that instance. To create the class instance from the document call one of the **Open** method overloads. The code in this topic uses the **[PresentationDocument.Open(String, Boolean)](/dotnet/api/documentformat.openxml.packaging.presentationdocument.open)** method, which takes a file path as the first parameter to specify the file to open, and a Boolean value as the second parameter to specify whether a document is editable. Set this second parameter to **false** to open the file for read-only access, or **true** if you want to open the file for read/write access. The code in this topic opens the file twice, once to count the number of slides and once to delete a specific slide. When you count the number of slides in a presentation, it is best to open the file for read-only access to protect the file against accidental writing. The following **using** statement opens the file for read-only access. In this code example, the **presentationFile** parameter is a string that represents the path for the file from which you want to open the document.

### [C#](#tab/cs-0)
```csharp
    // Open the presentation as read-only.
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
    {
        // Insert other code here.
    }                                                                           
```

### [Visual Basic](#tab/vb-0)
```vb
    ' Open the presentation as read-only.
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
        ' Insert other code here.
    End Using
```
***


To delete a slide from the presentation file, open it for read/write
access as shown in the following **using**
statement.

### [C#](#tab/cs-1)
```csharp
    // Open the source document as read/write.
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
    {
        // Place other code here.
    }
```

### [Visual Basic](#tab/vb-1)
```vb
    ' Open the source document as read/write.
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
        ' Place other code here.
    End Using
```
***


[!include[Using Statement](../includes/presentation/using-statement.md)] `presentationDocument`.

[!include[Structure](../includes/presentation/structure.md)]

## Counting the Number of Slides 

The sample code consists of two overloads of the **CountSlides** method. The first overload uses a **string** parameter and the second overload uses a **PresentationDocument** parameter. In the first **CountSlides** method, the sample code opens the presentation document in the **using** statement. Then it passes the **PresentationDocument** object to the second **CountSlides** method, which returns an integer number that represents the number of slides in the presentation.

### [C#](#tab/cs-2)
```csharp
    // Pass the presentation to the next CountSlides method
    // and return the slide count.
    return CountSlides(presentationDocument);
```

### [Visual Basic](#tab/vb-2)
```vb
    ' Pass the presentation to the next CountSlides method
    ' and return the slide count.
    Return CountSlides(presentationDocument)
```
***


In the second **CountSlides** method, the code
verifies that the **PresentationDocument**
object passed in is not **null**, and if it is
not, it gets a **PresentationPart** object from
the **PresentationDocument** object. By using
the **SlideParts** the code gets the slideCount
and returns it.

### [C#](#tab/cs-3)
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

### [Visual Basic](#tab/vb-3)
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
***


--------------------------------------------------------------------------------
## Deleting a Specific Slide 

The code for deleting a slide uses two overloads of the **DeleteSlide** method. The first overloaded **DeleteSlide** method takes two parameters: a string
that represents the presentation file name and path, and an integer that
represents the zero-based index position of the slide to delete. It
opens the presentation file for read/write access, gets a **PresentationDocument** object, and then passes that
object and the index number to the next overloaded **DeleteSlide** method, which performs the deletion.

### [C#](#tab/cs-4)
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

### [Visual Basic](#tab/vb-4)
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
***


The first section of the second overloaded **DeleteSlide** method uses the **CountSlides** method to get the number of slides in
the presentation. Then, it gets the list of slide IDs in the
presentation, identifies the specified slide in the slide list, and
removes the slide from the slide list.

### [C#](#tab/cs-5)
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

### [Visual Basic](#tab/vb-5)
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
***


The next section of the second overloaded **DeleteSlide** method removes all references to the
deleted slide from custom shows. It does that by iterating through the
list of custom shows and through the list of slides in each custom show.
It then declares and instantiates a linked list of slide list entries,
and finds references to the deleted slide by using the relationship ID
of that slide. It adds those references to the list of slide list
entries, and then removes each such reference from the slide list of its
respective custom show.

### [C#](#tab/cs-6)
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

### [Visual Basic](#tab/vb-6)
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
***


Finally, the code saves the modified presentation, and deletes the slide
part for the deleted slide.

### [C#](#tab/cs-7)
```csharp
    // Save the modified presentation.
    presentation.Save();

    // Get the slide part for the specified slide.
    SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;

    // Remove the slide part.
    presentationPart.DeletePart(slidePart);
    }
```

### [Visual Basic](#tab/vb-7)
```vb
    ' Save the modified presentation.
    presentation.Save()

    ' Get the slide part for the specified slide.
    Dim slidePart As SlidePart = TryCast(presentationPart.GetPartById(slideRelId), SlidePart)

    ' Remove the slide part.
    presentationPart.DeletePart(slidePart)
    End Sub
```
***


--------------------------------------------------------------------------------
## Sample Code 

The following is the complete sample code for the two overloaded
methods, **CountSlides** and **DeleteSlide**. To use the code, you can use the
following call as an example to delete the slide at index 2 in the
presentation file "Myppt6.pptx."

### [C#](#tab/cs-8)
```csharp
    DeleteSlide(@"C:\Users\Public\Documents\Myppt6.pptx", 2);
```

### [Visual Basic](#tab/vb-8)
```vb
    DeleteSlide("C:\Users\Public\Documents\Myppt6.pptx", 0)
```
***


You can also use the following call to count the number of slides in the
presentation.

### [C#](#tab/cs-9)
```csharp
    Console.WriteLine("Number of slides = {0}",
    CountSlides(@"C:\Users\Public\Documents\Myppt6.pptx"));
```

### [Visual Basic](#tab/vb-9)
```vb
    Console.WriteLine("Number of slides = {0}", _
    CountSlides("C:\Users\Public\Documents\Myppt6.pptx"))
```
***


It might be a good idea to count the number of slides before and after
performing the deletion.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/delete_a_slide_from/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/delete_a_slide_from/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also 



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
