---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 6de46612-f864-413f-a504-11ea85f1f88f
title: 'How to: Get all the text in a slide in a presentation (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Get all the text in a slide in a presentation (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to get all the text in a slide in a presentation
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using DocumentFormat.OpenXml.Presentation;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System
    Imports System.Collections.Generic
    Imports System.Linq
    Imports System.Text
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
that instance. To create the class instance from the document call the
<span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(System.String,System.Boolean)"><span
class="nolink">PresentationDocument.Open(String, Boolean)</span></span>
method that uses a file path, and a Boolean value as the second
parameter to specify whether a document is editable. To open a document
for read/write access, assign the value <span
class="keyword">true</span> to this parameter; for read-only access
assign it the value **false** as shown in the
following **using** statement. In this code,
the **file** parameter is a string that
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
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(file, False)
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
class="keyword">using</span> statement, in this case <span
class="keyword">presentationDocument</span>.


--------------------------------------------------------------------------------

The basic document structure of a <span
class="keyword">PresentationML</span> document consists of the main part
that contains the presentation definition. The following text from the
[ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337)
specification introduces the overall form of a <span
class="keyword">PresentationML</span> package.

> A PresentationML package's main part starts with a presentation root
> element. That element contains a presentation, which, in turn, refers
> to a <span class="term">slide</span> list, a <span class="term">slide
> master</span> list, a <span class="term">notes master</span> list, and
> a <span class="term">handout master</span> list. The slide list refers
> to all of the slides in the presentation; the slide master list refers
> to the entire slide masters used in the presentation; the notes master
> contains information about the formatting of notes pages; and the
> handout master describes how a handout looks.

> A <span class="term">handout</span> is a printed set of slides that
> can be provided to an <span class="term">audience</span> for future
> reference.

> As well as text and graphics, each slide can contain <span
> class="term">comments</span> and <span class="term">notes</span>, can
> have a <span class="term">layout</span>, and can be part of one or
> more <span class="term">custom presentations</span>. (A comment is an
> annotation intended for the person maintaining the presentation slide
> deck. A note is a reminder or piece of text intended for the presenter
> or the audience.)

> Other features that a PresentationML document can include the
> following: <span class="term">animation</span>, <span
> class="term">audio</span>, <span class="term">video</span>, and <span
> class="term">transitions</span> between slides.

> A PresentationML document is not stored as one large body in a single
> part. Instead, the elements that implement certain groupings of
> functionality are stored in separate parts. For example, all comments
> in a document are stored in one comment part while each slide has its
> own part.

> © ISO/IEC29500: 2008.

The following XML code segment represents a presentation that contains
two slides denoted by the ID 267 and 256.

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
content using strongly-typed classes that correspond to <span
class="keyword">PresentationML</span> elements. You can find these
classes in the <span sdata="cer"
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

The sample code consists of three overloads of the <span
class="keyword">GetAllTextInSlide</span> method. In the following
segment, the first overloaded method opens the source presentation that
contains the slide with text to get, and passes the presentation to the
second overloaded method, which gets the slide part. This method returns
the array of strings that the second method returns to it, each of which
represents a paragraph of text in the specified slide.

```csharp
    // Get all the text in a slide.
    public static string[] GetAllTextInSlide(string presentationFile, int slideIndex)
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
        {
            // Pass the presentation and the slide index
            // to the next GetAllTextInSlide method, and
            // then return the array of strings it returns. 
            return GetAllTextInSlide(presentationDocument, slideIndex);
        }
    }
```

```vb
    ' Get all the text in a slide.
    Public Shared Function GetAllTextInSlide(ByVal presentationFile As String, ByVal slideIndex As Integer) As String()
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            ' Pass the presentation and the slide index
            ' to the next GetAllTextInSlide method, and
            ' then return the array of strings it returns. 
            Return GetAllTextInSlide(presentationDocument, slideIndex)
        End Using
    End Function
```

The second overloaded method takes the presentation document passed in
and gets a slide part to pass to the third overloaded method. It returns
to the first overloaded method the array of strings that the third
overloaded method returns to it, each of which represents a paragraph of
text in the specified slide.

```csharp
    public static string[] GetAllTextInSlide(PresentationDocument presentationDocument, int slideIndex)
    {
        // Verify that the presentation document exists.
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Verify that the slide index is not out of range.
        if (slideIndex < 0)
        {
            throw new ArgumentOutOfRangeException("slideIndex");
        }

        // Get the presentation part of the presentation document.
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Verify that the presentation part and presentation exist.
        if (presentationPart != null && presentationPart.Presentation != null)
        {
            // Get the Presentation object from the presentation part.
            Presentation presentation = presentationPart.Presentation;

            // Verify that the slide ID list exists.
            if (presentation.SlideIdList != null)
            {
                // Get the collection of slide IDs from the slide ID list.
                var slideIds = presentation.SlideIdList.ChildElements;

                // If the slide ID is in range...
                if (slideIndex < slideIds.Count)
                {
                    // Get the relationship ID of the slide.
                    string slidePartRelationshipId = (slideIds[slideIndex] as SlideId).RelationshipId;

                    // Get the specified slide part from the relationship ID.
                    SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                    // Pass the slide part to the next method, and
                    // then return the array of strings that method
                    // returns to the previous method.
                    return GetAllTextInSlide(slidePart);
                }
            }
        }
        // Else, return null.
        return null;
    }
```

```vb
    Public Shared Function GetAllTextInSlide(ByVal presentationDocument As PresentationDocument, ByVal slideIndex As Integer) As String()
        ' Verify that the presentation document exists.
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Verify that the slide index is not out of range.
        If slideIndex < 0 Then
            Throw New ArgumentOutOfRangeException("slideIndex")
        End If

        ' Get the presentation part of the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Verify that the presentation part and presentation exist.
        If presentationPart IsNot Nothing AndAlso presentationPart.Presentation IsNot Nothing Then
            ' Get the Presentation object from the presentation part.
            Dim presentation As Presentation = presentationPart.Presentation

            ' Verify that the slide ID list exists.
            If presentation.SlideIdList IsNot Nothing Then
                ' Get the collection of slide IDs from the slide ID list.
                Dim slideIds = presentation.SlideIdList.ChildElements

                ' If the slide ID is in range...
                If slideIndex < slideIds.Count Then
                    ' Get the relationship ID of the slide.
                    Dim slidePartRelationshipId As String = (TryCast(slideIds(slideIndex), SlideId)).RelationshipId

                    ' Get the specified slide part from the relationship ID.
                    Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slidePartRelationshipId), SlidePart)

                    ' Pass the slide part to the next method, and
                    ' then return the array of strings that method
                    ' returns to the previous method.
                    Return GetAllTextInSlide(slidePart)
                End If
            End If
        End If

        ' Else, return null.
        Return Nothing
    End Function
```

The following code segment shows the third overloaded method, which
takes takes the slide part passed in, and returns to the second
overloaded method a string array of text paragraphs. It starts by
verifying that the slide part passed in exists, and then it creates a
linked list of strings. It iterates through the paragraphs in the slide
passed in, and using a **StringBuilder** object
to concatenate all the lines of text in a paragraph, it assigns each
paragraph to a string in the linked list. It then returns to the second
overloaded method an array of strings that represents all the text in
the specified slide in the presentation.

```csharp
    public static string[] GetAllTextInSlide(SlidePart slidePart)
    {
        // Verify that the slide part exists.
        if (slidePart == null)
        {
            throw new ArgumentNullException("slidePart");
        }

        // Create a new linked list of strings.
        LinkedList<string> texts = new LinkedList<string>();

        // If the slide exists...
        if (slidePart.Slide != null)
        {
            // Iterate through all the paragraphs in the slide.
            foreach (var paragraph in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                // Create a new string builder.                    
                StringBuilder paragraphText = new StringBuilder();

                // Iterate through the lines of the paragraph.
                foreach (var text in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                {
                    // Append each line to the previous lines.
                    paragraphText.Append(text.Text);
                }

                if (paragraphText.Length > 0)
                {
                    // Add each paragraph to the linked list.
                    texts.AddLast(paragraphText.ToString());
                }
            }
        }

        if (texts.Count > 0)
        {
            // Return an array of strings.
            return texts.ToArray();
        }
        else
        {
            return null;
        }
    }
```

```vb
    Public Shared Function GetAllTextInSlide(ByVal slidePart As SlidePart) As String()
        ' Verify that the slide part exists.
        If slidePart Is Nothing Then
            Throw New ArgumentNullException("slidePart")
        End If

        ' Create a new linked list of strings.
        Dim texts As New LinkedList(Of String)()

        ' If the slide exists...
        If slidePart.Slide IsNot Nothing Then
            ' Iterate through all the paragraphs in the slide.
            For Each paragraph In slidePart.Slide.Descendants(Of DocumentFormat.OpenXml.Drawing.Paragraph)()
                ' Create a new string builder.                    
                Dim paragraphText As New StringBuilder()

                ' Iterate through the lines of the paragraph.
                For Each text In paragraph.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)()
                    ' Append each line to the previous lines.
                    paragraphText.Append(text.Text)
                Next text

                If paragraphText.Length > 0 Then
                    ' Add each paragraph to the linked list.
                    texts.AddLast(paragraphText.ToString())
                End If
            Next paragraph
        End If

        If texts.Count > 0 Then
            ' Return an array of strings.
            Return texts.ToArray()
        Else
            Return Nothing
        End If
    End Function
```

--------------------------------------------------------------------------------

Following is the complete sample code that you can use to get all the
text in a specific slide in a presentation file. For example, you can
use the following **foreach** loop in your
program to get the array of strings returned by the method <span
class="keyword">GetAllTextInSlide</span>, which represents the text in
the second slide of the presentation file "Myppt8.pptx."

```csharp
    foreach (string s in GetAllTextInSlide(@"C:\Users\Public\Documents\Myppt8.pptx", 1))
        Console.WriteLine(s);
```

```vb
    For Each s As String In GetAllTextInSlide("C:\Users\Public\Documents\Myppt8.pptx", 1)
        Console.WriteLine(s)
    Next
```

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Get all the text in a slide.
    public static string[] GetAllTextInSlide(string presentationFile, int slideIndex)
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
        {
            // Pass the presentation and the slide index
            // to the next GetAllTextInSlide method, and
            // then return the array of strings it returns. 
            return GetAllTextInSlide(presentationDocument, slideIndex);
        }
    }
    public static string[] GetAllTextInSlide(PresentationDocument presentationDocument, int slideIndex)
    {
        // Verify that the presentation document exists.
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Verify that the slide index is not out of range.
        if (slideIndex < 0)
        {
            throw new ArgumentOutOfRangeException("slideIndex");
        }

        // Get the presentation part of the presentation document.
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Verify that the presentation part and presentation exist.
        if (presentationPart != null && presentationPart.Presentation != null)
        {
            // Get the Presentation object from the presentation part.
            Presentation presentation = presentationPart.Presentation;

            // Verify that the slide ID list exists.
            if (presentation.SlideIdList != null)
            {
                // Get the collection of slide IDs from the slide ID list.
                DocumentFormat.OpenXml.OpenXmlElementList slideIds = 
                    presentation.SlideIdList.ChildElements;

                // If the slide ID is in range...
                if (slideIndex < slideIds.Count)
                {
                    // Get the relationship ID of the slide.
                    string slidePartRelationshipId = (slideIds[slideIndex] as SlideId).RelationshipId;

                    // Get the specified slide part from the relationship ID.
                    SlidePart slidePart = 
                        (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                    // Pass the slide part to the next method, and
                    // then return the array of strings that method
                    // returns to the previous method.
                    return GetAllTextInSlide(slidePart);
                }
            }
        }

        // Else, return null.
        return null;
    }
    public static string[] GetAllTextInSlide(SlidePart slidePart)
    {
        // Verify that the slide part exists.
        if (slidePart == null)
        {
            throw new ArgumentNullException("slidePart");
        }

        // Create a new linked list of strings.
        LinkedList<string> texts = new LinkedList<string>();

        // If the slide exists...
        if (slidePart.Slide != null)
        {
            // Iterate through all the paragraphs in the slide.
            foreach (DocumentFormat.OpenXml.Drawing.Paragraph paragraph in 
                slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                // Create a new string builder.                    
                StringBuilder paragraphText = new StringBuilder();

                // Iterate through the lines of the paragraph.
                foreach (DocumentFormat.OpenXml.Drawing.Text text in 
                    paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                {
                    // Append each line to the previous lines.
                    paragraphText.Append(text.Text);
                }

                if (paragraphText.Length > 0)
                {
                    // Add each paragraph to the linked list.
                    texts.AddLast(paragraphText.ToString());
                }
            }
        }

        if (texts.Count > 0)
        {
            // Return an array of strings.
            return texts.ToArray();
        }
        else
        {
            return null;
        }
    }
```

```vb
    ' Get all the text in a slide.
    Public Function GetAllTextInSlide(ByVal presentationFile As String, ByVal slideIndex As Integer) As String()
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, False)
            ' Pass the presentation and the slide index
            ' to the next GetAllTextInSlide method, and
            ' then return the array of strings it returns. 
            Return GetAllTextInSlide(presentationDocument, slideIndex)
        End Using
    End Function
    Public Function GetAllTextInSlide(ByVal presentationDocument As PresentationDocument, ByVal slideIndex As Integer) As String()
        ' Verify that the presentation document exists.
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Verify that the slide index is not out of range.
        If slideIndex < 0 Then
            Throw New ArgumentOutOfRangeException("slideIndex")
        End If

        ' Get the presentation part of the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Verify that the presentation part and presentation exist.
        If presentationPart IsNot Nothing AndAlso presentationPart.Presentation IsNot Nothing Then
            ' Get the Presentation object from the presentation part.
            Dim presentation As Presentation = presentationPart.Presentation

            ' Verify that the slide ID list exists.
            If presentation.SlideIdList IsNot Nothing Then
                ' Get the collection of slide IDs from the slide ID list.
                Dim slideIds = presentation.SlideIdList.ChildElements

                ' If the slide ID is in range...
                If slideIndex < slideIds.Count Then
                    ' Get the relationship ID of the slide.
                    Dim slidePartRelationshipId As String = (TryCast(slideIds(slideIndex), SlideId)).RelationshipId

                    ' Get the specified slide part from the relationship ID.
                    Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slidePartRelationshipId), SlidePart)

                    ' Pass the slide part to the next method, and
                    ' then return the array of strings that method
                    ' returns to the previous method.
                    Return GetAllTextInSlide(slidePart)
                End If
            End If
        End If

        ' Else, return null.
        Return Nothing
    End Function
    Public Function GetAllTextInSlide(ByVal slidePart As SlidePart) As String()
        ' Verify that the slide part exists.
        If slidePart Is Nothing Then
            Throw New ArgumentNullException("slidePart")
        End If

        ' Create a new linked list of strings.
        Dim texts As New LinkedList(Of String)()

        ' If the slide exists...
        If slidePart.Slide IsNot Nothing Then
            ' Iterate through all the paragraphs in the slide.
            For Each paragraph In slidePart.Slide.Descendants(Of DocumentFormat.OpenXml.Drawing.Paragraph)()
                ' Create a new string builder.                    
                Dim paragraphText As New StringBuilder()

                ' Iterate through the lines of the paragraph.
                For Each Text In paragraph.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)()
                    ' Append each line to the previous lines.
                    paragraphText.Append(Text.Text)
                Next Text

                If paragraphText.Length > 0 Then
                    ' Add each paragraph to the linked list.
                    texts.AddLast(paragraphText.ToString())
                End If
            Next paragraph
        End If

        If texts.Count > 0 Then
            ' Return an array of strings.
            Return texts.ToArray()
        Else
            Return Nothing
        End If
    End Function
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
