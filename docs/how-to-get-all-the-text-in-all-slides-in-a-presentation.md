---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: debad542-5915-45ad-a71c-eeb95b40ec1a
title: 'How to: Get all the text in all slides in a presentation (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Get all the text in all slides in a presentation (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 to get
all of the text in all of the slides in a presentation programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using System.Linq;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Presentation;
    using A = DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml;
    using System.Text;
```

```vb
    Imports System
    Imports System.Linq
    Imports System.Collections.Generic
    Imports A = DocumentFormat.OpenXml.Drawing
    Imports DocumentFormat.OpenXml.Presentation
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml
    Imports System.Text
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
the <span class="term">presentationFile</span> parameter is a string
that represents the path for the file from which you want to open the
document.

```csharp
    // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
    {
        // Insert other code here.
    }
```

```vb
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
class="keyword">using</span> statement, in this case <span
class="keyword">presentationDocument</span>.


--------------------------------------------------------------------------------

The basic document structure of a <span
class="keyword">PresentationML</span> document consists of a number of
parts, among which is the main part that contains the presentation
definition. The following text from the [ISO/IEC
29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification
introduces the overall form of a <span
class="keyword">PresentationML</span> package.

A **PresentationML** package's main part starts
with a presentation root element. That element contains a presentation,
which, in turn, refers to a <span class="term">slide</span> list, a
<span class="term">slide master</span> list, a <span class="term">notes
master</span> list, and a <span class="term">handout master</span> list.
The slide list refers to all of the slides in the presentation; the
slide master list refers to the entire slide masters used in the
presentation; the notes master contains information about the formatting
of notes pages; and the handout master describes how a handout looks.

A <span class="term">handout</span> is a printed set of slides that can
be provided to an <span class="term">audience</span> for future
reference.

As well as text and graphics, each slide can contain <span
class="term">comments</span> and <span class="term">notes</span>, can
have a <span class="term">layout</span>, and can be part of one or more
<span class="term">custom presentations</span>. (A comment is an
annotation intended for the person maintaining the presentation slide
deck. A note is a reminder or piece of text intended for the presenter
or the audience.)

Other features that a PresentationML document can contain include the
following: <span class="term">animation</span>, <span
class="term">audio</span>, <span class="term">video</span>, and <span
class="term">transitions</span> between slides.

A PresentationML document is not stored as one large body in a single
part. Instead, the elements that implement certain groupings of
functionality are stored in separate parts. For example, all comments in
a document are stored in one comment part while each slide has its own
part.

© ISO/IEC29500: 2008.

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

The sample code starts by counting the number of slides in the
presentation. It does that by using two overloads of the method <span
class="keyword">CountSlides</span>. The first overload uses a string
parameter and the second overload uses a <span
class="keyword">PresentationDocument</span> parameter. In the first
**CountSlides** method, the sample code opens
the presentation document in the **using**
statement. Then it passes the <span
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
the **Count** method, which belongs to <span
class="keyword">SlideParts</span>, the code gets the <span
class="keyword">slidesCount</span> and returns it.

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

After counting the number of slides, the code uses the method <span
class="keyword">GetSlideIdAndText</span>to get the content of all the
slides. It starts with getting the relationship ID of the first slide,
and then gets the slide part from the relationship ID.

```csharp
    // Get the relationship ID of the first slide.
    PresentationPart part = ppt.PresentationPart;
    OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

    string relId = (slideIds[index] as SlideId).RelationshipId;

    // Get the slide part from the relationship ID.
    SlidePart slide = (SlidePart) part.GetPartById(relId);
```

```vb
    ' Get the relationship ID of the first slide.
    Dim part As PresentationPart = ppt.PresentationPart
    Dim slideIds As OpenXmlElementList = part.Presentation.SlideIdList.ChildElements

    Dim relId As String = TryCast(slideIds(index), SlideId).RelationshipId

    ' Get the slide part from the relationship ID.
    Dim slide As SlidePart = DirectCast(part.GetPartById(relId), SlidePart)
```

The code then declares a **StringBuilder**
object to store the inner text of the slide. It then goes through all
the slides and appends the text in each one to the <span
class="keyword">StringBuilder</span> object.

```csharp
    // Build a StringBuilder object.
    StringBuilder paragraphText = new StringBuilder();

    // Get the inner text of the slide:
    IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
    foreach (A.Text text in texts)
    {
        paragraphText.Append(text.Text);
    }
    sldText = paragraphText.ToString();
```

```vb
    ' Build a StringBuilder object.
    Dim paragraphText As New StringBuilder()

    ' Get the inner text of the slide:
    Dim texts As IEnumerable(Of A.Text) = slide.Slide.Descendants(Of A.Text)()
    For Each text As A.Text In texts
        paragraphText.Append(text.Text)
    Next
    sldText = paragraphText.ToString()
```

--------------------------------------------------------------------------------

The following code gets all the text in all the slides in a specific
presentation file. For example, you can enter the name of the
presentation file from the keyboard, and then use a <span
class="keyword">foreach</span> loop in your program to get the array of
strings returned by the method <span
class="keyword">GetSlideIdAndText</span> as shown in the following
example.

```csharp
    Console.Write("Please enter a presentation file name without extension: ");
    string fileName = Console.ReadLine();
    string file = @"C:\Users\Public\Documents\" + fileName + ".pptx";
    int numberOfSlides = CountSlides(file);
    System.Console.WriteLine("Number of slides = {0}", numberOfSlides);
    string slideText;
    for (int i = 0; i < numberOfSlides; i++)
    {
        GetSlideIdAndText(out slideText, file, i);
        System.Console.WriteLine("Slide #{0} contains: {1}", i + 1, slideText);
    }
    System.Console.ReadKey();
```

```vb
    Console.Write("Please enter a presentation file name without extension: ")
    Dim fileName As String = System.Console.ReadLine()
    Dim file As String = "C:\Users\Public\Documents\" + fileName + ".pptx"
    Dim numberOfSlides As Integer = CountSlides(file)
    System.Console.WriteLine("Number of slides = {0}", numberOfSlides)
    Dim slideText As String = Nothing
    For i As Integer = 0 To numberOfSlides - 1
        GetSlideIdAndText(slideText, file, i)
        System.Console.WriteLine("Slide #{0} contains: {1}", i + 1, slideText)
    Next
    System.Console.ReadKey()
```

The following is the complete sample code in both C\# and Visual Basic.

```csharp
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

    public static void GetSlideIdAndText(out string sldText, string docName, int index)
    {
        using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
        {
            // Get the relationship ID of the first slide.
            PresentationPart part = ppt.PresentationPart;
            OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

            string relId = (slideIds[index] as SlideId).RelationshipId;

            // Get the slide part from the relationship ID.
            SlidePart slide = (SlidePart) part.GetPartById(relId);

            // Build a StringBuilder object.
            StringBuilder paragraphText = new StringBuilder();

            // Get the inner text of the slide:
            IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
            foreach (A.Text text in texts)
            {
                paragraphText.Append(text.Text);
            }
            sldText = paragraphText.ToString();
        }              
    }
```

```vb
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
        ' Get the slide count from the SlideParts.
        If presentationPart IsNot Nothing Then
            slidesCount = presentationPart.SlideParts.Count()
        End If
        ' Return the slide count to the previous method.
        Return slidesCount
    End Function

    Public Sub GetSlideIdAndText(ByRef sldText As String, ByVal docName As String, ByVal index As Integer)
        Using ppt As PresentationDocument = PresentationDocument.Open(docName, False)
            ' Get the relationship ID of the first slide.
            Dim part As PresentationPart = ppt.PresentationPart
            Dim slideIds As OpenXmlElementList = part.Presentation.SlideIdList.ChildElements

            Dim relId As String = TryCast(slideIds(index), SlideId).RelationshipId

            ' Get the slide part from the relationship ID.
            Dim slide As SlidePart = DirectCast(part.GetPartById(relId), SlidePart)
            ' Build a StringBuilder object.
            Dim paragraphText As New StringBuilder()

            ' Get the inner text of the slide:
            Dim texts As IEnumerable(Of A.Text) = slide.Slide.Descendants(Of A.Text)()
            For Each text As A.Text In texts
                paragraphText.Append(text.Text)
            Next
            sldText = paragraphText.ToString()
        End Using
    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library
reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
