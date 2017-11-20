---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 8dc8a6ac-aa9e-47cc-b45e-e128fcec3c57
title: 'How to: Open a presentation document for read-only access (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Open a presentation document for read-only access (Open XML SDK)

This topic describes how to use the classes in the Open XML SDK 2.5 for
Office to programmatically open a presentation document for read-only
access.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Presentation;
    using A = DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml;
    using System.Text;
```

```vb
    Imports System
    Imports System.Collections.Generic
    Imports DocumentFormat.OpenXml.Presentation
    Imports A = DocumentFormat.OpenXml.Drawing
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml
    Imports System.Text
```

## How to Open a File for Read-Only Access

You may want to open a presentation document to read the slides. You
might want to extract information from a slide, copy a slide to a slide
library, or list the titles of the slides. In such cases you want to do
so in a way that ensures the document remains unchanged. You can do that
by opening the document for read-only access. This How-To topic
discusses several ways to programmatically open a read-only presentation
document.


## Create an Instance of the PresentationDocument Class 

In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.PresentationDocument"><span
class="nolink">PresentationDocument</span></span> class represents a
presentation document package. To work with a presentation document,
first create an instance of the <span
class="keyword">PresentationDocument</span> class, and then work with
that instance. To create the class instance from the document call one
of the <span sdata="cer"
target="Overload:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open"><span
class="nolink">Open</span></span> methods. Several Open methods are
provided, each with a different signature. The following table contains
a subset of the overloads for the **Open**
method that you can use to open the package.

| Name | Description |
|---|---|
| Open(String, Boolean) | Create a new instance of the **PresentationDocument** class from the specified file. |
| Open(Stream, Boolean) | Create a new instance of the **PresentationDocument** class from the I/O stream. |
| Open(Package) | Create a new instance of the **PresentationDocument** class from the specified package. |


The previous table includes two **Open**
methods that accept a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read-only access,
specify the value **false** for this parameter.

For example, you can open the presentation file as read-only and assign
it to a <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.PresentationDocument"><span
class="nolink">PresentationDocument</span></span> object as shown in the
following **using** statement. In this code,
the <span class="term">presentationFile</span> parameter is a string
that represents the path of the file from which you want to open the
document.

```csharp
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

You can also use the second overload of the <span
class="keyword">Open</span> method, in the table above, to create an
instance of the **PresentationDocument** class
based on an I/O stream. You might use this approach if you have a
Microsoft SharePoint Foundation 2010 application that uses stream I/O
and you want to use the Open XML SDK 2.5 to work with a document. The
following code segment opens a document based on a stream.

```csharp
    Stream stream = File.Open(strDoc, FileMode.Open);
    using (PresentationDocument presentationDocument =
        PresentationDocument.Open(stream, false)) 
    {
        // Place other code here.
    }
```

```vb
    Dim stream As Stream = File.Open(strDoc, FileMode.Open)
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(stream, False)
        ' Other code goes here.
    End Using
```

Suppose you have an application that employs the Open XML support in the
**System.IO.Packaging** namespace of the .NET
Framework Class Library, and you want to use the Open XML SDK 2.5 to
work with a package read-only. The Open XML SDK 2.5 includes a method
overload that accepts a **Package** as the only
parameter. There is no Boolean parameter to indicate whether the
document should be opened for editing. The recommended approach is to
open the package as read-only prior to creating the instance of the
**PresentationDocument** class. The following
code segment performs this operation.

```csharp
    Package presentationPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read);
    using (PresentationDocument presentationDocument =
        PresentationDocument.Open(presentationPackage))
    {
        // Other code goes here.
    }
```

```vb
    Dim presentationPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationPackage)
        ' Other code goes here.
    End Using
```

## Basic Presentation Document Structure

The basic document structure of a <span
class="keyword">PresentationML</span> document consists of a number of
parts, among which the main part is that contains the presentation
definition. The following text from the [ISO/IEC
29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification
introduces the overall form of a <span
class="keyword">PresentationML</span> package.

> The main part of a **PresentationML** package
> starts with a presentation root element. That element contains a
> presentation, which, in turn, refers to a <span
> class="keyword">slide</span> list, a <span class="term">slide
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
> more <span class="term">custom presentations</span>. A comment is an
> annotation intended for the person maintaining the presentation slide
> deck. A note is a reminder or piece of text intended for the presenter
> or the audience.

> Other features that a **PresentationML**
> document can include the following: <span
> class="term">animation</span>, <span class="term">audio</span>, <span
> class="term">video</span>, and <span class="term">transitions</span>
> between slides.

> A **PresentationML** document is not stored
> as one large body in a single part. Instead, the elements that
> implement certain groupings of functionality are stored in separate
> parts. For example, all comments in a document are stored in one
> comment part while each slide has its own part.

> © ISO/IEC29500: 2008.

This following XML code segment represents a presentation that contains
two slides denoted by the IDs 267 and 256. The <span
class="keyword">ID</span> property specifies the slide identifier that
contains a unique value throughout the presentation. The possible values
for this attribute are from 256 through 2147483647.

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

## How the Sample Code Works

In the sample code, after you open the presentation document in the
**using** statement for read-only access,
instantiate the **PresentationPart**, and open
the slide list. Then you get the relationship ID of the first slide.

```csharp
    // Get the relationship ID of the first slide.
    PresentationPart part = ppt.PresentationPart;
    OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
    string relId = (slideIds[index] as SlideId).RelationshipId;
```

```vb
    ' Get the relationship ID of the first slide.
    Dim part As PresentationPart = ppt.PresentationPart
    Dim slideIds As OpenXmlElementList = part.Presentation.SlideIdList.ChildElements
    Dim relId As String = (TryCast(slideIds(index), SlideId)).RelationshipId
```

From the relationship ID, <span class="term">relId</span>, you get the
slide part, and then the inner text of the slide by building a text
string using **StringBuilder**.

```csharp
    // Get the slide part from the relationship ID.
    SlidePart slide = (SlidePart)part.GetPartById(relId);

    // Build a StringBuilder object.
    StringBuilder paragraphText = new StringBuilder();

    // Get the inner text of the slide.
    IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
    foreach (A.Text text in texts)
    {
        paragraphText.Append(text.Text);
    }
    sldText = paragraphText.ToString();
```

```vb
    ' Get the slide part from the relationship ID.
    Dim slide As SlidePart = CType(part.GetPartById(relId), SlidePart)

    ' Build a StringBuilder object.
    Dim paragraphText As New StringBuilder()

    ' Get the inner text of the slide.
    Dim texts As IEnumerable(Of A.Text) = slide.Slide.Descendants(Of A.Text)()
    For Each text As A.Text In texts
        paragraphText.Append(text.Text)
    Next text
    sldText = paragraphText.ToString()
```

The inner text of the slide, which is an <span
class="keyword">out</span> parameter of the <span
class="keyword">GetSlideIdAndText</span> method, is passed back to the
main method to be displayed.

> [!IMPORTANT]
> This example displays only the text in the presentation file. Non-text parts, such as shapes or graphics, are not displayed.


## Sample Code

The following example opens a presentation file for read-only access and
gets the inner text of a slide at a specified index. To call the method
<span class="term">GetSlideIdAndText</span> pass in the full path of the
presentation document. Also pass in the **out**
parameter <span class="term">sldText</span>, which will be assigned a
value in the method itself, and then you can display its value in the
main program. For example, the following call to the <span
class="keyword">GetSlideIdAndText</span> method gets the inner text in
the second slide in a presentation file named "Myppt13.pptx".

> [!TIP]
> The most expected exception in this program is the **ArgumentOutOfRangeException</span> exception. It could be thrown if, for example, you have a file with two slides, and you wanted to display the text in slide number 4. Therefore, it is best to use a <span class="keyword">try</span> block when you call the <span class="keyword">GetSlideIdAndText** method as shown in the following example.

```csharp
    string file = @"C:\Users\Public\Documents\Myppt13.pptx";
    string slideText;
    int index = 1;
    try
    {
        GetSlideIdAndText(out slideText, file, index);
        Console.WriteLine("The text in the slide #{0} is: {1}", index + 1, slideText);
    }
    catch (ArgumentOutOfRangeException exp)
    {
        Console.WriteLine(exp.Message);
    }
```

```vb
    Dim file As String = "C:\Users\Public\Documents\Myppt13.pptx"
    Dim slideText As String = Nothing
    Dim index As Integer = 1
    Try
        GetSlideIdAndText(slideText, file, index)
        Console.WriteLine("The text in the slide #{0} is: {1}", index + 1, slideText)
    Catch exp As ArgumentOutOfRangeException
        Console.WriteLine(exp.Message)
    End Try
```

The following is the complete code listing in C\# and Visual Basic.

```csharp
    public static void GetSlideIdAndText(out string sldText, string docName, int index)
    {
        using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
        {
            // Get the relationship ID of the first slide.
            PresentationPart part = ppt.PresentationPart;
            OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
            string relId = (slideIds[index] as SlideId).RelationshipId;
            relId = (slideIds[index] as SlideId).RelationshipId;

            // Get the slide part from the relationship ID.
            SlidePart slide = (SlidePart)part.GetPartById(relId);

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
    Public Sub GetSlideIdAndText(ByRef sldText As String, ByVal docName As String, ByVal index As Integer)
        Using ppt As PresentationDocument = PresentationDocument.Open(docName, False)
            ' Get the relationship ID of the first slide.
            Dim part As PresentationPart = ppt.PresentationPart
            Dim slideIds As OpenXmlElementList = part.Presentation.SlideIdList.ChildElements
            Dim relId As String = TryCast(slideIds(index), SlideId).RelationshipId
            relId = TryCast(slideIds(index), SlideId).RelationshipId

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

## See also

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
