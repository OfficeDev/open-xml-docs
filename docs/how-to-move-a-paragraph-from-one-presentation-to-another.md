---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: ef817bef-27cd-4c2a-acf3-b7bba17e6e1e
title: 'How to: Move a paragraph from one presentation to another (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
localization_priority: Normal
---
# Move a paragraph from one presentation to another (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to move a paragraph from one presentation to another presentation
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml.Presentation;
    using DocumentFormat.OpenXml.Packaging;
    using Drawing = DocumentFormat.OpenXml.Drawing;
```

```vb
    Imports System.Linq
    Imports DocumentFormat.OpenXml.Presentation
    Imports DocumentFormat.OpenXml.Packaging
    Imports Drawing = DocumentFormat.OpenXml.Drawing
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

This following XML code segment represents a presentation that contains
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


## Structure of the Shape Text Body

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the structure of this element.

> This element specifies the existence of text to be contained within
> the corresponding shape. All visible text and visible text related
> properties are contained within this element. There can be multiple
> paragraphs and within paragraphs multiple runs of text.
> 
> © ISO/IEC29500: 2008.

The following table lists the child elements of the shape text body and
the description of each.

| Child Element | Description |
|---|---|
| bodyPr | Body Properties |
| lstStyle | Text List Styles |
| p | Text Paragraphs |

The following XML Schema fragment defines the contents of this element:

```xml
    <complexType name="CT_TextBody">
       <sequence>
           <element name="bodyPr" type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
           <element name="lstStyle" type="CT_TextListStyle" minOccurs="0" maxOccurs="1"/>
           <element name="p" type="CT_TextParagraph" minOccurs="1" maxOccurs="unbounded"/>
       </sequence>
    </complexType>
```

## How the Sample Code Works

The code in this topic consists of two methods, **MoveParagraphToPresentation** and **GetFirstSlide**. The first method takes two string
parameters: one that represents the source file, which contains the
paragraph to move, and one that represents the target file, to which the
paragraph is moved. The method opens both presentation files and then
calls the **GetFirstSlide** method to get the
first slide in each file. It then gets the first **TextBody** shape in each slide and the first
paragraph in the source shape. It performs a **deep
clone** of the source paragraph, copying not only the source **Paragraph** object itself, but also everything
contained in that object, including its text. It then inserts the cloned
paragraph in the target file and removes the source paragraph from the
source file, replacing it with a placeholder paragraph. Finally, it
saves the modified slides in both presentations.

```csharp
    // Moves a paragraph range in a TextBody shape in the source document
    // to another TextBody shape in the target document.
    public static void MoveParagraphToPresentation(string sourceFile, string targetFile)
    {
        // Open the source file as read/write.
        using (PresentationDocument sourceDoc = PresentationDocument.Open(sourceFile, true))
        {
            // Open the target file as read/write.
            using (PresentationDocument targetDoc = PresentationDocument.Open(targetFile, true))
            {
                // Get the first slide in the source presentation.
                SlidePart slide1 = GetFirstSlide(sourceDoc);

                // Get the first TextBody shape in it.
                TextBody textBody1 = slide1.Slide.Descendants<TextBody>().First();

                // Get the first paragraph in the TextBody shape.
                // Note: "Drawing" is the alias of namespace DocumentFormat.OpenXml.Drawing
                Drawing.Paragraph p1 = textBody1.Elements<Drawing.Paragraph>().First();

                // Get the first slide in the target presentation.
                SlidePart slide2 = GetFirstSlide(targetDoc);

                // Get the first TextBody shape in it.
                TextBody textBody2 = slide2.Slide.Descendants<TextBody>().First();

                // Clone the source paragraph and insert the cloned. paragraph into the target TextBody shape.
                // Passing "true" creates a deep clone, which creates a copy of the 
                // Paragraph object and everything directly or indirectly referenced by that object.
                textBody2.Append(p1.CloneNode(true));

                // Remove the source paragraph from the source file.
                textBody1.RemoveChild<Drawing.Paragraph>(p1);

                // Replace the removed paragraph with a placeholder.
                textBody1.AppendChild<Drawing.Paragraph>(new Drawing.Paragraph());

                // Save the slide in the source file.
                slide1.Slide.Save();

                // Save the slide in the target file.
                slide2.Slide.Save();
            }
        }
    }
```

```vb
    ' Moves a paragraph range in a TextBody shape in the source document
    ' to another TextBody shape in the target document.
    Public Shared Sub MoveParagraphToPresentation(ByVal sourceFile As String, ByVal targetFile As String)
        ' Open the source file as read/write.
        Using sourceDoc As PresentationDocument = PresentationDocument.Open(sourceFile, True)
            ' Open the target file as read/write.
            Using targetDoc As PresentationDocument = PresentationDocument.Open(targetFile, True)
                ' Get the first slide in the source presentation.
                Dim slide1 As SlidePart = GetFirstSlide(sourceDoc)

                ' Get the first TextBody shape in it.
                Dim textBody1 As TextBody = slide1.Slide.Descendants(Of TextBody)().First()

                ' Get the first paragraph in the TextBody shape.
                ' Note: "Drawing" is the alias of namespace DocumentFormat.OpenXml.Drawing
                Dim p1 As Drawing.Paragraph = textBody1.Elements(Of Drawing.Paragraph)().First()

                ' Get the first slide in the target presentation.
                Dim slide2 As SlidePart = GetFirstSlide(targetDoc)

                ' Get the first TextBody shape in it.
                Dim textBody2 As TextBody = slide2.Slide.Descendants(Of TextBody)().First()

                ' Clone the source paragraph and insert the cloned. paragraph into the target TextBody shape.
                ' Passing "true" creates a deep clone, which creates a copy of the 
                ' Paragraph object and everything directly or indirectly referenced by that object.
                textBody2.Append(p1.CloneNode(True))

                ' Remove the source paragraph from the source file.
                textBody1.RemoveChild(Of Drawing.Paragraph)(p1)

                ' Replace the removed paragraph with a placeholder.
                textBody1.AppendChild(Of Drawing.Paragraph)(New Drawing.Paragraph())

                ' Save the slide in the source file.
                slide1.Slide.Save()

                ' Save the slide in the target file.
                slide2.Slide.Save()
            End Using
        End Using
    End Sub
```

The **GetFirstSlide** method takes the **PresentationDocument** object passed in, gets its
presentation part, and then gets the ID of the first slide in its slide
list. It then gets the relationship ID of the slide, gets the slide part
from the relationship ID, and returns the slide part to the calling
method.

```csharp
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
    ' Get the slide part of the first slide in the presentation document.
    Public Shared Function GetFirstSlide(ByVal presentationDocument As PresentationDocument) As SlidePart
        ' Get relationship ID of the first slide
        Dim part As PresentationPart = presentationDocument.PresentationPart
        Dim slideId As SlideId = part.Presentation.SlideIdList.GetFirstChild(Of SlideId)()
        Dim relId As String = slideId.RelationshipId

        ' Get the slide part by the relationship ID.
        Dim slidePart As SlidePart = CType(part.GetPartById(relId), SlidePart)

        Return slidePart
    End Function
```

## Sample Code

By using this sample code, you can move a paragraph from one
presentation to another. In your program, you can use the following call
to the **MoveParagraphToPresentation** method
to move the first paragraph from the presentation file "Myppt4.pptx" to
the presentation file "Myppt12.pptx".

```csharp
    string sourceFile = @"C:\Users\Public\Documents\Myppt4.pptx";
    string targetFile = @"C:\Users\Public\Documents\Myppt12.pptx";
    MoveParagraphToPresentation(sourceFile, targetFile);
```

```vb
    Dim sourceFile As String = "C:\Users\Public\Documents\Myppt4.pptx"
    Dim targetFile As String = "C:\Users\Public\Documents\Myppt12.pptx"
    MoveParagraphToPresentation(sourceFile, targetFile)
```

After you run the program take a look on the content of both the source
and the target files to see the moved paragraph.

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Moves a paragraph range in a TextBody shape in the source document
    // to another TextBody shape in the target document.
    public static void MoveParagraphToPresentation(string sourceFile, string targetFile)
    {
        // Open the source file as read/write.
        using (PresentationDocument sourceDoc = PresentationDocument.Open(sourceFile, true))
        {
            // Open the target file as read/write.
            using (PresentationDocument targetDoc = PresentationDocument.Open(targetFile, true))
            {
                // Get the first slide in the source presentation.
                SlidePart slide1 = GetFirstSlide(sourceDoc);

                // Get the first TextBody shape in it.
                TextBody textBody1 = slide1.Slide.Descendants<TextBody>().First();

                // Get the first paragraph in the TextBody shape.
                // Note: "Drawing" is the alias of namespace DocumentFormat.OpenXml.Drawing
                Drawing.Paragraph p1 = textBody1.Elements<Drawing.Paragraph>().First();

                // Get the first slide in the target presentation.
                SlidePart slide2 = GetFirstSlide(targetDoc);

                // Get the first TextBody shape in it.
                TextBody textBody2 = slide2.Slide.Descendants<TextBody>().First();

                // Clone the source paragraph and insert the cloned. paragraph into the target TextBody shape.
                // Passing "true" creates a deep clone, which creates a copy of the 
                // Paragraph object and everything directly or indirectly referenced by that object.
                textBody2.Append(p1.CloneNode(true));

                // Remove the source paragraph from the source file.
                textBody1.RemoveChild<Drawing.Paragraph>(p1);

                // Replace the removed paragraph with a placeholder.
                textBody1.AppendChild<Drawing.Paragraph>(new Drawing.Paragraph());

                // Save the slide in the source file.
                slide1.Slide.Save();

                // Save the slide in the target file.
                slide2.Slide.Save();
            }
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
    ' Moves a paragraph range in a TextBody shape in the source document
    ' to another TextBody shape in the target document.
    Public Sub MoveParagraphToPresentation(ByVal sourceFile As String, ByVal targetFile As String)

        ' Open the source file.
        Dim sourceDoc As PresentationDocument = PresentationDocument.Open(sourceFile, True)

        ' Open the target file.
        Dim targetDoc As PresentationDocument = PresentationDocument.Open(targetFile, True)

        ' Get the first slide in the source presentation.
        Dim slide1 As SlidePart = GetFirstSlide(sourceDoc)

        ' Get the first TextBody shape in it.
        Dim textBody1 As TextBody = slide1.Slide.Descendants(Of TextBody).First()

        ' Get the first paragraph in the TextBody shape.
        ' Note: Drawing is the alias of the namespace DocumentFormat.OpenXml.Drawing
        Dim p1 As Drawing.Paragraph = textBody1.Elements(Of Drawing.Paragraph).First()

        ' Get the first slide in the target presentation.
        Dim slide2 As SlidePart = GetFirstSlide(targetDoc)

        ' Get the first TextBody shape in it.
        Dim textBody2 As TextBody = slide2.Slide.Descendants(Of TextBody).First()

        ' Clone the source paragraph and insert the cloned paragraph into the target TextBody shape.
        textBody2.Append(p1.CloneNode(True))

        ' Remove the source paragraph from the source file.
        textBody1.RemoveChild(Of Drawing.Paragraph)(p1)

        ' Replace it with an empty one, because a paragraph is required for a TextBody shape.
        textBody1.AppendChild(Of Drawing.Paragraph)(New Drawing.Paragraph())

        ' Save the slide in the source file.
        slide1.Slide.Save()

        ' Save the slide in the target file.
        slide2.Slide.Save()

    End Sub
    ' Get the slide part of the first slide in the presentation document.
    Public Function GetFirstSlide(ByVal presentationDoc As PresentationDocument) As SlidePart

        ' Get relationship ID of the first slide.
        Dim part As PresentationPart = presentationDoc.PresentationPart
        Dim slideId As SlideId = part.Presentation.SlideIdList.GetFirstChild(Of SlideId)()
        Dim relId As String = slideId.RelationshipId

        ' Get the slide part by the relationship ID.
        Dim slidePart As SlidePart = CType(part.GetPartById(relId), SlidePart)

        Return slidePart

    End Function
```

## See also

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
