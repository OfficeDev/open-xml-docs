---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: d6daf04e-3e45-4570-a184-8f0449c7ab91
title: 'How to: Get all the external hyperlinks in a presentation (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Get all the external hyperlinks in a presentation (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to get all the external hyperlinks in a presentation
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Packaging;
    using Drawing = DocumentFormat.OpenXml.Drawing;
```

```vb
    Imports System
    Imports System.Collections.Generic 
    Imports DocumentFormat.OpenXml.Packaging
    Imports Drawing = DocumentFormat.OpenXml.Drawing
```

--------------------------------------------------------------------------------
## Getting a PresentationDocument Object
In the Open XML SDK, the [PresentationDocument](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.presentationdocument.aspx) class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the
[PresentationDocument.Open(String, Boolean)](https://msdn.microsoft.com/en-us/library/office/cc562287.aspx)
method that uses a file path, and a Boolean value as the second
parameter to specify whether a document is editable. Set this second
parameter to **false** to open the file for
read-only access, or **true** if you want to
open the file for read/write access. In this topic, it is best to open
the file for read-only access to protect the file against accidental
writing. The following **using** statement
opens the file for read-only access. In this code segment, the **fileName** parameter is a string that represents the
path for the file from which you want to open the document.

```csharp
    // Open the presentation file as read-only.
    using (PresentationDocument document = PresentationDocument.Open(fileName, false))
    {
        // Insert other code here.
    }
```

```vb
    ' Open the presentation file as read-only.
    Using document As PresentationDocument = PresentationDocument.Open(fileName, False)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **document**.


--------------------------------------------------------------------------------
## Basic Presentation Document Structure
The basic document structure of a **PresentationML** document consists of the main part
that contains the presentation definition. The following text from the
[ISO/IEC 29500](https://www.iso.org/standard/71691.html)
specification introduces the overall form of a **PresentationML** package.

> A **PresentationML** package's main part
> starts with a presentation root element. That element contains a
> presentation, which, in turn, refers to a <span
> class="keyword">slide** list, a *slide master* list, a *notes
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
content using strongly-typed classes that correspond to **PresentationML** elements. You can find these
classes in the [DocumentFormat.OpenXml.Presentation](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.aspx)
namespace. The following table lists the class names of the classes that
correspond to the **sld**, **sldLayout**, **sldMaster**, and **notesMaster** elements.

| PresentationML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| sld | [Slide](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.slide.aspx) | Presentation Slide. It is the root element of SlidePart. |
| sldLayout | [SlideLayout](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.slidelayout.aspx) | Slide Layout. It is the root element of SlideLayoutPart. |
| sldMaster | [SlideMaster](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.slidemaster.aspx) | Slide Master. It is the root element of SlideMasterPart. |
| notesMaster | [NotesMaster](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.notesmaster.aspx) | Notes Master (or handoutMaster). It is the root element of NotesMasterPart. |


--------------------------------------------------------------------------------
## Structure of the Hyperlink Element
In this how-to code example, you are going to work with external
hyperlinks. Therefore, it is best to familiarize yourself with the
hyperlink element. The following text from the [ISO/IEC
29500](https://www.iso.org/standard/71691.html) specification
introduces the **id** (Hyperlink Target).

> Specifies the ID of the relationship whose target shall be used as the
> target for thishyperlink.

> If this attribute is omitted, then there shall be no external
> hyperlink target for the current hyperlink - a location in the current
> document can still be target via the anchor attribute. If this
> attribute exists, it shall supersede the value in the anchor
> attribute.

> [*Example*: Consider the following <span
> class="keyword">PresentationML** fragment for a hyperlink:

```xml
    <w:hyperlink r:id="rId9">
      <w:r>
        <w:t>http://www.example.com</w:t>
      </w:r>
    </w:hyperlink>
```

> The **id** attribute value of **rId9** specifies that relationship in the
> associated relationship part item with a corresponding Id attribute
> value must be navigated to when this hyperlink is invoked. For
> example, if the following XML is present in the associated
> relationship part item:

```xml
    <Relationships xmlns="…">
      <Relationship Id="rId9" Mode="External"
    Target=http://www.example.com />
    </Relationships>
```

> The target of this hyperlink would therefore be the target of
> relationship **rId9** - in this case,
> http://www.example.com. *end example*]

> The possible values for this attribute are defined by the
> ST\_RelationshipId simple type(§22.8.2.1).

> © ISO/IEC29500: 2008.


--------------------------------------------------------------------------------
## How the Sample Code Works
The sample code in this topic consists of one method that takes as a
parameter the full path of the presentation file. It iterates through
all the slides in the presentation and returns a list of strings that
represent the Universal Resource Identifiers (URIs) of all the external
hyperlinks in the presentation.

```csharp
    // Iterate through all the slide parts in the presentation part.
    foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
    {
        IEnumerable<Drawing.HyperlinkType> links = slidePart.Slide.Descendants<Drawing.HyperlinkType>();

        // Iterate through all the links in the slide part.
        foreach (Drawing.HyperlinkType link in links)
        {

            // Iterate through all the external relationships in the slide part. 
            foreach (HyperlinkRelationship relation in slidePart.HyperlinkRelationships)
            {
                // If the relationship ID matches the link ID…
                if (relation.Id.Equals(link.Id))
                {
                    // Add the URI of the external relationship to the list of strings.
                    ret.Add(relation.Uri.AbsoluteUri);

                }
```

```vb
    ' Iterate through all the slide parts in the presentation part.
    For Each slidePart As SlidePart In document.PresentationPart.SlideParts
        Dim links As IEnumerable(Of Drawing.HyperlinkType) = slidePart.Slide.Descendants(Of Drawing.HyperlinkType)()

        ' Iterate through all the links in the slide part.
        For Each link As Drawing.HyperlinkType In links

            ' Iterate through all the external relationships in the slide part. 
            For Each relation As HyperlinkRelationship In slidePart.HyperlinkRelationships
                ' If the relationship ID matches the link ID…
                If relation.Id.Equals(link.Id) Then
                    ' Add the URI of the external relationship to the list of strings.
                    ret.Add(relation.Uri.AbsoluteUri)
                End If
```

--------------------------------------------------------------------------------
## Sample Code
Following is the complete code sample that you can use to return the
list of all external links in a presentation. You can use the following
loop in your program to call the **GetAllExternalHyperlinksInPresentation** method to
get the list of URIs in your presentation.

```csharp
    string fileName = @"C:\Users\Public\Documents\Myppt7.pptx";
    foreach (string s in GetAllExternalHyperlinksInPresentation(fileName))
        Console.WriteLine(s);
```

```vb
    Dim fileName As String
    fileName = "C:\Users\Public\Documents\Myppt7.pptx"
    For Each s As String In GetAllExternalHyperlinksInPresentation(fileName)
        Console.WriteLine(s)
    Next
```

```csharp
    // Returns all the external hyperlinks in the slides of a presentation.
    public static IEnumerable<String> GetAllExternalHyperlinksInPresentation(string fileName)
    {
        // Declare a list of strings.
        List<string> ret = new List<string>();

        // Open the presentation file as read-only.
        using (PresentationDocument document = PresentationDocument.Open(fileName, false))
        {
            // Iterate through all the slide parts in the presentation part.
            foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
            {
                IEnumerable<Drawing.HyperlinkType> links = slidePart.Slide.Descendants<Drawing.HyperlinkType>();

                // Iterate through all the links in the slide part.
                foreach (Drawing.HyperlinkType link in links)
                {
                    // Iterate through all the external relationships in the slide part. 
                    foreach (HyperlinkRelationship relation in slidePart.HyperlinkRelationships)
                    {
                        // If the relationship ID matches the link ID…
                        if (relation.Id.Equals(link.Id))
                        {
                            // Add the URI of the external relationship to the list of strings.
                            ret.Add(relation.Uri.AbsoluteUri);
                        }
                    }
                }
            }
        }

        // Return the list of strings.
        return ret;
    }
```

```vb
    ' Returns all the external hyperlinks in the slides of a presentation.
    Public Function GetAllExternalHyperlinksInPresentation(ByVal fileName As String) As IEnumerable

        ' Declare a list of strings.
        Dim ret As List(Of String) = New List(Of String)

        ' Open the presentation file as read-only.
        Dim document As PresentationDocument = PresentationDocument.Open(fileName, False)

        Using (document)

            ' Iterate through all the slide parts in the presentation part.
            For Each slidePart As SlidePart In document.PresentationPart.SlideParts
                Dim links As IEnumerable = slidePart.Slide.Descendants(Of Drawing.HyperlinkType)()

                ' Iterate through all the links in the slide part.
                For Each link As Drawing.HyperlinkType In links

                    ' Iterate through all the external relationships in the slide part.
                    For Each relation As HyperlinkRelationship In slidePart.HyperlinkRelationships
                        ' If the relationship ID matches the link ID…
                        If relation.Id.Equals(link.Id) Then

                            ' Add the URI of the external relationship to the list of strings.
                            ret.Add(relation.Uri.AbsoluteUri)
                        End If
                    Next
                Next
            Next


            ' Return the list of strings.
            Return ret

        End Using
    End Function
```

--------------------------------------------------------------------------------
## See also
#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
