---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 82deb499-7479-474d-9d89-c4847e6f3649
title: Working with presentations
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Working with presentations

This topic discusses the Open XML SDK for Office [Presentation](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.presentation) class and how it relates to
the Open XML File Format PresentationML schema. For more information
about the overall structure of the parts and elements that make up a
PresentationML document, see [Structure of a
PresentationML document](structure-of-a-presentationml-document.md).


---------------------------------------------------------------------------------
## Presentations in PresentationML
The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML \<presentation\>
element used to represent a presentation in a PresentationML document as
follows:

This element specifies within it fundamental presentation-wide
properties.

**Example:** Consider the following presentation with a single slide master
and two slides. In addition to these commonly used elements there can
also be the specification of other properties such as slide size, notes
size and default text styles.

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p="">  
    <p:sldMasterIdLst>  
        <p:sldMasterId id="2147483648" r:id="rId1">  
    </p:sldMasterIdLst>  
    <p:sldIdLst>  
        <p:sldId id="256" r:id="rId3"/>  
        <p:sldId id="257" r:id="rId4"/>  
    </p:sldIdLst>  
    <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>  
    <p:notesSz cx="6858000" cy="9144000"/>  
    <p:defaultTextStyle>  
        …  
    </p:defaultTextStyle>  
</p:presentation>
```

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The \<presentation\> element typically contains child elements that list
slide masters, slides, and custom slide shows contained within the
presentation. In addition, it also commonly contains elements that
specify other properties of the presentation, such as slide size, notes
size, and default text styles.

The \<presentation\> element is the root element of the PresentationML
Presentation part. For more information about the overall structure of
the parts and elements that make up a PresentationML document, see
[Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

The following table lists some of the most common child elements of the
\<presentation\> element used when working with presentations and the
Open XML SDK classes that correspond to them.


| **PresentationML Element** |                                                     **Open XML SDK Class**                                                      |
|----------------------------|-------------------------------------------------------------------------------------------------------------------------------------|
|     \<sldMasterIdLst\>     |   [SlideMasterIdList](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.slidemasteridlist)   |
|      \<sldMasterId\>       |       [SlideMasterId](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.slidemasterid)       |
|        \<sldIdLst\>        |         [SlideIdList](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.slideidlist)         |
|         \<sldId\>          |             [SlideId](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.slideid)             |
|    \<notesMasterIdLst\>    |   [NotesMasterIdList](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.notesmasteridlist)   |
|   \<handoutMasterIdLst\>   | [HandoutMasterIdList](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.handoutmasteridlist) |
|      \<custShowLst\>       |      [CustomShowList](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.customshowlist)      |
|         \<sldSz\>          |           [SlideSize](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.slidesize)           |
|        \<notesSz\>         |           [NotesSize](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.notessize)           |
|    \<defaultTextStyle\>    |    [DefaultTextStyle](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.defaulttextstyle)    |

--------------------------------------------------------------------------------
## Open XML SDK Presentation Class
The Open XML SDK**Presentation** class
represents the \<presentation\> element defined in the Open XML File
Format schema for PresentationML documents. Use the **Presentation** class to manipulate an individual
\<presentation\> element in a PresentationML document.

Classes commonly associated with the **Presentation** class are shown in the following
sections.

### SlideMasterIdList Class

All slides that share the same master inherit the same layout from that
master. The **SlideMasterIdList** class
corresponds to the \<sldMasterIdList\> element. The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML \<sldMasterIdList\>
element used to represent a slide master ID list in a PresentationML
document as follows:

This element specifies a list of identification information for the
slide master slides that are available within the corresponding
presentation. A slide master is a slide that is specifically designed to
be a template for all related child layout slides.

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### SlideMasterId Class

The **SlideMasterId** class corresponds to the
\<sldMasterId\> element. The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML \<sldMasterId\>
element used to represent a slide master ID in a PresentationML document
as follows:

This element specifies a slide master that is available within the
corresponding presentation. A slide master is a slide that is
specifically designed to be a template for all related child layout
slides.

**Example**: Consider the following specification of a slide master within
a presentation

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p=""
embedTrueTypeFonts="1">  
    …  
    <p:sldMasterIdLst>  
        <p:sldMasterId id="2147483648" r:id="rId1"/>  
    </p:sldMasterIdLst>  
    …  
</p:presentation>
```

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### SlideIdList Class

The **SlideIdList** class corresponds to the
\<sldIdLst\> element. The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML \<sldIdLst\> element
used to represent a slide ID list in a PresentationML document as
follows:

This element specifies a list of identification information for the
slides that are available within the corresponding presentation. A slide
contains the information that is specific to a single slide such as
slide-specific shape and text information.

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### SlideId Class

The **SlideId** class corresponds to the
\<sldId\> element. The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML \<sldId\> element
used to represent a slide ID in a PresentationML document as follows:

This element specifies a presentation slide that is available within the
corresponding presentation. A slide contains the information that is
specific to a single slide such as slide-specific shape and text
information.

**Example**: Consider the following specification of a slide master within
a presentation

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p=""
embedTrueTypeFonts="1">  
    …  
    <p:sldIdLst>  
        <p:sldId id="256" r:id="rId3"/>  
        <p:sldId id="257" r:id="rId4"/>  
        <p:sldId id="258" r:id="rId5"/>  
        <p:sldId id="259" r:id="rId6"/>  
        <p:sldId id="260" r:id="rId7"/>  
    </p:sldIdLst>  
    ...  
</p:presentation>
```

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### NotesMasterIdList Class

The **NotesMasterIdList** class corresponds to
the \<notesMasterIdLst\> element. The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML \<notesMasterIdLst\>
element used to represent a notes master ID list in a PresentationML
document as follows:

This element specifies a list of identification information for the
notes master slides that are available within the corresponding
presentation. A notes master is a slide that is specifically designed
for the printing of the slide along with any attached notes.

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### HandoutMasterIdList Class

The **HandoutMasterIdList** class corresponds
to the \<handoutMasterIdLst\> element. The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML
\<handoutMasterIdLst\> element used to represent a handout master ID
list in a PresentationML document as follows:

This element specifies a list of identification information for the
handout master slides that are available within the corresponding
presentation. A handout master is a slide that is specifically designed
for printing as a handout.

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### CustomShowList Class

The **CustomShowList** class corresponds to the
\<custShowLst\> element. The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML \<custShowLst\>
element used to represent a custom show list in a PresentationML
document as follows:

This element specifies a list of all custom shows that are available
within the corresponding presentation. A custom show is a defined slide
sequence that allows for the displaying of the slides with the
presentation in any arbitrary order.

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### SlideSize Class

The **SlideSize** class corresponds to the
\<sldSz\> element. The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML \<sldSz\> element
used to represent presentation slide size in a PresentationML document
as follows:

This element specifies the size of the presentation slide surface.
Objects within a presentation slide can be specified outside these
extents, but this is the size of background surface that is shown when
the slide is presented or printed.

**Example**: Consider the following specifying of the size of a
presentation slide.

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p=""
embedTrueTypeFonts="1">  
    …  
    <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>  
    …  
</p:presentation>  
```

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### NotesSize Class

The **NotesSize** class corresponds to the
\<notesSz\> element. The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML \<notesSz\> element
used to represent notes slide size in a PresentationML document as
follows:

This element specifies the size of slide surface used for notes slides
and handout slides. Objects within a notes slide can be specified
outside these extents, but the notes slide has a background surface of
the specified size when presented or printed. This element is intended
to specify the region to which content is fitted in any special format
of printout the application might choose to generate, such as an outline
handout.

**Example**: Consider the following specifying of the size of a notes
slide.

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p=""
embedTrueTypeFonts="1">  
    …  
    <p:notesSz cx="9144000" cy="6858000"/>  
    …  
</p:presentation>
```

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### DefaultTextStyle Class

The DefaultTextStyle class corresponds to the \<defaultTextStyle\>
element. The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML \<defaultTextStyle\>
element used to represent default text style in a PresentationML
document as follows:

This element specifies the default text styles that are to be used
within the presentation. The text style defined here can be referenced
when inserting a new slide if that slide is not associated with a master
slide or if no styling information has been otherwise specified for the
text within the presentation slide.

© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]


--------------------------------------------------------------------------------
## Working with the Presentation Class
As shown in the Open XML SDK code example that follows, every instance
of the **Presentation** class is associated
with an instance of the [PresentationPart](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.presentationpart) class, which represents a
presentation part, one of the required parts of a PresentationML
presentation file package.

The **Presentation** class, which represents
the \<presentation\> element, is therefore also associated with a series
of other classes that represent the child elements of the
\<presentation\> element. Among these classes, as shown in the following
code example, are the **SlideMasterIdList**,
**SlideIdList**, **SlideSize**, **NotesSize**, and **DefaultTextStyle** classes.


--------------------------------------------------------------------------------
## Open XML SDK Code Example
The following code example from the article [How to: Create a presentation document by providing a file name](how-to-create-a-presentation-document-by-providing-a-file-name.md) uses the [Create(String, PresentationDocumentType)](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.presentationdocument.create)
method of the [PresentationDocument](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.presentationdocument) class of the Open XML
SDK to create an instance of that same class that has the specified
name and file path. Then it uses the [AddPresentationPart()](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.presentationdocument.addpresentationpart) method to add an
instance of the [PresentationPart](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.presentationpart) class to the document
file. Next, it creates an instance of the [Presentation](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.presentation) class that represents the
presentation. It passes a reference to the **PresentationPart** class instance to the
**CreatePresentationParts** procedure, which creates the other required
parts of the presentation file. The **CreatePresentation** procedure
cleans up by closing the **PresentationDocument** class instance that it
opened previously.

The **CreatePresentationParts** procedure creates instances of the **SlideMasterIdList**, **SlideIdList**, **SlideSize**, **NotesSize**, and **DefaultTextStyle** classes and appends them to the
presentation.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/working_with_presentations/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/working_with_presentations/vb/Program.vb)]

---------------------------------------------------------------------------------
## Resulting PresentationML
When the Open XML SDK code is run, the following XML is written to the
PresentationML document referenced in the code.

```xml
    <?xml version="1.0" encoding="utf-8" ?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
     <p:sldMasterIdLst>
      <p:sldMasterId id="2147483648" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
     </p:sldMasterIdLst>
     <p:sldIdLst>
      <p:sldId id="256" r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
     </p:sldIdLst>
     <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
     <p:notesSz cx="6858000" cy="9144000"/>
     <p:defaultTextStyle/>
    </p:presentation>
```

--------------------------------------------------------------------------------
## See also


[About the Open XML SDK for Office](../about-the-open-xml-sdk.md)  

[How to: Create a presentation document by providing a file name](how-to-create-a-presentation-document-by-providing-a-file-name.md)  

[How to: Insert a new slide into a presentation](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Retrieve the number of slides in a presentation document](how-to-retrieve-the-number-of-slides-in-a-presentation-document.md)  

[How to: Apply a theme to a presentation](how-to-apply-a-theme-to-a-presentation.md)  
