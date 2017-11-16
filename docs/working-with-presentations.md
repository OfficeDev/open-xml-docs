---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 82deb499-7479-474d-9d89-c4847e6f3649
title: Working with presentations (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Working with presentations (Open XML SDK)

This topic discusses the Open XML SDK 2.5 for Office <span sdata="cer"
target="T:DocumentFormat.OpenXml.Presentation.Presentation"><span
class="nolink">Presentation</span></span> class and how it relates to
the Open XML File Format PresentationML schema. For more information
about the overall structure of the parts and elements that make up a
PresentationML document, see <span sdata="link">[Structure of a
PresentationML document (Open XML SDK)](structure-of-a-presentationml-document.md)</span>.


---------------------------------------------------------------------------------

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<presentation\>
element used to represent a presentation in a PresentationML document as
follows:

This element specifies within it fundamental presentation-wide
properties.

Example: Consider the following presentation with a single slide master
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

© ISO/IEC29500: 2008.

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
Open XML SDK 2.5 classes that correspond to them.

**PresentationML Element**|**Open XML SDK 2.5 Class**
---|---
\<sldMasterIdLst\>|SlideMasterIdList
\<sldMasterId\>|SlideMasterId
\<sldIdLst\>|SlideIdList
\<sldId\>|SlideId
\<notesMasterIdLst\>|NotesMasterIdList
\<handoutMasterIdLst\>|HandoutMasterIdList
\<custShowLst\>|CustomShowList
\<sldSz\>|SlideSize
\<notesSz\>|NotesSize
\<defaultTextStyle\>|DefaultTextStyle


--------------------------------------------------------------------------------

The Open XML SDK 2.5**Presentation** class
represents the \<presentation\> element defined in the Open XML File
Format schema for PresentationML documents. Use the <span
class="keyword">Presentation</span> class to manipulate an individual
\<presentation\> element in a PresentationML document.

Classes commonly associated with the <span
class="keyword">Presentation</span> class are shown in the following
sections.

### SlideMasterIdList Class

All slides that share the same master inherit the same layout from that
master. The **SlideMasterIdList** class
corresponds to the \<sldMasterIdList\> element. The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<sldMasterIdList\>
element used to represent a slide master ID list in a PresentationML
document as follows:

This element specifies a list of identification information for the
slide master slides that are available within the corresponding
presentation. A slide master is a slide that is specifically designed to
be a template for all related child layout slides.

© ISO/IEC29500: 2008.

### SlideMasterId Class

The **SlideMasterId** class corresponds to the
\<sldMasterId\> element. The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<sldMasterId\>
element used to represent a slide master ID in a PresentationML document
as follows:

This element specifies a slide master that is available within the
corresponding presentation. A slide master is a slide that is
specifically designed to be a template for all related child layout
slides.

Example: Consider the following specification of a slide master within
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

© ISO/IEC29500: 2008.

### SlideIdList Class

The **SlideIdList** class corresponds to the
\<sldIdLst\> element. The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<sldIdLst\> element
used to represent a slide ID list in a PresentationML document as
follows:

This element specifies a list of identification information for the
slides that are available within the corresponding presentation. A slide
contains the information that is specific to a single slide such as
slide-specific shape and text information.

© ISO/IEC29500: 2008.

### SlideId Class

The **SlideId** class corresponds to the
\<sldId\> element. The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<sldId\> element
used to represent a slide ID in a PresentationML document as follows:

This element specifies a presentation slide that is available within the
corresponding presentation. A slide contains the information that is
specific to a single slide such as slide-specific shape and text
information.

Example: Consider the following specification of a slide master within
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

© ISO/IEC29500: 2008.

### NotesMasterIdList Class

The **NotesMasterIdList** class corresponds to
the \<notesMasterIdLst\> element. The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<notesMasterIdLst\>
element used to represent a notes master ID list in a PresentationML
document as follows:

This element specifies a list of identification information for the
notes master slides that are available within the corresponding
presentation. A notes master is a slide that is specifically designed
for the printing of the slide along with any attached notes.

© ISO/IEC29500: 2008.

### HandoutMasterIdList Class

The **HandoutMasterIdList** class corresponds
to the \<handoutMasterIdLst\> element. The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML
\<handoutMasterIdLst\> element used to represent a handout master ID
list in a PresentationML document as follows:

This element specifies a list of identification information for the
handout master slides that are available within the corresponding
presentation. A handout master is a slide that is specifically designed
for printing as a handout.

© ISO/IEC29500: 2008.

### CustomShowList Class

The **CustomShowList** class corresponds to the
\<custShowLst\> element. The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<custShowLst\>
element used to represent a custom show list in a PresentationML
document as follows:

This element specifies a list of all custom shows that are available
within the corresponding presentation. A custom show is a defined slide
sequence that allows for the displaying of the slides with the
presentation in any arbitrary order.

© ISO/IEC29500: 2008.

### SlideSize Class

The **SlideSize** class corresponds to the
\<sldSz\> element. The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<sldSz\> element
used to represent presentation slide size in a PresentationML document
as follows:

This element specifies the size of the presentation slide surface.
Objects within a presentation slide can be specified outside these
extents, but this is the size of background surface that is shown when
the slide is presented or printed.

Example: Consider the following specifying of the size of a
presentation slide.

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p=""
embedTrueTypeFonts="1">  
    …  
    <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>  
    …  
</p:presentation>  
```

© ISO/IEC29500: 2008.

### NotesSize Class

The **NotesSize** class corresponds to the
\<notesSz\> element. The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
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

Example: Consider the following specifying of the size of a notes
slide.

```xml
<p:presentation xmlns:a="" xmlns:r="" xmlns:p=""
embedTrueTypeFonts="1">  
    …  
    <p:notesSz cx="9144000" cy="6858000"/>  
    …  
</p:presentation>
```

© ISO/IEC29500: 2008.

### DefaultTextStyle Class

The DefaultTextStyle class corresponds to the \<defaultTextStyle\>
element. The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<defaultTextStyle\>
element used to represent default text style in a PresentationML
document as follows:

This element specifies the default text styles that are to be used
within the presentation. The text style defined here can be referenced
when inserting a new slide if that slide is not associated with a master
slide or if no styling information has been otherwise specified for the
text within the presentation slide.

© ISO/IEC29500: 2008.


--------------------------------------------------------------------------------

As shown in the Open XML SDK code example that follows, every instance
of the **Presentation** class is associated
with an instance of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.PresentationPart"><span
class="nolink">PresentationPart</span></span> class, which represents a
presentation part, one of the required parts of a PresentationML
presentation file package.

The **Presentation** class, which represents
the \<presentation\> element, is therefore also associated with a series
of other classes that represent the child elements of the
\<presentation\> element. Among these classes, as shown in the following
code example, are the **SlideMasterIdList**,
**SlideIdList**, <span
class="keyword">SlideSize</span>, <span
class="keyword">NotesSize</span>, and <span
class="keyword">DefaultTextStyle</span> classes.


--------------------------------------------------------------------------------

The following code example from the article <span sdata="link">[How to: Create a presentation document by providing a file name (Open XML SDK)](how-to-create-a-presentation-document-by-providing-a-file-name.htm)</span> uses the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.Create(System.String,DocumentFormat.OpenXml.PresentationDocumentType)"><span
class="nolink">Create(String, PresentationDocumentType)</span></span>
method of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.PresentationDocument"><span
class="nolink">PresentationDocument</span></span> class of the Open XML
SDK 2.5 to create an instance of that same class that has the specified
name and file path. Then it uses the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.AddPresentationPart"><span
class="nolink">AddPresentationPart()</span></span> method to add an
instance of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.PresentationPart"><span
class="nolink">PresentationPart</span></span> class to the document
file. Next, it creates an instance of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Presentation.Presentation"><span
class="nolink">Presentation</span></span> class that represents the
presentation. It passes a reference to the <span
class="keyword">PresentationPart</span> class instance to the
**CreatePresentationParts** procedure, which creates the other required
parts of the presentation file. The **CreatePresentation** procedure
cleans up by closing the <span
class="keyword">PresentationDocument</span> class instance that it
opened previously.

The **CreatePresentationParts** procedure creates instances of the <span
class="keyword">SlideMasterIdList</span>, <span
class="keyword">SlideIdList</span>, <span
class="keyword">SlideSize</span>, <span
class="keyword">NotesSize</span>, and <span
class="keyword">DefaultTextStyle</span> classes and appends them to the
presentation.

```csharp
    public static void CreatePresentation(string filepath)
            {
                // Create a presentation at a specified file path. The presentation document type is pptx, by default.
                PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
                PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                CreatePresentationParts(presentationPart);

                // Close the presentation handle.
                presentationDoc.Close();
            } 
    private static void CreatePresentationParts(PresentationPart presentationPart)
            {
                SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
                SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
                SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
                NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
                DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

               presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

             // Code to create other parts of the presentation file goes here.
            }
```

```vb
    Public Shared Sub CreatePresentation(ByVal filepath As String)

                ' Create a presentation at a specified file path. The presentation document type is pptx, by default.
                Dim presentationDoc As PresentationDocument = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation)
                Dim presentationPart As PresentationPart = presentationDoc.AddPresentationPart()
                presentationPart.Presentation = New Presentation()

                CreatePresentationParts(presentationPart)

                ' Close the presentation handle.
                presentationDoc.Close()
            End Sub
    Private Shared Sub CreatePresentationParts(ByVal presentationPart As PresentationPart)
                Dim slideMasterIdList1 As New SlideMasterIdList(New SlideMasterId() With { _
                 .Id = DirectCast(2147483648UI, UInt32Value), _
                 .RelationshipId = "rId1" _
                })
                Dim slideIdList1 As New SlideIdList(New SlideId() With { _
                 .Id = DirectCast(256UI, UInt32Value), _
                 .RelationshipId = "rId2" _
                })
                Dim slideSize1 As New SlideSize() With { _
                 .Cx = 9144000, _
                 .Cy = 6858000, _
                 .Type = SlideSizeValues.Screen4x3 _
                }
                Dim notesSize1 As New NotesSize() With { _
                 .Cx = 6858000, _
                 .Cy = 9144000 _
                }
                Dim defaultTextStyle1 As New DefaultTextStyle()

                presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1)

             ' Code to create other parts of the presentation file goes here.
            End Sub
```

---------------------------------------------------------------------------------

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

#### Concepts

[About the Open XML SDK 2.5 for Office](about-the-open-xml-sdk-2-5.md)  

[How to: Create a presentation document by providing a file name (Open XML SDK)](how-to-create-a-presentation-document-by-providing-a-file-name.md)  

[How to: Insert a new slide into a presentation (Open XML SDK)](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation (Open XML SDK)](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Retrieve the number of slides in a presentation document (Open XML SDK)](how-to-retrieve-the-number-of-slides-in-a-presentation-document.md)  

[How to: Apply a theme to a presentation (Open XML SDK)](how-to-apply-a-theme-to-a-presentation.md)  
