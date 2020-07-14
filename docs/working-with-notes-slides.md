---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 56d28bc5-c9ea-4c0e-b2f5-20be9c16d290
title: Working with notes slides (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
localization_priority: Normal
---
# Working with notes slides (Open XML SDK)

This topic discusses the Open XML SDK 2.5 for Office [NotesSlide](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.notesslide.aspx) class and how it relates to the
Open XML File Format PresentationML schema.


--------------------------------------------------------------------------------
## Notes Slides in PresentationML
The [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<notes\> element
used to represent notes slides in a PresentationML document as follows:

This element specifies the existence of a notes slide along with its
corresponding data. Contained within a notes slide are all the common
slide elements along with additional properties that are specific to the
notes element.

**Example**: Consider the following PresentationML notes slide:

```xml
<p:notes>  
    <p:cSld>  
        …  
    </p:cSld>  
    …  
</p:notes>
```

In the above example a notes element specifies the existence of a notes
slide with all of its parts. Notice the cSld element, which specifies
the common elements that can appear on any slide type and then any
elements specify additional non-common properties for this notes slide.

© ISO/IEC29500: 2008.

The \<notes\> element is the root element of the PresentationML Notes
Slide part. For more information about the overall structure of the
parts and elements that make up a PresentationML document, see
[Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

The following table lists the child elements of the \<notes\> element
used when working with notes slides and the Open XML SDK 2.5 classes
that correspond to them.


| **PresentationML Element** |                                                               **Open XML SDK 2.5 Class**                                                                |
|----------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------|
|       \<clrMapOvr\>        |              [ColorMapOverride](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.colormapoverride.aspx)              |
|          \<cSld\>          |               [CommonSlideData](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.commonslidedata.aspx)               |
|         \<extLst\>         | [ExtensionListWithModification](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.extensionlistwithmodification.aspx) |

The following table from the [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the attributes of the \<notes\> element.


|                    **Attributes**                     |                                                                                     **Description**                                                                                      |
|-------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| showMasterPhAnim (Show Master Placeholder Animations) | Specifies whether or not to display animations on placeholders from the master slide.<br/>The possible values for this attribute are defined by the W3C XML Schema **boolean** datatype. |
|           showMasterSp (Show Master Shapes)           |       Specifies if shapes on the master slide should be shown on slides or not.<br/>The possible values for this attribute are defined by the W3C XML Schema **boolean** datatype.       |

© ISO/IEC29500: 2008.


---------------------------------------------------------------------------------
## Open XML SDK 2.5 NotesSlide Class
The OXML SDK **NotesSlide** class represents
the \<notes\> element defined in the Open XML File Format schema for
PresentationML documents. Use the **NotesSlide** class to manipulate individual
\<notes\> elements in a PresentationML document.

Classes that represent child elements of the \<notes\> element and that
are therefore commonly associated with the **NotesSlide** class are shown in the following list.

### ColorMapOverride Class

The **ColorMapOverride** class corresponds to
the \<clrMapOvr\> element. The following information from the [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<clrMapOvr\> element:

This element provides a mechanism with which to override the color
schemes listed within the \<ClrMap\> element. If the
\<masterClrMapping\> child element is present, the color scheme defined
by the master is used. If the \<overrideClrMapping\> child element is
present, it defines a new color scheme specific to the parent notes
slide, presentation slide, or slide layout.

© ISO/IEC29500: 2008.

### CommonSlideData Class

The **CommonSlideData** class corresponds to
the \<cSld\> element. The following information from the [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<cSld\> element:

This element specifies a container for the type of slide information
that is relevant to all of the slide types. All slides share a common
set of properties that is independent of the slide type; the description
of these properties for any particular slide is stored within the
slide's \<cSld\> container. Slide data specific to the slide type
indicated by the parent element is stored elsewhere.

The actual data in \<cSld\> describe only the particular parent slide;
it is only the type of information stored that is common across all
slides.

© ISO/IEC29500: 2008.

### ExtensionListWithModification Class

The **ExtensionListWithModification** class
corresponds to the \<extLst\>element. The following information from the
[ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<extLst\> element:

This element specifies the extension list with modification ability
within which all future extensions of element type \<ext\> are defined.
The extension list along with corresponding future extensions is used to
extend the storage capabilities of the PresentationML framework. This
allows for various new kinds of data to be stored natively within the
framework.

[Note: Using this extLst element allows the generating application to
store whether this extension property has been modified. end note]

© ISO/IEC29500: 2008.


---------------------------------------------------------------------------------
## Working with the NotesSlide Class
As shown in the Open XML SDK code sample that follows, every instance of
the **NotesSlide** class is associated with an
instance of the [NotesSlidePart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.notesslidepart.aspx) class, which represents a
notes slide part, one of the parts of a PresentationML presentation file
package, and a part that is required for each notes slide in a
presentation file. Each **NotesSlide** class
instance may also be associated with an instance of the [NotesMaster](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.notesmaster.aspx) class, which in turn is
associated with a similarly named presentation part, represented by the
[NotesMasterPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.notesmasterpart.aspx) class.

The **NotesSlide** class, which represents the
\<notes\> element, is therefore also associated with a series of other
classes that represent the child elements of the \<notes\> element.
Among these classes, as shown in the following code sample, are the
**CommonSlideData** class and the **ColorMapOverride** class. The [ShapeTree](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.shapetree.aspx) class and the [Shape](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.shape.aspx) classes are in turn associated with
the **CommonSlideData** class.


--------------------------------------------------------------------------------
## Open XML SDK Code Example
The following method adds a new notes slide part to an existing
presentation and creates an instance of an Open XML SDK 2.5**NotesSlide** class in the new notes slide part. The
**NotesSlide** class constructor creates
instances of the **CommonSlideData** class and
the **ColorMap** class. The **CommonSlideData** class constructor creates an
instance of the [ShapeTree](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.shapetree.aspx) class, whose constructor in turn
creates additional class instances: an instance of the [NonVisualGroupShapeProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.nonvisualgroupshapeproperties.aspx) class, an
instance of the [GroupShapeProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.groupshapeproperties.aspx) class, and an instance
of the [Shape](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.shape.aspx) class.

The namespace represented by the letter *P* in the code is the [DocumentFormat.OpenXml.Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.aspx)
namespace.

```csharp
    private static NotesSlidePart CreateNotesSlidePart(SlidePart slidePart1)
        {
            NotesSlidePart notesSlidePart1 = slidePart1.AddNewPart<NotesSlidePart>("rId6");
            NotesSlide notesSlide = new NotesSlide(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(new EndParagraphRunProperties()))))),
            new ColorMapOverride(new MasterColorMapping()));
            notesSlidePart1.NotesSlide = notesSlide;
            return notesSlidePart1;
        }
```

```vb
    Private Shared Function CreateNotesSlidePart(ByVal slidePart1 As SlidePart) As NotesSlidePart
            Dim notesSlidePart1 As NotesSlidePart = slidePart1.AddNewPart(Of NotesSlidePart)("rId6")
            Dim notesSlide As New NotesSlide(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
             .Id = DirectCast(1UI, UInt32Value), _
             .Name = "" _
            }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New  _
                GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
             .Id = DirectCast(2UI, UInt32Value), _
             .Name = "" _
            }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
             .NoGrouping = True _
            }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New  _
                P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph(New EndParagraphRunProperties()))))),
            New ColorMapOverride(New MasterColorMapping()))
            notesSlidePart1.NotesSlide = notesSlide
            Return notesSlidePart1
        End Function
```

---------------------------------------------------------------------------------
## Generated PresentationML
When the Open XML SDK 2.0 code is run, the following XML is written to
the PresentationML document referenced in the code.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <p:notes xmlns:p="https://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr>
            <p:cNvPr id="1"
                     name="" />
            <p:cNvGrpSpPr />
            <p:nvPr />
          </p:nvGrpSpPr>
          <p:grpSpPr>
            <a:xfrm xmlns:a="https://schemas.openxmlformats.org/drawingml/2006/main" />
          </p:grpSpPr>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2"
                       name="" />
              <p:cNvSpPr>
                <a:spLocks noGrp="1"
                           xmlns:a="https://schemas.openxmlformats.org/drawingml/2006/main" />
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph />
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr />
            <p:txBody>
              <a:bodyPr xmlns:a="https://schemas.openxmlformats.org/drawingml/2006/main" />
              <a:lstStyle xmlns:a="https://schemas.openxmlformats.org/drawingml/2006/main" />
              <a:p xmlns:a="https://schemas.openxmlformats.org/drawingml/2006/main">
                <a:endParaRPr />
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping xmlns:a="https://schemas.openxmlformats.org/drawingml/2006/main" />
      </p:clrMapOvr>
    </p:notes>
```

--------------------------------------------------------------------------------
## See also


[About the Open XML SDK 2.5 for Office](about-the-open-xml-sdk.md)  

[How to: Create a Presentation by Providing a File Name](how-to-create-a-presentation-document-by-providing-a-file-name.md)  

[How to: Insert a new slide into a presentation (Open XML SDK)](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation (Open XML SDK)](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Apply a theme to a presentation (Open XML SDK)](how-to-apply-a-theme-to-a-presentation.md)  
