---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: ee6c905b-26c5-4aed-a414-9aa826364a23
title: Working with presentation slides
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 09/20/2024
ms.localizationpriority: high
---
# Working with presentation slides

This topic discusses the Open XML SDK for Office <xref:DocumentFormat.OpenXml.Presentation.Slide> class and how it relates to the Open
XML File Format PresentationML schema. For more information about the
overall structure of the parts and elements that make up a
PresentationML document, see [Structure of a PresentationML document](structure-of-a-presentationml-document.md).


---------------------------------------------------------------------------------
## Presentation Slides in PresentationML
The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML `<sld/>` element used
to represent a presentation slide in a PresentationML document as
follows:

This element specifies a slide within a slide list. The slide list is
used to specify an ordering of slides.

Example: Consider the following custom show with an ordering of slides.

```xml
<p:custShowLst>  
    <p:custShow name="Custom Show 1" id="0">  
        <p:sldLst>  
            <p:sld r:id="rId4"/>  
            <p:sld r:id="rId3"/>  
            <p:sld r:id="rId2"/>  
            <p:sld r:id="rId5"/>  
        </p:sldLst>  
    </p:custShow>  
</p:custShowLst>
```

In the above example the order specified to present the slides is slide
4, then 3, 2, and finally 5.


&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The `<sld/>` element is the root element of the PresentationML Slide
part. For more information about the overall structure of the parts and
elements that make up a PresentationML document, see [Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

The following table lists the child elements of the `<sld/>` element used
when working with presentation slides and the Open XML SDK classes
that correspond to them.


| **PresentationML Element** |                                                               **Open XML SDK Class**                                                                |
|----------------------------|--------------------------------------------------------|
|       `<clrMapOvr />`        |              <xref:DocumentFormat.OpenXml.Presentation.ColorMapOverride>              |
|          `<cSld />`          |               <xref:DocumentFormat.OpenXml.Presentation.CommonSlideData>               |
|         `<extLst />`         | <xref:DocumentFormat.OpenXml.Presentation.ExtensionListWithModification> |
|         `<timing />`         |                        <xref:DocumentFormat.OpenXml.Presentation.Timing>                       |
|       `<transition />`       |                    <xref:DocumentFormat.OpenXml.Presentation.Transition>                    |

--------------------------------------------------------------------------------
## Open XML SDK Slide Class
The Open XML SDK `Slide` class represents the `<sld/>` element
defined in the Open XML File Format schema for PresentationML documents.
Use the `Slide` object to manipulate
individual `<sld/>` elements in a PresentationML document.

Classes commonly associated with the `Slide`
class are shown in the following sections.

### ColorMapOverride Class

The `ColorMapOverride` class corresponds to
the `<clrMapOvr/>` element. The following information from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification introduces the `<clrMapOvr/>` element:

This element provides a mechanism with which to override the color
schemes listed within the `<ClrMap/>` element. If the
`<masterClrMapping/>` child element is present, the color scheme defined
by the master is used. If the `<overrideClrMapping/>` child element is
present, it defines a new color scheme specific to the parent notes
slide, presentation slide, or slide layout.

&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### CommonSlideData Class

The `CommonSlideData` class corresponds to
the `<cSld/>` element. The following information from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification introduces the `<cSld/>` element:

This element specifies a container for the type of slide information
that is relevant to all of the slide types. All slides share a common
set of properties that is independent of the slide type; the description
of these properties for any particular slide is stored within the
slide's `<cSld/>` container. Slide data specific to the slide type
indicated by the parent element is stored elsewhere.

The actual data in `<cSld/>` describe only the particular parent slide;
it is only the type of information stored that is common across all
slides.

&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### ExtensionListWithModification Class

The `ExtensionListWithModification` class
corresponds to the `<extLst/>` element. The following information from the
[!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification introduces the `<extLst/>` element:

This element specifies the extension list with modification ability
within which all future extensions of element type `<ext/>` are defined.
The extension list along with corresponding future extensions is used to
extend the storage capabilities of the PresentationML framework. This
allows for various new kinds of data to be stored natively within the
framework.

[Note: Using this `extLst` element allows the generating application to
store whether this extension property has been modified. end note]

&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### Timing Class

The `Timing` class corresponds to the
`<timing/>` element. The following information from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification introduces the `<timing/>` element:

This element specifies the timing information for handling all
animations and timed events within the corresponding slide. This
information is tracked via time nodes within the `<timing/>` element.
More information on the specifics of these time nodes and how they are
to be defined can be found within the Animation section of the
PresentationML framework.

&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

### Transition Class

The `Transition` class corresponds to the
`<transition/>` element. The following information from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification introduces the `<transition/>` element:

This element specifies the kind of slide transition that should be used
to transition to the current slide from the previous slide. That is, the
transition information is stored on the slide that appears after the
transition is complete.

&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]


--------------------------------------------------------------------------------
## Working with the Slide Class

As shown in the Open XML SDK code example that follows, every instance
of the `Slide` class is associated with an
instance of the <xref:DocumentFormat.OpenXml.Packaging.SlidePart> class, which represents a slide
part, one of the required parts of a PresentationML presentation file
package. Each instance of the `Slide` class
must also be associated with instances of the <xref:DocumentFormat.OpenXml.Presentation.SlideLayout> and <xref:DocumentFormat.OpenXml.Presentation.SlideMaster> classes, which are in turn
associated with similarly named required presentation parts, represented
by the <xref:DocumentFormat.OpenXml.Packaging.SlideLayoutPart> and <xref:DocumentFormat.OpenXml.Packaging.SlideMasterPart> classes.

The `Slide` class, which represents the
`<sld/>` element, is therefore also associated with a series of other
classes that represent the child elements of the `<sld/>` element. Among
these classes, as shown in the following code example, are the `CommonSlideData` class, the `ColorMapOverride` class, the <xref:DocumentFormat.OpenXml.Presentation.ShapeTree> class, and the <xref:DocumentFormat.OpenXml.Presentation.Shape> class.


--------------------------------------------------------------------------------
## Open XML SDK Code Example
The following method from the article [How to: Create a presentation document by providing a file name](how-to-create-a-presentation-document-by-providing-a-file-name.md) adds a new slide
part to an existing presentation and creates an instance of the Open XML
SDK `Slide` class in the new slide part.
The `Slide` class constructor creates
instances of the `CommonSlideData` and `ColorMapOverride` classes. The `CommonSlideData` class constructor creates an
instance of the `ShapeTree` class, whose
constructor, in turn, creates additional class instances: an instance of
the <xref:DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties> class, the <xref:DocumentFormat.OpenXml.Presentation.GroupShapeProperties> class, and the `Shape` class.

All of these class instances and instances of the classes that represent
the child elements of the `<sld/>` element are required to create the
minimum number of XML elements necessary to represent a new slide.

The namespace represented by the letter *P* in the code is the <xref:DocumentFormat.OpenXml.Presentation>
namespace.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/presentation/create_by_providing_a_file_name/cs/Program.cs#snippet102)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/presentation/create_by_providing_a_file_name/vb/Program.vb#snippet102)]
***


To add another shape to the shape tree and, hence, to the slide,
instantiate a second `Shape` object by
passing an additional parameter that contains the following code to the
`ShapeTree` constructor.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/create_by_providing_a_file_name/cs/Program.cs#snippet103)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/create_by_providing_a_file_name/vb/Program.vb#snippet103)]
***


---------------------------------------------------------------------------------
## Generated PresentationML
When the Open XML SDK code in the method is run, the following XML code
is written to the PresentationML document file referenced in the code.

```xml
    <?xml version="1.0" encoding="utf-8" ?> 
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr>
            <p:cNvPr id="1" name="" /> 
            <p:cNvGrpSpPr /> 
            <p:nvPr /> 
          </p:nvGrpSpPr>
            <p:grpSpPr>
              <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /> 
            </p:grpSpPr>
            <p:sp>
              <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1" /> 
              <p:cNvSpPr>
                <a:spLocks noGrp="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /> 
              </p:cNvSpPr>
                <p:nvPr>
                  <p:ph /> 
                </p:nvPr>
                </p:nvSpPr>
                <p:spPr /> 
                <p:txBody>
                 <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /> 
             <a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /> 
                 <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                   <a:endParaRPr lang="en-US" /> 
          </a:p>
                </p:txBody>
             </p:sp>
           </p:spTree>
        </p:cSld>
        <p:clrMapOvr>
          <a:masterClrMapping xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /> 
        </p:clrMapOvr>
    </p:sld>
```

--------------------------------------------------------------------------------
## See also


[About the Open XML SDK for Office](../about-the-open-xml-sdk.md)  

[How to: Insert a new slide into a presentation](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Retrieve the number of slides in a presentation document](how-to-retrieve-the-number-of-slides-in-a-presentation-document.md)  

[How to: Apply a theme to a presentation](how-to-apply-a-theme-to-a-presentation.md)  

[How to: Create a presentation document by providing a file name](how-to-create-a-presentation-document-by-providing-a-file-name.md)  
