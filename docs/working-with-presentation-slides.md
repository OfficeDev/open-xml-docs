---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: ee6c905b-26c5-4aed-a414-9aa826364a23
title: Working with presentation slides (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Working with presentation slides (Open XML SDK)

This topic discusses the Open XML SDK 2.5 for Office <span sdata="cer"
target="T:DocumentFormat.OpenXml.Presentation.Slide"><span
class="nolink">Slide</span></span> class and how it relates to the Open
XML File Format PresentationML schema. For more information about the
overall structure of the parts and elements that make up a
PresentationML document, see <span sdata="link">[Structure of a
PresentationML document (Open XML SDK)](structure-of-a-presentationml-document.md)</span>.


---------------------------------------------------------------------------------

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<sld\> element used
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


© ISO/IEC29500: 2008.

The \<sld\> element is the root element of the PresentationML Slide
part. For more information about the overall structure of the parts and
elements that make up a PresentationML document, see [Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

The following table lists the child elements of the \<sld\> element used
when working with presentation slides and the Open XML SDK 2.5 classes
that correspond to them.

**PresentationML Element**|**Open XML SDK 2.5 Class**
---|---
\<clrMapOvr\>|ColorMapOverride
\<cSld\>|CommonSlideData
\<extLst\>|ExtensionListWithModification
\<timing\>|Timing
\<transition\>|Transition


--------------------------------------------------------------------------------

The Open XML SDK 2.5**Slide** class represents the \<sld\> element
defined in the Open XML File Format schema for PresentationML documents.
Use the **Slide** object to manipulate
individual \<sld\> elements in a PresentationML document.

Classes commonly associated with the **Slide**
class are shown in the following sections.

### ColorMapOverride Class

The **ColorMapOverride** class corresponds to
the \<clrMapOvr\> element. The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
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
the \<cSld\> element. The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
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
[ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
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

### Timing Class

The **Timing** class corresponds to the
\<timing\> element. The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<timing\> element:

This element specifies the timing information for handling all
animations and timed events within the corresponding slide. This
information is tracked via time nodes within the \<timing\> element.
More information on the specifics of these time nodes and how they are
to be defined can be found within the Animation section of the
PresentationML framework.

© ISO/IEC29500: 2008.

### Transition Class

The **Transition** class corresponds to the
\<transition\> element. The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<transition\> element:

This element specifies the kind of slide transition that should be used
to transition to the current slide from the previous slide. That is, the
transition information is stored on the slide that appears after the
transition is complete.

© ISO/IEC29500: 2008.


--------------------------------------------------------------------------------

As shown in the Open XML SDK code example that follows, every instance
of the **Slide** class is associated with an
instance of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.SlidePart"><span
class="nolink">SlidePart</span></span> class, which represents a slide
part, one of the required parts of a PresentationML presentation file
package. Each instance of the **Slide** class
must also be associated with instances of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Presentation.SlideLayout"><span
class="nolink">SlideLayout</span></span> and <span sdata="cer"
target="T:DocumentFormat.OpenXml.Presentation.SlideMaster"><span
class="nolink">SlideMaster</span></span> classes, which are in turn
associated with similarly named required presentation parts, represented
by the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.SlideLayoutPart"><span
class="nolink">SlideLayoutPart</span></span> and <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.SlideMasterPart"><span
class="nolink">SlideMasterPart</span></span> classes.

The **Slide** class, which represents the
\<sld\> element, is therefore also associated with a series of other
classes that represent the child elements of the \<sld\> element. Among
these classes, as shown in the following code example, are the <span
class="keyword">CommonSlideData</span> class, the <span
class="keyword">ColorMapOverride</span> class, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Presentation.ShapeTree"><span
class="nolink">ShapeTree</span></span> class, and the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Presentation.Shape"><span
class="nolink">Shape</span></span> class.


--------------------------------------------------------------------------------

The following method from the article <span sdata="link">[How to: Create a presentation document by providing a file name (Open XML SDK)](how-to-create-a-presentation-document-by-providing-a-file-name.md)</span> adds a new slide
part to an existing presentation and creates an instance of the Open XML
SDK 2.5**Slide** class in the new slide part.
The **Slide** class constructor creates
instances of the **CommonSlideData** and <span
class="keyword">ColorMapOverride</span> classes. The <span
class="keyword">CommonSlideData</span> class constructor creates an
instance of the **ShapeTree** class, whose
constructor, in turn, creates additional class instances: an instance of
the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties"><span
class="nolink">NonVisualGroupShapeProperties</span></span> class, the
<span sdata="cer"
target="T:DocumentFormat.OpenXml.Presentation.GroupShapeProperties"><span
class="nolink">GroupShapeProperties</span></span> class, and the <span
class="keyword">Shape</span> class.

All of these class instances and instances of the classes that represent
the child elements of the \<sld\> element are required to create the
minimum number of XML elements necessary to represent a new slide.

The namespace represented by the letter *P* in the code is the <span
sdata="cer" target="N:DocumentFormat.OpenXml.Presentation"><span
class="nolink">DocumentFormat.OpenXml.Presentation</span></span>
namespace.

```csharp
    private static SlidePart CreateSlidePart(PresentationPart presentationPart)        
            {
                SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
                    slidePart1.Slide = new Slide(
                            new CommonSlideData(
                                new ShapeTree(
                                    new P.NonVisualGroupShapeProperties(
                                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                        new P.NonVisualGroupShapeDrawingProperties(),
                                        new ApplicationNonVisualDrawingProperties()),
                                    new GroupShapeProperties(new TransformGroup()),
                                    new P.Shape(
                                        new P.NonVisualShapeProperties(
                                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                        new P.ShapeProperties(),
                                        new P.TextBody(
                                            new BodyProperties(),
                                            new ListStyle(),
                                            new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
                            new ColorMapOverride(new MasterColorMapping()));
                    return slidePart1;
             }
```

```vb
    Private Shared Function CreateSlidePart(ByVal presentationPart As PresentationPart) As SlidePart
                Dim slidePart1 As SlidePart = presentationPart.AddNewPart(Of SlidePart)("rId2")
                slidePart1.Slide = New Slide(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                 .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New GroupShapeProperties(New TransformGroup()), _
                   New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "Title 1" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), _
                   New ListStyle(), New Paragraph(New EndParagraphRunProperties() With { _
                  .Language = "en-US" _
                }))))), New ColorMapOverride(New MasterColorMapping()))
                Return slidePart1
            End Function
```

To add another shape to the shape tree and, hence, to the slide,
instantiate a second **Shape** object by
passing an additional parameter that contains the following code to the
**ShapeTree** constructor.

```csharp
    new P.Shape(
              new P.NonVisualShapeProperties(
                  new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                  new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                  new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                  new BodyProperties(),
                  new ListStyle(),
                  new Paragraph(new EndParagraphRunProperties() { Language = "en-US" })))
```

```vb
    New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "Title 1" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), _
                   New ListStyle(), New Paragraph(New EndParagraphRunProperties() With { _
                  .Language = "en-US" })))
```

---------------------------------------------------------------------------------

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

#### Concepts

[About the Open XML SDK 2.5 for Office](about-the-open-xml-sdk-2-5.md)  

[How to: Insert a new slide into a presentation (Open XML SDK)](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation (Open XML SDK)](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Retrieve the number of slides in a presentation document (Open XML SDK)](how-to-retrieve-the-number-of-slides-in-a-presentation-document.md)  

[How to: Apply a theme to a presentation (Open XML SDK)](how-to-apply-a-theme-to-a-presentation.md)  

[How to: Create a presentation document by providing a file name (Open XML SDK)](how-to-create-a-presentation-document-by-providing-a-file-name.md)  
