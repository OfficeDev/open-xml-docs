---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 7dfd78a3-e233-4abd-8c17-1e384780d3ec
title: Working with slide masters (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Working with slide masters (Open XML SDK)

This topic discusses the Open XML SDK 2.5 for Office **[SlideMaster](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.slidemaster.aspx)** class and how it relates to the
Open XML File Format PresentationML schema.


---------------------------------------------------------------------------------
## Slide Masters in PresentationML 

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML \<sldMaster\>
element used to represent slide layouts in a PresentationML document as
follows.

This element specifies an instance of a slide master slide. Within a
slide master slide are contained all elements that describe the objects
and their corresponding formatting for within a presentation slide.
Within a slide master slide are two main elements. The cSld element
specifies the common slide elements such as shapes and their attached
text bodies. Then the txStyles element specifies the formatting for the
text within each of these shapes. The other properties within a slide
master slide specify other properties for within a presentation slide
such as color information, headers and footers, as well as timing and
transition information for all corresponding presentation slides.

© ISO/IEC29500: 2008.

The \<sldMaster\> element is the root element of the PresentationML
Slide Master part. For more information about the overall structure of
the parts and elements that make up a PresentationML document, see
[Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

The following table lists the child elements of the \<sldMaster\>
element used when working with slide masters and the Open XML SDK 2.5
classes that correspond to them.

**PresentationML Element**|**Open XML SDK 2.5 Class**
---|---
\<clrMap\>|[ColorMap](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.colormap.aspx)
\<cSld\>|[CommonSlideData](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.commonslidedata.aspx)
\<extLst\>|[ExtensionListWithModification](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.extensionlistwithmodification.aspx)
\<hf\>|[HeaderFooter](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.headerfooter.aspx)
\<sldLayoutIdLst\>|[SlideLayoutIdList](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.slidelayoutidlist.aspx)
\<timing\>|[Timing](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.timing.aspx)
\<transition\>|[Transition](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.transition.aspx)
\<txStyles\>|[TextStyles](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.textstyles.aspx)

The following table from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the attributes of the \<sldMaster\> element.

**Attributes**|**Description**
---|---
preserve (Preserve Slide Master)|Specifies whether the corresponding slide layout is deleted when all the slides that follow that layout are deleted. If this attribute is not specified then a value of **false** should be assumed by the generating application. This would mean that the slide would in fact be deleted if no slides within the presentation were related to it.<br/>The possible values for this attribute are defined by the W3C XML Schema **Boolean** data type.


© ISO/IEC29500: 2008.


--------------------------------------------------------------------------------
## Open XML SDK 2.5 SlideMaster Class 

The Open XML SDK 2.5**SlideMaster** class
represents the \<sldMaster\> element defined in the Open XML File Format
schema for PresentationML documents. Use the **SlideMaster** class to manipulate individual
\<sldMaster\> elements in a PresentationML document.

Classes that represent child elements of the \<sldMaster\> element and
that are therefore commonly associated with the **SlideMaster** class are shown in the following
list.

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

### HeaderFooter Class

The **HeaderFooter** class corresponds to the
\<hf\> element. The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<hf\> element:

This element specifies the header and footer information for a slide.
Headers and footers consist of placeholders for text that should be
consistent across all slides and slide types, such as a date and time,
slide numbering, and custom header and footer text.

© ISO/IEC29500: 2008.

### SlideLayoutIdList Class

The **SlideLayoutIdList** class corresponds to
the \<sldLayoutIdLst\> element. The following information from the
[ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<sldLayoutIdLst\> element:

This element specifies the existence of the slide layout identification
list. This list is contained within the slide master and is used to
determine which layouts are being used within the slide master file.
Each layout within the list of slide layouts has its own identification
number and relationship identifier that uniquely identifies it within
both the presentation document and the particular master slide within
which it is used.

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

### TextStyles Class

The **TextStyles** class corresponds to the
\<txStyles\> element. The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<txStyles\> element:

This element specifies the text styles within a slide master. Within
this element is the styling information for title text, the body text
and other slide text as well. This element is only for use within the
Slide Master and thus sets the text styles for the corresponding
presentation slides.

Example: Consider the case where we would like to specify the title
text for a master slide.

```xml
<p:txStyles>  
    <p:titleStyle>
        <a:lvl1pPr algn="ctr" rtl="0" latinLnBrk="0">  
            <a:spcBef>  
                <a:spcPct val="0"/>  
            </a:spcBef>  
            <a:buNone/>  
            <a:defRPr sz="4400" kern="1200">  
                <a:solidFill>vv  
                    <a:schemeClr val="tx1"/>  
                </a:solidFill\>  
                <a:latin typeface="+mj-lt"/>  
                <a:ea typeface="+mj-ea"/>  
                <a:cs typeface="+mj-cs"/>  
            </a:defRPr>  
        </a:lvl1pPr>  
    </p:titleStyle>  
</p:txStyles>
```

In the above example the title text is set according to the above
formatting for all related slides within the presentation. end example]

© ISO/IEC29500: 2008.


--------------------------------------------------------------------------------
## Working with the SlideMaster Class 

As shown in the Open XML SDK code sample that follows, every instance of
the **SlideMaster** class is associated with an
instance of the **[SlideMasterPart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.slidemasterpart.aspx)** class, which represents a
slide master part, one of the required parts of a PresentationML
presentation file package. Each **SlideMaster**
class instance must also be associated with instances of the **[SlideLayout](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.slidelayout.aspx)** and <**[Slide](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.slide.aspx)** classes, which are in turn associated
with similarly named required presentation parts, represented by the
**[SlideLayoutPart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.slidelayoutpart.aspx)** and **[SlidePart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.slidepart.aspx)** classes.

The **SlideMaster** class, which represents the
\<sldMaster\> element, is therefore also associated with a series of
other classes that represent the child elements of the \<sldMaster\>
element. Among these classes, as shown in the following code sample, are
the **CommonSlideData** class, the **ColorMap** class, the **[ShapeTree](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.shapetree.aspx)** class, and the **[Shape](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.shape.aspx)** class.


--------------------------------------------------------------------------------
## Open XML SDK Code Example 

The following method from the article <span sdata="link">[How to: Create a presentation document by providing a file name (Open XML SDK)](how-to-create-a-presentation-document-by-providing-a-file-name.md) adds a new slide
master part to an existing presentation and creates an instance of an
Open XML SDK 2.5**SlideMaster** class in the
new slide master part. The **SlideMaster**
class constructor creates instances of the **CommonSlideData** class and the **ColorMap**, **SlideLayoutIdList**, and **TextStyles** classes. The **CommonSlideData** class constructor creates an
instance of the **[ShapeTree](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.shapetree.aspx)** class, whose constructor in turn
creates additional class instances: an instance of the **[NonVisualGroupShapeProperties](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.nonvisualgroupshapeproperties.aspx)** class, an
instance of the **[GroupShapeProperties](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.groupshapeproperties.aspx)** class, and an instance
of the **[Shape](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.shape.aspx)** class, among others.

The namespace represented by the letter *P* in the code is the **[DocumentFormat.OpenXml.Presentation](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.presentation.aspx)**
namespace.

```csharp
    private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
       {
           SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
           SlideMaster slideMaster = new SlideMaster(
           new CommonSlideData(new ShapeTree(
             new P.NonVisualGroupShapeProperties(
             new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
             new P.NonVisualGroupShapeDrawingProperties(),
             new ApplicationNonVisualDrawingProperties()),
             new GroupShapeProperties(new TransformGroup()),
             new P.Shape(
             new P.NonVisualShapeProperties(
               new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
               new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
               new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
             new P.ShapeProperties(),
             new P.TextBody(
               new BodyProperties(),
               new ListStyle(),
               new Paragraph())))),
           new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
           new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
           new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
           slideMasterPart1.SlideMaster = slideMaster;
           return slideMasterPart1;
        }
```

```vb
    Private Shared Function CreateSlideMasterPart(ByVal slideLayoutPart1 As SlideLayoutPart) As SlideMasterPart
                Dim slideMasterPart1 As SlideMasterPart = slideLayoutPart1.AddNewPart(Of SlideMasterPart)("rId1")
                Dim slideMaster As New SlideMaster(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), _
                    New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "Title Placeholder 1" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With { _
                  .Type = PlaceholderValues.Title _
                })), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph())))), New P.ColorMap() With { _
                  .Background1 = D.ColorSchemeIndexValues.Light1, _
                  .Text1 = D.ColorSchemeIndexValues.Dark1, _
                  .Background2 = D.ColorSchemeIndexValues.Light2, _
                  .Text2 = D.ColorSchemeIndexValues.Dark2, _
                  .Accent1 = D.ColorSchemeIndexValues.Accent1, _
                  .Accent2 = D.ColorSchemeIndexValues.Accent2, _
                  .Accent3 = D.ColorSchemeIndexValues.Accent3, _
                  .Accent4 = D.ColorSchemeIndexValues.Accent4, _
                  .Accent5 = D.ColorSchemeIndexValues.Accent5, _
                  .Accent6 = D.ColorSchemeIndexValues.Accent6, _
                  .Hyperlink = D.ColorSchemeIndexValues.Hyperlink, _
                  .FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink _
                }, New SlideLayoutIdList(New SlideLayoutId() With { _
                  .Id = CType(2147483649UI, UInt32Value), _
                  .RelationshipId = "rId1" _
                }), New TextStyles(New TitleStyle(), New BodyStyle(), New OtherStyle()))
                slideMasterPart1.SlideMaster = slideMaster

                Return slideMasterPart1
            End Function
```

---------------------------------------------------------------------------------
## Generated PresentationML 

When the Open XML SDK code is run, the following XML is written to the
PresentationML document referenced in the code.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr>
            <p:cNvPr id="1"
                     name="" />
            <p:cNvGrpSpPr />
            <p:nvPr />
          </p:nvGrpSpPr>
          <p:grpSpPr>
            <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
          </p:grpSpPr>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2"
                       name="Title Placeholder 1" />
              <p:cNvSpPr>
                <a:spLocks noGrp="1"
                           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="title" />
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr />
            <p:txBody>
              <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
              <a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
              <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMap bg1="lt1"
                tx1="dk1"
                bg2="lt2"
                tx2="dk2"
                accent1="accent1"
                accent2="accent2"
                accent3="accent3"
                accent4="accent4"
                accent5="accent5"
                accent6="accent6"
                hlink="hlink"
                folHlink="folHlink" />
      <p:sldLayoutIdLst>
        <p:sldLayoutId id="2147483649"
                       r:id="rId1"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
      </p:sldLayoutIdLst>
      <p:txStyles>
        <p:titleStyle />
        <p:bodyStyle />
        <p:otherStyle />
      </p:txStyles>
    </p:sldMaster>
```

--------------------------------------------------------------------------------
## See also 

#### Concepts

[About the Open XML SDK 2.5 for Office](about-the-open-xml-sdk.md)  

[How to: Create a Presentation by Providing a File Name](how-to-create-a-presentation-document-by-providing-a-file-name.md)  

[How to: Insert a new slide into a presentation (Open XML SDK)](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation (Open XML SDK)](how-to-delete-a-slide-from-a-presentation.md)  
