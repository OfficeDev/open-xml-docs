---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: fb4b293c-9a23-44b7-8af6-afe5fac6611a
title: Working with handout master slides (Open XML SDK)
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Working with handout master slides (Open XML SDK)

This topic discusses the Open XML SDK for Office [HandoutMaster](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.handoutmaster.aspx) class and how it relates to the Open XML File Format PresentationML schema. For more information about the overall structure of the parts and elements that make up a PresentationML document, see [Structure of a PresentationML document (Open XML SDK)](structure-of-a-presentationml-document.md).

## Handout Master Slides in PresentationML
The [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML \<handoutMaster\>
element used to represent a handout master slide in a PresentationML document as follows:

This element specifies an instance of a handout master slide. Within a
handout master slide are contained all elements that describe the
objects and their corresponding formatting for within a handout slide.
Within a handout master slide the cSld element specifies the common
slide elements such as shapes and their attached text bodies. There are
other properties within a handout master slide but cSld encompasses the
majority of the intended purpose for a handoutMaster slide.

© ISO/IEC29500: 2008.

The following table lists the child elements of the \<handoutMaster\>
element used when working with handout master slides and the Open XML
SDK classes that correspond to them.


| **PresentationML Element** |                                                               **Open XML SDK Class**                                                                |
|----------------------------|--------------------------------------------------------------------------------------------------------------------------------------------|
|         \<clrMap\>         |                      [ColorMap](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.colormap.aspx)                      |
|          \<cSld\>          |               [CommonSlideData](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.commonslidedata.aspx)               |
|         \<extLst\>         | [ExtensionListWithModification](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.extensionlistwithmodification.aspx) |
|           \<hf\>           |                  [HeaderFooter](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.headerfooter.aspx)                  |

## Open XML SDK HandoutMaster Class

The Open XML SDK **HandoutMaster** class represents the \<handoutMaster\> element defined in the Open XML File Format schema for PresentationML documents. Use the **HandoutMaster** class to manipulate individual \<handoutMaster\> elements in a PresentationML document.

Classes commonly associated with the **HandoutMaster** class are shown in the following sections.

### ColorMap Class

The **ColorMap** class corresponds to the \<clrMap\> element. The following information from the [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification introduces the \<clrMap\> element:

This element specifies the mapping layer that transforms one color
scheme definition to another. Each attribute represents a color name
that can be referenced in this master, and the value is the
corresponding color in the theme.

**Example**: Consider the following mapping of colors that applies to a
slide master:

```xml
<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1"  
accent2="accent2" accent3="accent3" accent4="accent4"
accent5="accent5"  
accent6="accent6" hlink="hlink" folHlink="folHlink"/>
```

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

> [!NOTE]
> Using this extLst element allows the generating application to
store whether this extension property has been modified.

### HeaderFooter Class

The **HeaderFooter** class corresponds to the
\<hf\> element. The following information from the [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the \<hf\> element:

This element specifies the header and footer information for a slide.
Headers and footers consist of placeholders for text that should be
consistent across all slides and slide types, such as a date and time,
slide numbering, and custom header and footer text.


## Working with the HandoutMaster Class

As shown in the Open XML SDK code sample that follows, every instance of
the **HandoutMaster** class is associated with
an instance of the [HandoutMasterPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.handoutmasterpart.aspx) class, which represents a
handout master part, one of the parts of a PresentationML presentation
file package, and a part that is required for a presentation file that
contains handouts.

The **HandoutMaster** class, which represents
the \<handoutMaster\> element, is therefore also associated with a
series of other classes that represent the child elements of the
\<handoutMaster\> element. Among these classes, as shown in the
following code sample, are the **CommonSlideData** class, the **ColorMap** class, the [ShapeTree](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.shapetree.aspx) class, and the [Shape](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.shape.aspx) class.

## Open XML SDK Code Example

The following method adds a new handout master part to an existing
presentation and creates an instance of an Open XML SDK**HandoutMaster** class in the new handout master
part. The **HandoutMaster** class constructor
creates instances of the **CommonSlideData**
class and the **ColorMap** class. The **CommonSlideData** class constructor creates an
instance of the [ShapeTree](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.shapetree.aspx) class, whose constructor, in
turn, creates additional class instances: an instance of the [NonVisualGroupShapeProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.nonvisualgroupshapeproperties.aspx) class, an
instance of the [GroupShapeProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.groupshapeproperties.aspx) class, and an instance
of the [Shape](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.shape.aspx) class.

The namespace represented by the letter *P* in the code is the [DocumentFormat.OpenXml.Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.aspx)
namespace.

### [C#](#tab/cs)
[!code-csharp[](../samples/presentation/working_with_handout_master_slides/cs/Program.cs)]
### [Visual Basic](#tab/vb)
[!code-vb[](../samples/presentation/working_with_handout_master_slides/vb/Program.vb)]

## Generated PresentationML

When the Open XML SDK code is run, the following XML is written to
the PresentationML document referenced in the code.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <p:handoutMaster xmlns:p="https://schemas.openxmlformats.org/presentationml/2006/main">
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
                       name="Title 1" />
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
                <a:endParaRPr lang="en-US" />
              </a:p>
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
    </p:handoutMaster>
```

## See also

[About the Open XML SDK for Office](about-the-open-xml-sdk.md)
[How to: Create a presentation document by providing a file name (Open XML SDK)](how-to-create-a-presentation-document-by-providing-a-file-name.md)
[How to: Insert a new slide into a presentation (Open XML SDK)](how-to-insert-a-new-slide-into-a-presentation.md)
[How to: Delete a slide from a presentation (Open XML SDK)](how-to-delete-a-slide-from-a-presentation.md)
[How to: Retrieve the number of slides in a presentation document (Open XML SDK)](how-to-retrieve-the-number-of-slides-in-a-presentation-document.md)
[How to: Apply a theme to a presentation (Open XML SDK)](how-to-apply-a-theme-to-a-presentation.md)
