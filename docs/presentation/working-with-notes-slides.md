---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 56d28bc5-c9ea-4c0e-b2f5-20be9c16d290
title: Working with notes slides
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/20/2024
ms.localizationpriority: medium
---
# Working with notes slides

This topic discusses the Open XML SDK for Office <xref:DocumentFormat.OpenXml.Presentation.NotesSlide> class and how it relates to the
Open XML File Format PresentationML schema.


--------------------------------------------------------------------------------

## Notes Slides in PresentationML

The [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the Open XML PresentationML `<notes/>` element
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

&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The `<notes/>` element is the root element of the PresentationML Notes
Slide part. For more information about the overall structure of the
parts and elements that make up a PresentationML document, see
[Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

The following table lists the child elements of the `<notes/>` element
used when working with notes slides and the Open XML SDK classes
that correspond to them.


| **PresentationML Element** |    **Open XML SDK Class**    |
|----------------------------|-------------------------------------|
|       `<clrMapOvr/>`        |              <xref:DocumentFormat.OpenXml.Presentation.ColorMapOverride>              |
|          `<cSld/>`          |               <xref:DocumentFormat.OpenXml.Presentation.CommonSlideData>               |
|         `<extLst/>`         | <xref:DocumentFormat.OpenXml.Presentation.ExtensionListWithModification> |

The following table from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]
specification describes the attributes of the `<notes/>` element.


|                    **Attributes**                     | **Description**    |
|-------------------------------------------------------|---------------------|
| showMasterPhAnim (Show Master Placeholder Animations) | Specifies whether or not to display animations on placeholders from the master slide.<br/><br/>The possible values for this attribute are defined by the W3C XML Schema `boolean` datatype. |
|           showMasterSp (Show Master Shapes)           |       Specifies if shapes on the master slide should be shown on slides or not.<br/><br/>The possible values for this attribute are defined by the W3C XML Schema `boolean` datatype.       |

&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]


---------------------------------------------------------------------------------

## Open XML SDK NotesSlide Class

The OXML SDK `NotesSlide` class represents
the `<notes/>` element defined in the Open XML File Format schema for
PresentationML documents. Use the `NotesSlide` class to manipulate individual
`<notes/>` elements in a PresentationML document.

Classes that represent child elements of the `<notes/>` element and that
are therefore commonly associated with the `NotesSlide` class are shown in the following list.

- ### ColorMapOverride Class

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

- ### CommonSlideData Class

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

- ### ExtensionListWithModification Class

  The `ExtensionListWithModification` class
  corresponds to the `<extLst/>`element. The following information from the
  [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification introduces the `<extLst/>` element:
  
  This element specifies the extension list with modification ability
  within which all future extensions of element type `<ext/>` are defined.
  The extension list along with corresponding future extensions is used to
  extend the storage capabilities of the PresentationML framework. This
  allows for various new kinds of data to be stored natively within the
  framework.
  
  [Note: Using this extLst element allows the generating application to
  store whether this extension property has been modified. end note]
  
  &copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]


---------------------------------------------------------------------------------
## Working with the NotesSlide Class

As shown in the Open XML SDK code sample that follows, every instance of
the `NotesSlide` class is associated with an
instance of the <xref:DocumentFormat.OpenXml.Packaging.NotesSlidePart> class, which represents a
notes slide part, one of the parts of a PresentationML presentation file
package, and a part that is required for each notes slide in a
presentation file. Each `NotesSlide` class
instance may also be associated with an instance of the <xref:DocumentFormat.OpenXml.Presentation.NotesMaster> class, which in turn is
associated with a similarly named presentation part, represented by the
<xref:DocumentFormat.OpenXml.Packaging.NotesMasterPart> class.

The `NotesSlide` class, which represents the
`<notes/>` element, is therefore also associated with a series of other
classes that represent the child elements of the `<notes/>` element.
Among these classes, as shown in the following code sample, are the
`CommonSlideData` class and the `ColorMapOverride` class. The <xref:DocumentFormat.OpenXml.Presentation.ShapeTree> class and the <xref:DocumentFormat.OpenXml.Presentation.Shape> classes are in turn associated with
the `CommonSlideData` class.


--------------------------------------------------------------------------------
## Open XML SDK Code Example

> In the following code snippets `P` represents the <a href="xref:DocumentFormat.OpenXml.Presentation?displayProperty=fullName"/> namespace and `D` represents the <a href="xref:DocumentFormat.OpenXml.Drawing?displayProperty=fullName"/> namespace.

In the snippet below, a presentation is opened with `Presentation.Open` and the first <xref:DocumentFormat.OpenXml.Packaging.SlidePart>
is retrieved or added if the presentation does not already have a `SlidePart`.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/working_with_notes_slides/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/working_with_notes_slides/vb/Program.vb#snippet1)]
***


In this snippet the a `NoteSlide` is added to the `NoteSlidePart` if one does not already exist. 
The `NotesSlide` class constructor creates instances of the `CommonSlideData` class. 
The `CommonSlideData` class constructor creates an instance of the <xref:DocumentFormat.OpenXml.Presentation.ShapeTree> class, whose constructor in turn
creates additional class instances: an instance of the <xref:DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties> class, an
instance of the <xref:DocumentFormat.OpenXml.Presentation.GroupShapeProperties> class, and an instance
of the <xref:DocumentFormat.OpenXml.Presentation.Shape> class.


### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/working_with_notes_slides/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/working_with_notes_slides/vb/Program.vb#snippet2)]
***

The `Shape` constructor creates an instance of <xref:DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties>, <xref:DocumentFormat.OpenXml.Presentation.ShapeProperties>, and <xref:DocumentFormat.OpenXml.Presentation.TextBody> classes along with their required child elements. The `TextBody`
contains the <xref:DocumentFormat.OpenXml.Drawing.Paragraph>, that has a <xref:DocumentFormat.OpenXml.Drawing.Run>, which contains the text of the note. The slide part is then added to the notes slide part.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/working_with_notes_slides/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/working_with_notes_slides/vb/Program.vb#snippet3)]
***

The notes slide part created with the code above contains the following XML

```xml
<?xml version="1.0" encoding="utf-8"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="test Placeholder 3"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:r>
              <a:t>This is a test note!</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>
```


The following snippets add the required <xref:DocumentFormat.OpenXml.Packaging.NotesMasterPart> and <xref:DocumentFormat.OpenXml.Packaging.ThemePart> if they are missing.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/working_with_notes_slides/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/working_with_notes_slides/vb/Program.vb#snippet4)]
***

## Sample code

The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/working_with_notes_slides/cs/Program.cs#snippet0)]
### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/working_with_notes_slides/vb/Program.vb#snippet0)]
***

--------------------------------------------------------------------------------
## See also


[About the Open XML SDK for Office](../about-the-open-xml-sdk.md)  

[How to: Create a Presentation by Providing a File Name](how-to-create-a-presentation-document-by-providing-a-file-name.md)  

[How to: Insert a new slide into a presentation](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Apply a theme to a presentation](how-to-apply-a-theme-to-a-presentation.md)  
