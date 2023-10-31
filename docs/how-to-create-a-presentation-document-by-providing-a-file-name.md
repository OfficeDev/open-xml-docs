---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 3d4a800e-64f0-4715-919f-a8f7d92a5c37
title: 'How to: Create a presentation document by providing a file name (Open XML SDK)'
description: 'Learn how to create a presentation document by providing a file name using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: high
---

# Create a presentation document by providing a file name (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK to
create a presentation document programmatically.

To use the sample code in this topic, you must install the [Open XML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml). You
must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (Installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;
    using P = DocumentFormat.OpenXml.Presentation;
    using D = DocumentFormat.OpenXml.Drawing;
```

```vb
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXML.Drawing 
    Imports DocumentFormat.OpenXml.Presentation
    Imports P = DocumentFormat.OpenXml.Presentation
    Imports D = DocumentFormat.OpenXml.Drawing
```

--------------------------------------------------------------------------------

## Create a Presentation

A presentation file, like all files defined by the Open XML standard,
consists of a package file container. This is the file that users see in
their file explorer; it usually has a .pptx extension. The package file
is represented in the Open XML SDK by the [PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx) class. The
presentation document contains, among other parts, a presentation part.
The presentation part, represented in the Open XML SDK by the [PresentationPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationpart.aspx) class, contains the basic
*PresentationML* definition for the slide presentation. PresentationML
is the markup language used for creating presentations. Each package can
contain only one presentation part, and its root element must be
\<presentation\>.

The API calls used to create a new presentation document package are
relatively simple. The first step is to call the static [Create(String,PresentationDocumentType)](https://msdn.microsoft.com/library/office/cc535977.aspx)
method of the [PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx) class, as shown here
in the **CreatePresentation** procedure, which is the first part of the
complete code sample presented later in the article. The
**CreatePresentation** code calls the override of the **Create** method that takes as arguments the path to
the new document and the type of presentation document to be created.
The types of presentation documents available in that argument are
defined by a [PresentationDocumentType](https://msdn.microsoft.com/library/office/documentformat.openxml.presentationdocumenttype.aspx) enumerated value.

Next, the code calls [AddPresentationPart()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.addpresentationpart.aspx), which creates and
returns a **PresentationPart**. After the **PresentationPart** class instance is created, a new
root element for the presentation is added by setting the [Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationpart.presentation.aspx) property equal to the instance
of the [Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.presentation.aspx) class returned from a call to
the **Presentation** class constructor.

In order to create a complete, useable, and valid presentation, the code
must also add a number of other parts to the presentation package. In
the example code, this is taken care of by a call to a utility function
named **CreatePresentationsParts**. That function then calls a number of
other utility functions that, taken together, create all the
presentation parts needed for a basic presentation, including slide,
slide layout, slide master, and theme parts.

```csharp
    public static void CreatePresentation(string filepath)
    {
        // Create a presentation at a specified file path. The presentation document type is pptx, by default.
        PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
        PresentationPart presentationPart = presentationDoc.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        CreatePresentationParts(presentationPart);

        // Close the presentation handle
        presentationDoc.Close();
    }
```

Using the Open XML SDK, you can create presentation structure and
content by using strongly-typed classes that correspond to
PresentationML elements. You can find these classes in the [DocumentFormat.OpenXml.Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.aspx)
namespace. The following table lists the names of the classes that
correspond to the presentation, slide, slide master, slide layout, and
theme elements. The class that corresponds to the theme element is
actually part of the [DocumentFormat.OpenXml.Drawing](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.aspx) namespace.
Themes are common to all Open XML markup languages.

| PresentationML Element | Open XML SDK Class |
|---|---|
| &lt;presentation&gt; | [Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.presentation.aspx) |
| &lt;sld&gt; | [Slide](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slide.aspx) |
| &lt;sldMaster&gt; | [SlideMaster](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slidemaster.aspx) |
| &lt;sldLayout&gt; | [SlideLayout](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slidelayout.aspx) |
| &lt;theme&gt; | [Theme](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.theme.aspx) |

The PresentationML code that follows is the XML in the presentation part
(in the file presentation.xml) for a simple presentation that contains
two slides.

```xml
    <p:presentation xmlns:p="…" … >
      <p:sldMasterIdLst>
        <p:sldMasterId xmlns:rel="https://…/relationships" rel:id="rId1"/>
      </p:sldMasterIdLst>
      <p:notesMasterIdLst>
        <p:notesMasterId xmlns:rel="https://…/relationships" rel:id="rId4"/>
      </p:notesMasterIdLst>
      <p:handoutMasterIdLst>
        <p:handoutMasterId xmlns:rel="https://…/relationships" rel:id="rId5"/>
      </p:handoutMasterIdLst>
      <p:sldIdLst>
        <p:sldId id="267" xmlns:rel="https://…/relationships" rel:id="rId2"/>
        <p:sldId id="256" xmlns:rel="https://…/relationships" rel:id="rId3"/>
      </p:sldIdLst>
      <p:sldSz cx="9144000" cy="6858000"/>
      <p:notesSz cx="6858000" cy="9144000"/>
    </p:presentation>
```

--------------------------------------------------------------------------------

## Sample Code

Following is the complete sample C\# and VB code to create a
presentation, given a file path.

### [C#](#tab/cs)
[!code-csharp[](../samples/presentation/create_by_providing_a_file_name/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/presentation/create_by_providing_a_file_name/vb/Program.vb)]

--------------------------------------------------------------------------------

## See also 

[About the Open XML SDK for Office](about-the-open-xml-sdk.md)  

[Structure of a PresentationML Document](structure-of-a-presentationml-document.md)  

[How to: Insert a new slide into a presentation (Open XML SDK)](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation (Open XML SDK)](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Retrieve the number of slides in a presentation document (Open XML SDK)](how-to-retrieve-the-number-of-slides-in-a-presentation-document.md)  

[How to: Apply a theme to a presentation (Open XML SDK)](how-to-apply-a-theme-to-a-presentation.md)  

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
