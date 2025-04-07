---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 5471f369-ad02-41c3-a5d3-ebaf618d185a
title: 'How to: Add transitions between slides in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 04/03/2025
ms.localizationpriority: medium
---

# Add Transitions between slides in a presentation

This topic shows how to use the classes in the Open XML SDK to
add transition between all slides in a presentation programmatically.

## Getting a Presentation Object 

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document, call the
<xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*> method, that uses a file path, and a
Boolean value as the second parameter to specify whether a document is
editable. To open a document for read/write, specify the value `true` for this parameter as shown in the following
`using` statement. In this code, the file parameter, is a string that represents the path for the file from which you want to open the document.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/add_transition/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/add_transition/vb/Program.vb#snippet1)]
***

[!include[Using Statement](../includes/presentation/using-statement.md)] `ppt`.

## The Structure of the Transition

Transition element `<transition>` specifies the kind of slide transition that should be used to transition to the current slide from the
previous slide. That is, the transition information is stored on the slide that appears after the transition is
complete.

The following table lists the attributes of the Transition along
with the description of each.

| Attribute | Description |
|---|---|
| advClick (Advance on Click) | Specifies whether a mouse click advances the slide or not. If this attribute is not specified then a value of true is assumed. |
| advTm (Advance after time) | Specifies the time, in milliseconds, after which the transition should start. This setting can be used in conjunction with the advClick attribute. If this attribute is not specified then it is assumed that no auto-advance occurs. |
| spd (Transition Speed) |Specifies the transition speed that is to be used when transitioning from the current slide to the next. |

[*Example*: Consider the following example

```xml
      <p:transition spd="slow" advClick="1" advTm="3000">
        <p:randomBar dir="horz"/>
      </p:transition>
```
In the above example, the transition speed `<speed>` is set to slow (available options: slow, med, fast). Advance on Click `<advClick>` is set to true, and Advance after time `<advTm>` is set to 3000 milliseconds. The Random Bar child element `<randomBar>` describes the randomBar slide transition effect, which uses a set of randomly placed horizontal `<dir="horz">` or vertical `<dir="vert">` bars on the slide that continue to be added until the new slide is fully shown. *end example*]

A full list of Transition's child elements can be viewed here: <xref:DocumentFormat.OpenXml.Presentation.Transition>

## The Structure of the Alternate Content

Office Open XML defines a mechanism for the storage of content that is not defined by the ISO/IEC 29500 Office Open XML specification, such as extensions developed by future software applications that leverage the Office Open XML formats. This mechanism allows for the storage of a series of alternative representations of content, from which the consuming application can use the first alternative whose requirements are met.

Consider an application that creates a new transition object intended to specify the duration of the transition. This functionality is not defined in the Office Open XML specification. Using an AlternateContent block as follows allows specifying the duration `<p14:dur>` in milliseconds.

[*Example*: 
```xml
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
   xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
    <mc:Choice Requires="p14">
      <p:transition spd="slow" p14:dur="2000" advClick="1" advTm="3000">
        <p:randomBar/>
      </p:transition>
    </mc:Choice>
    <mc:Fallback>
      <p:transition spd="slow" advClick="1" advTm="3000">
        <p:randomBar/>
      </p:transition>
    </mc:Fallback>
  </mc:AlternateContent>
```

The Choice element in the above example requires the <xref:DocumentFormat.OpenXml.Linq.P14.dur*> attribute to specify the duration of the transition, and the Fallback element allows clients that do not support this namespace to see an appropriate alternative representation. *end example*]

More details on the P14 class can be found here: <xref:DocumentFormat.OpenXml.Presentation.Linq.P14>

## How the Sample Code Works ##
After opening the presentation file for read/write access in the using statement, the code gets the presentation part from the presentation document. Then, it retrieves the relationship IDs of all slides in the presentation and gets the slides part from the relationship ID. The code then checks if there are no existing transitions set on the slides and replaces them with a new RandomBarTransition.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/add_transition/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/add_transition/vb/Program.vb#snippet2)]
***

If there are currently no transitions on the slide, code creates new transition. In both cases as a fallback transition,
RandomBarTransition is used but without `P14:dur`(duration) to allow grater support for clients that aren't supporting this namespace

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/add_transition/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/add_transition/vb/Program.vb#snippet3)]
***

## Sample Code

Following is the complete sample code that you can use to add RandomBarTransition to all slides.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/add_transition/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/add_transition/vb/Program.vb#snippet0)]
***

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)




