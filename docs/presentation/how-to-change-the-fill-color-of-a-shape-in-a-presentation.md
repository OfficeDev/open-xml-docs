---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: e3adae2b-c7e8-45f4-b1fc-93d937f4b3b1
title: 'How to: Change the fill color of a shape in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/02/2025
ms.localizationpriority: medium
---

# Change the fill color of a shape in a presentation

This topic shows how to use the classes in the Open XML SDK to
change the fill color of a shape on the first slide in a presentation
programmatically.



## Getting a Presentation Object 

In the Open XML SDK, the `PresentationDocument` class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call the
`Open` method that uses a file path, and a
Boolean value as the second parameter to specify whether a document is
editable. To open a document for read/write, specify the value `true` for this parameter as shown in the following
`using` statement. In this code, the file
parameter is a string that represents the path for the file from which
you want to open the document.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/presentation/change_the_fill_color_of_a_shape/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/change_the_fill_color_of_a_shape/vb/Program.vb#snippet1)]
***


[!include[Using Statement](../includes/presentation/using-statement.md)] `ppt`.


## The Structure of the Shape Tree 

The basic document structure of a PresentationML document consists of a
number of parts, among which is the Shape Tree (`<spTree/>`) element.

The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces the overall form of a `PresentationML` package.

> This element specifies all shapes within a slide. Contained within
> here are all the shapes, either grouped or not, that can be referenced
> on a given slide. As most objects within a slide are shapes, this
> represents the majority of content within a slide. Text and effects
> are attached to shapes that are contained within the `spTree` element.
> 
> [*Example*: Consider the following `PresentationML` slide

```xml
    <p:sld>
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr>
          ..
          </p:nvGrpSpPr>
          <p:grpSpPr>
          ..
          </p:grpSpPr>
          <p:sp>
          ..
          </p:sp>
        </p:spTree>
      </p:cSld>
      ..
    </p:sld>
```

> In the above example the shape tree specifies all the shape properties
> for this slide. *end example*]
> 
> &copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The following table lists the child elements of the Shape Tree along
with the description of each.

| Element | Description |
|---|---|
| cxnSp | Connection Shape |
| extLst | Extension List with Modification Flag |
| graphicFrame | Graphic Frame |
| grpSp | Group Shape |
| grpSpPr | Group Shape Properties |
| nvGrpSpPr | Non-Visual Properties for a Group Shape |
| pic | Picture |
| sp | Shape |

The following XML Schema fragment defines the contents of this element.

```xml
    <complexType name="CT_GroupShape">
       <sequence>
           <element name="nvGrpSpPr" type="CT_GroupShapeNonVisual" minOccurs="1" maxOccurs="1"/>
           <element name="grpSpPr" type="a:CT_GroupShapeProperties" minOccurs="1" maxOccurs="1"/>
           <choice minOccurs="0" maxOccurs="unbounded">
              <element name="sp" type="CT_Shape"/>
              <element name="grpSp" type="CT_GroupShape"/>
              <element name="graphicFrame" type="CT_GraphicalObjectFrame"/>
              <element name="cxnSp" type="CT_Connector"/>
              <element name="pic" type="CT_Picture"/>
           </choice>
           <element name="extLst" type="CT_ExtensionListModify" minOccurs="0" maxOccurs="1"/>
       </sequence>
    </complexType>
```

## How the Sample Code Works

After opening the presentation file for read/write access in the `using` statement, the code gets the presentation
part from the presentation document. Then it gets the relationship ID of
the first slide, and gets the slide part from the relationship ID.

> [!NOTE]
> The test file must have a shape on the first slide.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/change_the_fill_color_of_a_shape/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/change_the_fill_color_of_a_shape/vb/Program.vb#snippet2)]
***


The code then gets the shape tree that contains the shape whose fill
color is to be changed, and gets the first shape in the shape tree. It
then gets the shape properties of the shape and the solid fill reference of the shape properties,
and assigns a new fill color to the shape. There is no need to explicitly
save the file when inside of a using.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/change_the_fill_color_of_a_shape/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/change_the_fill_color_of_a_shape/vb/Program.vb#snippet3)]
***


## Sample Code

Following is the complete sample code that you can use to change the
fill color of a shape in a presentation. 
### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/change_the_fill_color_of_a_shape/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/change_the_fill_color_of_a_shape/vb/Program.vb#snippet0)]
***

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)




