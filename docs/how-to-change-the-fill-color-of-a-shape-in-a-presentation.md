---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: e3adae2b-c7e8-45f4-b1fc-93d937f4b3b1
title: 'How to: Change the fill color of a shape in a presentation (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---

# Change the fill color of a shape in a presentation (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK to
change the fill color of a shape on the first slide in a presentation
programmatically.



## Getting a Presentation Object 

In the Open XML SDK, the **PresentationDocument** class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the
**Open** method that uses a file path, and a
Boolean value as the second parameter to specify whether a document is
editable. To open a document for read/write, specify the value **true** for this parameter as shown in the following
**using** statement. In this code, the file
parameter is a string that represents the path for the file from which
you want to open the document.

```csharp
    using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
    {
        // Insert other code here.
    }
```

```vb
    Using ppt As PresentationDocument = PresentationDocument.Open(docName, True)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case *ppt*.


## The Structure of the Shape Tree 

The basic document structure of a PresentationML document consists of a
number of parts, among which is the Shape Tree (**sp
Tree**) element.

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the overall form of a **PresentationML** package.

> This element specifies all shapes within a slide. Contained within
> here are all the shapes, either grouped or not, that can be referenced
> on a given slide. As most objects within a slide are shapes, this
> represents the majority of content within a slide. Text and effects
> are attached to shapes that are contained within the <span
> class="keyword">spTree** element.
> 
> [*Example*: Consider the following <span
> class="keyword">PresentationML** slide

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
> © ISO/IEC29500: 2008.

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

After opening the presentation file for read/write access in the **using** statement, the code gets the presentation
part from the presentation document. Then it gets the relationship ID of
the first slide, and gets the slide part from the relationship ID.

> [!NOTE]
> The test file must have a filled shape as the first shape on the first slide.

```csharp
    using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
    {
        // Get the relationship ID of the first slide.
        PresentationPart part = ppt.PresentationPart;
        OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
        string relId = (slideIds[0] as SlideId).RelationshipId;

        // Get the slide part from the relationship ID.
        SlidePart slide = (SlidePart)part.GetPartById(relId);
```

```vb
    Using ppt As PresentationDocument = PresentationDocument.Open(docName, True)

        ' Get the relationship ID of the first slide.
        Dim part As PresentationPart = ppt.PresentationPart
        Dim slideIds As OpenXmlElementList = part.Presentation.SlideIdList.ChildElements
        Dim relId As String = CType(slideIds(0), SlideId).RelationshipId

        ' Get the slide part from the relationship ID.
        Dim slide As SlidePart = CType(part.GetPartById(relId), SlidePart)
```

The code then gets the shape tree that contains the shape whose fill
color is to be changed, and gets the first shape in the shape tree. It
then gets the style of the shape and the fill reference of the style,
and assigns a new fill color to the shape. Finally it saves the modified
presentation.

```csharp
    if (slide != null)
    {
        // Get the shape tree that contains the shape to change.
        ShapeTree tree = slide.Slide.CommonSlideData.ShapeTree;

        // Get the first shape in the shape tree.
        Shape shape = tree.GetFirstChild<Shape>();

        if (shape != null)
        {
            // Get the style of the shape.
            ShapeStyle style = shape.ShapeStyle;

            // Get the fill reference.
            Drawing.FillReference fillRef = style.FillReference;

            // Set the fill color to SchemeColor Accent 6;
            fillRef.SchemeColor = new Drawing.SchemeColor();
            fillRef.SchemeColor.Val = Drawing.SchemeColorValues.Accent6;

            // Save the modified slide.
            slide.Slide.Save();
        }
    }
```

```vb
    If (Not (slide) Is Nothing) Then

        ' Get the shape tree that contains the shape to change.
        Dim tree As ShapeTree = slide.Slide.CommonSlideData.ShapeTree

        ' Get the first shape in the shape tree.
        Dim shape As Shape = tree.GetFirstChild(Of Shape)()

        If (Not (shape) Is Nothing) Then

            ' Get the style of the shape.
            Dim style As ShapeStyle = shape.ShapeStyle

            ' Get the fill reference.
            Dim fillRef As Drawing.FillReference = style.FillReference

            ' Set the fill color to SchemeColor Accent 6;
            fillRef.SchemeColor = New Drawing.SchemeColor
            fillRef.SchemeColor.Val = Drawing.SchemeColorValues.Accent6

            ' Save the modified slide.
            slide.Slide.Save()
        End If
    End If
```

## Sample Code

Following is the complete sample code that you can use to change the
fill color of a shape in a presentation. In your program, you can invoke
the method **SetPPTShapeColor** to change the
fill color in the file "Myppt3.pptx" by using the following call.

```csharp
    string docName = @"C:\Users\Public\Documents\Myppt3.pptx";
    SetPPTShapeColor(docName);
```

```vb
    Dim docName As String = "C:\Users\Public\Documents\Myppt3.pptx"
    SetPPTShapeColor(docName)
```

After running the program, examine the file "Myppt3.pptx" to see the
change in the fill color.

### [C#](#tab/cs)
[!code-csharp[](../samples/presentation/change_the_fill_color_of_a_shape/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/presentation/change_the_fill_color_of_a_shape/vb/Program.vb)]

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)




