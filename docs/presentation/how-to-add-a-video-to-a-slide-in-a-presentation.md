---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 403abe97-7ab2-40ba-92c0-d6312a6d10c8
title: 'How to: Add a video to a slide in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 02/25/2025
ms.localizationpriority: medium
---

# Add a video to a slide in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to add a video to the first slide in a presentation
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
[!code-csharp[](../../samples/presentation/add_video/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/presentation/add_video/vb/Program.vb#snippet1)]
***


[!include[Using Statement](../includes/presentation/using-statement.md)] `ppt`.


## The Structure of the Video From File

The PresentationML document consists of a number of parts, among which is the Picture (`<pic/>`) element.

The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification introduces the overall form of a `PresentationML` package.

Video File (`<videoFile/>`) specifies the presence of a video file. It is defined within the non-visual properties of an object. The video is attached to an object, representing it shall be attached to an object as this is how it is represented within the document. The actual playing of the video however is done within the timing node list that is specified under the timing element.

Consider the following ``Picture`` object that has a video attached to it.

```xml
<p:pic>  
  <p:nvPicPr>  
    <p:cNvPr id="7" name="Rectangle 6">  
      <a:hlinkClick r:id="" action="ppaction://media"/>  
    </p:cNvPr>  
    <p:cNvPicPr>  
      <a:picLocks noRot="1"/>  
    </p:cNvPicPr>  
    <p:nvPr>  
      <a:videoFile r:link="rId1"/>  
    </p:nvPr>  
  </p:nvPicPr>  
</p:pic>
```

In the above example, we see that there is a single videoFile element attached to this picture. This picture is placed within the document just as a normal picture or shape would be. The id of this picture, namely 7 in this case, is used to refer to this videoFile element from within the timing node list. The Linked relationship id is used to retrieve the actual video file for playback purposes. 

&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The following XML Schema fragment defines the contents of videoFile.

```xml
<xsd:complexType name="CT_TLMediaNodeVideo">
	<xsd:sequence>
		<xsd:element name="cMediaNode" type="CT_TLCommonMediaNodeData" minOccurs="1" maxOccurs="1"/>
	</xsd:sequence>
	<xsd:attribute name="fullScrn" type="xsd:boolean" use="optional" default="false"/>
</xsd:complexType>
```

## How the Sample Code Works

After opening the presentation file for read/write access in the `using` statement, the code gets the presentation
part from the presentation document. Then it gets the relationship ID of
the last slide, and gets the slide part from the relationship ID.


### [C#](#tab/cs-2)
[!code-csharp[](../../samples/presentation/add_video/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/presentation/add_video/vb/Program.vb#snippet2)]
***

The code first creates a media data part for the video file to be added. With the video file stream open, it feeds the media data part object. Next, video and media relationship references are added to the slide using the provided embedId for future reference to the video file and mediaEmbedId for media reference.

An image part is then added with a sample picture to be used as a placeholder for the video. A picture object is created with various elements, such as Non-Visual Drawing Properties (`<cNvPr/>`), which specify non-visual canvas properties. This allows for additional information that does not affect the appearance of the picture to be stored. The `<videoFile/>` element, explained above, is also included. The HyperLinkOnClick (`<hlinkClick/>`) element specifies the on-click hyperlink information to be applied to a run of text or image. When the hyperlink text or image is clicked, the link is fetched. Non-Visual Picture Drawing Properties (`<cNvPicPr/>`) specify the non-visual properties for the picture canvas. For a detailed explanation of the elements used, please refer to [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)]

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/add_video/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/add_video/vb/Program.vb#snippet3)]
***

Next Media(CT_Media) element is created with use of previously referenced mediaEmbedId(Embedded Picture Reference). Blip element is also added, this element specifies the existence of an image (binary large image or picture) and contains a reference to the image data. Blip's Embed attribute is used to specify an placeholder image in the Image Part created previously.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/presentation/add_video/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/presentation/add_video/vb/Program.vb#snippet4)]
***

All other elements such Offset(`<off/>`), Stretch(`<stretch/>`), fillRectangle(`<fillRect/>`), are appended to the ShapeProperties(`<spPr/>`) and ShapeProperties are appended to the Picture element(`<pic/>`). Finally the picture element that incudes video is added to the ShapeTree(`<sp/>`) of the slide.

Following is the complete sample code that you can use to add video to the slide.

## Sample Code

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/add_video/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/add_video/vb/Program.vb#snippet0)]
***

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
