---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: ae8c98d9-dd11-4b75-804c-165095d60ffd
title: 'How to: Insert a picture into a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
localization_priority: Priority
---
# Insert a picture into a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically add a picture to a word processing document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.IO;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using A = DocumentFormat.OpenXml.Drawing;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
```

```vb
    Imports System.IO
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
    Imports A = DocumentFormat.OpenXml.Drawing
    Imports DW = DocumentFormat.OpenXml.Drawing.Wordprocessing
    Imports PIC = DocumentFormat.OpenXml.Drawing.Pictures
```

--------------------------------------------------------------------------------
## Opening an Existing Document for Editing
To open an existing document, instantiate the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class as shown in
the following **using** statement. In the same
statement, open the word processing file at the specified **filepath** by using the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562234.aspx) method, with the
Boolean parameter set to **true** in order to
enable editing the document.

```csharp
    using (WordprocessingDocument wordprocessingDocument =
           WordprocessingDocument.Open(filepath, true)) 
    { 
        // Insert other code here. 
    }
```

```vb
    Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
        ' Insert other code here. 
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
that is used by the Open XML SDK to clean up resources) is automatically
called when the closing brace is reached. The block that follows the
**using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case
*wordprocessingDocument*. Because the **WordprocessingDocument** class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because
**Dispose** is automatically called when you
exit the block, you do not have to explicitly call **Save** and **Close**─as
long as you use **using**.


--------------------------------------------------------------------------------
## The XML Representation of the Graphic Object
The following text from the [ISO/IEC
29500](https://www.iso.org/standard/71691.html) specification
introduces the Graphic Object Data element.

> This element specifies the reference to a graphic object within the
> document. This graphic object is provided entirely by the document
> authors who choose to persist this data within the document.
> 
> [*Note*: Depending on the type of graphical object used not every
> generating application that supports the OOXML framework will have the
> ability to render the graphical object. *end note*]
> 
> © ISO/IEC29500: 2008.

The following XML Schema fragment defines the contents of this element

```xml
    <complexType name="CT_GraphicalObjectData">
       <sequence>
           <any minOccurs="0" maxOccurs="unbounded" processContents="strict"/>
       </sequence>
       <attribute name="uri" type="xsd:token"/>
    </complexType>
```

--------------------------------------------------------------------------------
## How the Sample Code Works
After you have opened the document, add the [ImagePart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.imagepart.aspx) object to the [MainDocumentPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.maindocumentpart.aspx) object by using a file
stream as shown in the following code segment.

```csharp
    MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
    using (FileStream stream = new FileStream(fileName, FileMode.Open))
    {
        imagePart.FeedData(stream);
    }
    AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
```

```vb
    Dim mainPart As MainDocumentPart = wordprocessingDocument.MainDocumentPart
    Dim imagePart As ImagePart = mainPart.AddImagePart(ImagePartType.Jpeg)
    Using stream As New FileStream(fileName, FileMode.Open)
        imagePart.FeedData(stream)
    End Using
    AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart))
```

To add the image to the body, first define the reference of the image.
Then, append the reference to the body. The element should be in a [Run](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.run.aspx).

```csharp
    // Define the reference of the image.
    var element =
         new Drawing(
             new DW.Inline(
                 new DW.Extent() { Cx = 990000L, Cy = 792000L },
                 new DW.EffectExtent()
                 {
                     LeftEdge = 0L,
                     TopEdge = 0L,
                     RightEdge = 0L,
                     BottomEdge = 0L
                 },
                 new DW.DocProperties()
                 {
                     Id = (UInt32Value)1U,
                     Name = "Picture 1"
                 },
                 new DW.NonVisualGraphicFrameDrawingProperties(
                     new A.GraphicFrameLocks() { NoChangeAspect = true }),
                 new A.Graphic(
                     new A.GraphicData(
                         new PIC.Picture(
                             new PIC.NonVisualPictureProperties(
                                 new PIC.NonVisualDrawingProperties()
                                 {
                                     Id = (UInt32Value)0U,
                                     Name = "New Bitmap Image.jpg"
                                 },
                                 new PIC.NonVisualPictureDrawingProperties()),
                             new PIC.BlipFill(
                                 new A.Blip(
                                     new A.BlipExtensionList(
                                         new A.BlipExtension()
                                         {
                                             Uri =
                                               "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                         })
                                 )
                                 {
                                     Embed = relationshipId,
                                     CompressionState =
                                     A.BlipCompressionValues.Print
                                 },
                                 new A.Stretch(
                                     new A.FillRectangle())),
                             new PIC.ShapeProperties(
                                 new A.Transform2D(
                                     new A.Offset() { X = 0L, Y = 0L },
                                     new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                 new A.PresetGeometry(
                                     new A.AdjustValueList()
                                 ) { Preset = A.ShapeTypeValues.Rectangle }))
                     ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
             )
             {
                 DistanceFromTop = (UInt32Value)0U,
                 DistanceFromBottom = (UInt32Value)0U,
                 DistanceFromLeft = (UInt32Value)0U,
                 DistanceFromRight = (UInt32Value)0U,
                 EditId = "50D07946"
             });

    // Append the reference to the body. The element should be in 
    // a DocumentFormat.OpenXml.Wordprocessing.Run.
    wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
```

```vb
    ' Define the image reference.
    Dim element = New Drawing( _
                          New DW.Inline( _
                      New DW.Extent() With {.Cx = 990000L, .Cy = 792000L}, _
                      New DW.EffectExtent() With {.LeftEdge = 0L, .TopEdge = 0L, .RightEdge = 0L, .BottomEdge = 0L}, _
                      New DW.DocProperties() With {.Id = CType(1UI, UInt32Value), .Name = "Picture1"}, _
                      New DW.NonVisualGraphicFrameDrawingProperties( _
                          New A.GraphicFrameLocks() With {.NoChangeAspect = True} _
                          ), _
                      New A.Graphic(New A.GraphicData( _
                                    New PIC.Picture( _
                                        New PIC.NonVisualPictureProperties( _
                                            New PIC.NonVisualDrawingProperties() With {.Id = 0UI, .Name = "Koala.jpg"}, _
                                            New PIC.NonVisualPictureDrawingProperties() _
                                            ), _
                                        New PIC.BlipFill( _
                                            New A.Blip( _
                                                New A.BlipExtensionList( _
                                                    New A.BlipExtension() With {.Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"}) _
                                                ) With {.Embed = relationshipId, .CompressionState = A.BlipCompressionValues.Print}, _
                                            New A.Stretch( _
                                                New A.FillRectangle() _
                                                ) _
                                            ), _
                                        New PIC.ShapeProperties( _
                                            New A.Transform2D( _
                                                New A.Offset() With {.X = 0L, .Y = 0L}, _
                                                New A.Extents() With {.Cx = 990000L, .Cy = 792000L}), _
                                            New A.PresetGeometry( _
                                                New A.AdjustValueList() _
                                                ) With {.Preset = A.ShapeTypeValues.Rectangle} _
                                            ) _
                                        ) _
                                    ) With {.Uri = "https://schemas.openxmlformats.org/drawingml/2006/picture"} _
                                ) _
                            ) With {.DistanceFromTop = 0UI, _
                                    .DistanceFromBottom = 0UI, _
                                    .DistanceFromLeft = 0UI, _
                                    .DistanceFromRight = 0UI} _
                        )

    ' Append the reference to the body, the element should be in 
    ' a DocumentFormat.OpenXml.Wordprocessing.Run.
    wordDoc.MainDocumentPart.Document.Body.AppendChild(New Paragraph(New Run(element)))
```

--------------------------------------------------------------------------------
## Sample Code
The following code example adds a picture to an existing word document.
In your code, you can call the **InsertAPicture** method by passing in the path of
the word document, and the path of the file that contains the picture.
For example, the following call inserts the picture "MyPic.jpg" into the
file "Word9.docx," located at the specified paths.

```csharp
    string document = @"C:\Users\Public\Documents\Word9.docx";
    string fileName = @"C:\Users\Public\Documents\MyPic.jpg";
    InsertAPicture(document, fileName);
```

```vb
    Dim document As String = "C:\Users\Public\Documents\Word9.docx"
    Dim fileName As String = "C:\Users\Public\Documents\MyPic.jpg"
    InsertAPicture(document, fileName)
```

After you run the code, look at the file "Word9.docx" to see the
inserted picture.

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    public static void InsertAPicture(string document, string fileName)
    {
        using (WordprocessingDocument wordprocessingDocument = 
            WordprocessingDocument.Open(document, true))
        {
            MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
        }
    }

    private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
    {
        // Define the reference of the image.
        var element =
             new Drawing(
                 new DW.Inline(
                     new DW.Extent() { Cx = 990000L, Cy = 792000L },
                     new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, 
                         RightEdge = 0L, BottomEdge = 0L },
                     new DW.DocProperties() { Id = (UInt32Value)1U, 
                         Name = "Picture 1" },
                     new DW.NonVisualGraphicFrameDrawingProperties(
                         new A.GraphicFrameLocks() { NoChangeAspect = true }),
                     new A.Graphic(
                         new A.GraphicData(
                             new PIC.Picture(
                                 new PIC.NonVisualPictureProperties(
                                     new PIC.NonVisualDrawingProperties() 
                                        { Id = (UInt32Value)0U, 
                                            Name = "New Bitmap Image.jpg" },
                                     new PIC.NonVisualPictureDrawingProperties()),
                                 new PIC.BlipFill(
                                     new A.Blip(
                                         new A.BlipExtensionList(
                                             new A.BlipExtension() 
                                                { Uri = 
                                                    "{28A0092B-C50C-407E-A947-70E740481C1C}" })
                                     ) 
                                     { Embed = relationshipId, 
                                         CompressionState = 
                                         A.BlipCompressionValues.Print },
                                     new A.Stretch(
                                         new A.FillRectangle())),
                                 new PIC.ShapeProperties(
                                     new A.Transform2D(
                                         new A.Offset() { X = 0L, Y = 0L },
                                         new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                     new A.PresetGeometry(
                                         new A.AdjustValueList()
                                     ) { Preset = A.ShapeTypeValues.Rectangle }))
                         ) { Uri = "https://schemas.openxmlformats.org/drawingml/2006/picture" })
                 ) { DistanceFromTop = (UInt32Value)0U, 
                     DistanceFromBottom = (UInt32Value)0U, 
                     DistanceFromLeft = (UInt32Value)0U, 
                     DistanceFromRight = (UInt32Value)0U, EditId = "50D07946" });

       // Append the reference to body, the element should be in a Run.
       wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
    }
```

```vb
    Public Sub InsertAPicture(ByVal document As String, ByVal fileName As String)
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            Dim mainPart As MainDocumentPart = wordprocessingDocument.MainDocumentPart

            Dim imagePart As ImagePart = mainPart.AddImagePart(ImagePartType.Jpeg)

            Using stream As New FileStream(fileName, FileMode.Open)
                imagePart.FeedData(stream)
            End Using

            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart))
        End Using
    End Sub

    Private Sub AddImageToBody(ByVal wordDoc As WordprocessingDocument, ByVal relationshipId As String)
        ' Define the reference of the image.
        Dim element = New Drawing( _
                              New DW.Inline( _
                          New DW.Extent() With {.Cx = 990000L, .Cy = 792000L}, _
                          New DW.EffectExtent() With {.LeftEdge = 0L, .TopEdge = 0L, .RightEdge = 0L, .BottomEdge = 0L}, _
                          New DW.DocProperties() With {.Id = CType(1UI, UInt32Value), .Name = "Picture1"}, _
                          New DW.NonVisualGraphicFrameDrawingProperties( _
                              New A.GraphicFrameLocks() With {.NoChangeAspect = True} _
                              ), _
                          New A.Graphic(New A.GraphicData( _
                                        New PIC.Picture( _
                                            New PIC.NonVisualPictureProperties( _
                                                New PIC.NonVisualDrawingProperties() With {.Id = 0UI, .Name = "Koala.jpg"}, _
                                                New PIC.NonVisualPictureDrawingProperties() _
                                                ), _
                                            New PIC.BlipFill( _
                                                New A.Blip( _
                                                    New A.BlipExtensionList( _
                                                        New A.BlipExtension() With {.Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"}) _
                                                    ) With {.Embed = relationshipId, .CompressionState = A.BlipCompressionValues.Print}, _
                                                New A.Stretch( _
                                                    New A.FillRectangle() _
                                                    ) _
                                                ), _
                                            New PIC.ShapeProperties( _
                                                New A.Transform2D( _
                                                    New A.Offset() With {.X = 0L, .Y = 0L}, _
                                                    New A.Extents() With {.Cx = 990000L, .Cy = 792000L}), _
                                                New A.PresetGeometry( _
                                                    New A.AdjustValueList() _
                                                    ) With {.Preset = A.ShapeTypeValues.Rectangle} _
                                                ) _
                                            ) _
                                        ) With {.Uri = "https://schemas.openxmlformats.org/drawingml/2006/picture"} _
                                    ) _
                                ) With {.DistanceFromTop = 0UI, _
                                        .DistanceFromBottom = 0UI, _
                                        .DistanceFromLeft = 0UI, _
                                        .DistanceFromRight = 0UI} _
                            )

        ' Append the reference to body, the element should be in a Run.
        wordDoc.MainDocumentPart.Document.Body.AppendChild(New Paragraph(New Run(element)))
    End Sub
```

--------------------------------------------------------------------------------
## See also


- [Open XML SDK 2.5 class library reference](https://docs.microsoft.com/office/open-xml/open-xml-sdk)
