using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using ShapeTree = DocumentFormat.OpenXml.Presentation.ShapeTree;
using ShapeProperties = DocumentFormat.OpenXml.Presentation.ShapeProperties;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties;
using NonVisualPictureProperties = DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties;
using NonVisualPictureDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties;
using Picture = DocumentFormat.OpenXml.Presentation.Picture;
using BlipFill = DocumentFormat.OpenXml.Presentation.BlipFill;
using DocumentFormat.OpenXml.Packaging;
using ApplicationNonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties;

// <Snippet0>
AddVideo(args[0], args[1], args[2]);

static void AddVideo(string filePath, string videoFilePath, string coverPicPath)
{

    string imgEmbedId = "rId4", embedId = "rId3", mediaEmbedId = "rId2";
    UInt32Value shapeId = 5;
    // <Snippet1>
    using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
    // </Snippet1>
    {

        if (presentationDocument.PresentationPart == null || presentationDocument.PresentationPart.Presentation.SlideIdList == null)
        {
            throw new NullReferenceException("Presentation Part is empty or there are no slides in it");
        }
        // <Snippet2>
        //Get presentation part
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        //Get slides ids.
        OpenXmlElementList slidesIds = presentationPart.Presentation.SlideIdList.ChildElements;


        //Get relationsipId of the last slide
        string? videoSldRelationshipId = ((SlideId) slidesIds[slidesIds.ToArray().Length - 1]).RelationshipId;

        if (videoSldRelationshipId == null)
        {
            throw new NullReferenceException("Slide id not found");
        }

        //Get slide part by relationshipID
        SlidePart? slidePart = (SlidePart) presentationPart.GetPartById(videoSldRelationshipId);
        // </Snippet2>

        // <Snippet3>
        // Create video Media Data Part (content type, extension)
        MediaDataPart mediaDataPart = presentationDocument.CreateMediaDataPart("video/mp4", ".mp4");

        //Get the video file and feed the stream
        using (Stream mediaDataPartStream = File.OpenRead(videoFilePath))
        {
            mediaDataPart.FeedData(mediaDataPartStream);
        }
        //Adds a VideoReferenceRelationship to the MainDocumentPart
        slidePart.AddVideoReferenceRelationship(mediaDataPart, embedId);

        //Adds a MediaReferenceRelationship to the SlideLayoutPart
        slidePart.AddMediaReferenceRelationship(mediaDataPart, mediaEmbedId);

        NonVisualDrawingProperties nonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = shapeId, Name = "video" };
        A.VideoFromFile videoFromFile = new A.VideoFromFile() { Link = embedId };

        ApplicationNonVisualDrawingProperties appNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();
        appNonVisualDrawingProperties.Append(videoFromFile);
       
        //adds sample image to the slide with id to be used as reference in blip
        ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png, imgEmbedId);
        using (Stream data = File.OpenRead(coverPicPath))
        {
            imagePart.FeedData(data);
        }
       
        if (slidePart!.Slide!.CommonSlideData!.ShapeTree == null)
        {
            throw new NullReferenceException("Presentation shape tree is empty");
        }

        //Getting existing shape tree object from PowerPoint
        ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

        // specifies the existence of a picture within a presentation.
        // It can have non-visual properties, a picture fill as well as shape properties attached to it.
        Picture picture = new Picture();
        NonVisualPictureProperties nonVisualPictureProperties = new NonVisualPictureProperties();

        A.HyperlinkOnClick hyperlinkOnClick = new A.HyperlinkOnClick() { Id = "", Action = "ppaction://media" };
        nonVisualDrawingProperties.Append(hyperlinkOnClick);

        NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties = new NonVisualPictureDrawingProperties();
        A.PictureLocks pictureLocks = new A.PictureLocks() { NoChangeAspect = true };
        nonVisualPictureDrawingProperties.Append(pictureLocks);

        ApplicationNonVisualDrawingPropertiesExtensionList appNonVisualDrawingPropertiesExtensionList = new ApplicationNonVisualDrawingPropertiesExtensionList();
        ApplicationNonVisualDrawingPropertiesExtension appNonVisualDrawingPropertiesExtension = new ApplicationNonVisualDrawingPropertiesExtension() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };
        // </Snippet3>

        // <Snippet4>
        P14.Media media = new() { Embed = mediaEmbedId };
        media.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        appNonVisualDrawingPropertiesExtension.Append(media);
        appNonVisualDrawingPropertiesExtensionList.Append(appNonVisualDrawingPropertiesExtension);
        appNonVisualDrawingProperties.Append(appNonVisualDrawingPropertiesExtensionList);

        nonVisualPictureProperties.Append(nonVisualDrawingProperties);
        nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        nonVisualPictureProperties.Append(appNonVisualDrawingProperties);

        //Prepare shape properties to display picture
        BlipFill blipFill = new BlipFill();
        A.Blip blip = new A.Blip() { Embed = imgEmbedId };
        // </Snippet4>

        A.Stretch stretch = new A.Stretch();
        A.FillRectangle fillRectangle = new A.FillRectangle();
        A.Transform2D transform2D = new A.Transform2D();
        A.Offset offset = new A.Offset() { X = 1524000L, Y = 857250L };
        A.Extents extents = new A.Extents() { Cx = 9144000L, Cy = 5143500L };
        A.PresetGeometry presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
        A.AdjustValueList adjValueList = new A.AdjustValueList();

        stretch.Append(fillRectangle);
        blipFill.Append(blip);
        blipFill.Append(stretch);
        transform2D.Append(offset);
        transform2D.Append(extents);
        presetGeometry.Append(adjValueList);

        ShapeProperties shapeProperties = new ShapeProperties();
        shapeProperties.Append(transform2D);
        shapeProperties.Append(presetGeometry);

        //adds all elements to the slide's shape tree
        picture.Append(nonVisualPictureProperties);
        picture.Append(blipFill);
        picture.Append(shapeProperties);

        shapeTree.Append(picture);

    }
}
// </Snippet0>