
    private static SlidePart CreateSlidePart(PresentationPart presentationPart)        
            {
                SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
                    slidePart1.Slide = new Slide(
                            new CommonSlideData(
                                new ShapeTree(
                                    new P.NonVisualGroupShapeProperties(
                                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                        new P.NonVisualGroupShapeDrawingProperties(),
                                        new ApplicationNonVisualDrawingProperties()),
                                    new GroupShapeProperties(new TransformGroup()),
                                    new P.Shape(
                                        new P.NonVisualShapeProperties(
                                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                        new P.ShapeProperties(),
                                        new P.TextBody(
                                            new BodyProperties(),
                                            new ListStyle(),
                                            new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
                            new ColorMapOverride(new MasterColorMapping()));
                    return slidePart1;
             }

    new P.Shape(
              new P.NonVisualShapeProperties(
                  new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                  new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                  new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                  new BodyProperties(),
                  new ListStyle(),
                  new Paragraph(new EndParagraphRunProperties() { Language = "en-US" })))