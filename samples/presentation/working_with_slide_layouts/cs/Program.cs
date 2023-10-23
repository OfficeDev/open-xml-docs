
    private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
            {
                SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
                SlideLayout slideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                  new P.NonVisualGroupShapeProperties(
                  new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                  new P.NonVisualGroupShapeDrawingProperties(),
                  new ApplicationNonVisualDrawingProperties()),
                  new GroupShapeProperties(new TransformGroup()),
                  new P.Shape(
                  new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                  new P.ShapeProperties(),
                  new P.TextBody(
                    new BodyProperties(),
                    new ListStyle(),
                    new Paragraph(new EndParagraphRunProperties()))))),
                new ColorMapOverride(new MasterColorMapping()));
                slideLayoutPart1.SlideLayout = slideLayout;
                return slideLayoutPart1;
             }

    private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
            {
                SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
                SlideLayout slideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                  new P.NonVisualGroupShapeProperties(
                  new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                  new P.NonVisualGroupShapeDrawingProperties(),
                  new ApplicationNonVisualDrawingProperties()),
                  new GroupShapeProperties(new TransformGroup()),
                  new P.Shape(
                  new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                  new P.ShapeProperties(),
                  new P.TextBody(
                    new BodyProperties(),
                    new ListStyle(),
                    new Paragraph(new EndParagraphRunProperties()))))),
                new ColorMapOverride(new MasterColorMapping()));
                slideLayoutPart1.SlideLayout = slideLayout;
                return slideLayoutPart1;
             }