using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Linq;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using SlideHelpers;
using DocumentFormat.OpenXml.Drawing;


string pptxPath = args[0];
// <Snippet0>
// <Snippet1>
using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxPath, true) ?? throw new Exception("Presentation Document does not exist"))
{
    // Get the first slide in the presentation or use the InsertNewSlide.InsertNewSlideFromPresentation helper method to insert a new slide.
    SlidePart slidePart = presentationDocument.PresentationPart?.SlideParts.FirstOrDefault() ?? InsertNewSlide.InsertNewSlideFromPresentation(presentationDocument, 1, "my new slide");

    // Add a new NoteSlidePart if one does not already exist
    NotesSlidePart notesSlidePart = slidePart.NotesSlidePart ?? slidePart.AddNewPart<NotesSlidePart>();
    // </Snippet1>

    // <Snippet2>
    // Add a NoteSlide to the NoteSlidePart if one does not already exist.
    notesSlidePart.NotesSlide ??= new P.NotesSlide(
        new P.CommonSlideData(
            new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new D.Transform2D(
                        new D.Offset() { X = 0, Y = 0 },
                        new D.Extents() { Cx = 0, Cy = 0 },
                        new D.ChildOffset() { X = 0, Y = 0 },
                        new D.ChildExtents() { Cx = 0, Cy = 0 })),
                // <Snippet3>
                new P.Shape(
                    // </Snippet2>
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = 3, Name = "test Placeholder 3" },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.ShapeProperties(),
                    new P.TextBody(
                        new D.BodyProperties(),
                        new D.Paragraph(
                            new D.Run(
                                new D.Text("This is a test note!"))))))));

    notesSlidePart.AddPart(slidePart);
    // </Snippet3>

    // <Snippet4>
    // Add the required NotesMasterPart if it is missing
    NotesMasterPart notesMasterPart = notesSlidePart.NotesMasterPart ?? notesSlidePart.AddNewPart<NotesMasterPart>();

    // Add a NotesMaster to the NotesMasterPart if not present
    notesMasterPart.NotesMaster ??= new NotesMaster(
    new P.CommonSlideData(
        new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties() { Id = 1, Name = "tacocat" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties(
                    new P.PlaceholderShape() { Type = PlaceholderValues.Body, Index = 1 })),
            new P.GroupShapeProperties())),
    new P.ColorMap()
    {
        Background1 = D.ColorSchemeIndexValues.Light1,
        Background2 = D.ColorSchemeIndexValues.Light2,
        Text1 = D.ColorSchemeIndexValues.Dark1,
        Text2 = D.ColorSchemeIndexValues.Dark2,
        Accent1 = D.ColorSchemeIndexValues.Accent1,
        Accent2 = D.ColorSchemeIndexValues.Accent2,
        Accent3 = D.ColorSchemeIndexValues.Accent3,
        Accent4 = D.ColorSchemeIndexValues.Accent4,
        Accent5 = D.ColorSchemeIndexValues.Accent5,
        Accent6 = D.ColorSchemeIndexValues.Accent6,
        Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
        FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink,
    });

    // Add a new ThemePart for the NotesMasterPart
    ThemePart themePart = notesMasterPart.ThemePart ?? notesMasterPart.AddNewPart<ThemePart>();

    // Add the Theme if it is missing
    themePart.Theme ??= new Theme(
        new ThemeElements(
            new ColorScheme(
                new Dark1Color(
                    new SystemColor() { Val = SystemColorValues.WindowText }),
                new Light1Color(
                    new SystemColor() { Val = SystemColorValues.Window }),
                new Dark2Color(
                    new RgbColorModelHex() { Val = "f1d7be" }),
                new Light2Color(
                    new RgbColorModelHex() { Val = "171717" }),
                new Accent1Color(
                    new RgbColorModelHex() { Val = "ea9f7d" }),
                new Accent2Color(
                    new RgbColorModelHex() { Val = "168ecd" }),
                new Accent3Color(
                    new RgbColorModelHex() { Val = "e694db" }),
                new Accent4Color(
                    new RgbColorModelHex() { Val = "f0612a" }),
                new Accent5Color(
                    new RgbColorModelHex() { Val = "5fd46c" }),
                new Accent6Color(
                    new RgbColorModelHex() { Val = "b158d1" }),
                new D.Hyperlink(
                    new RgbColorModelHex() { Val = "699f82" }),
                new FollowedHyperlinkColor(
                    new RgbColorModelHex() { Val = "699f82" }))
            { Name = "Office2" },
            new D.FontScheme(
                new MajorFont(
                    new LatinFont(),
                    new EastAsianFont(),
                    new ComplexScriptFont()),
                new MinorFont(
                    new LatinFont(),
                    new EastAsianFont(),
                    new ComplexScriptFont()))
            { Name = "Office2" },
            new FormatScheme(
                new FillStyleList(
                    new NoFill(),
                    new SolidFill(),
                    new D.GradientFill(),
                    new D.BlipFill(),
                    new D.PatternFill(),
                    new GroupFill()),
                new LineStyleList(
                    new D.Outline(),
                    new D.Outline(),
                    new D.Outline()),
                new EffectStyleList(
                    new EffectStyle(
                        new EffectList()),
                    new EffectStyle(
                        new EffectList()),
                    new EffectStyle(
                        new EffectList())),
                new BackgroundFillStyleList(
                    new NoFill(),
                    new SolidFill(),
                    new D.GradientFill(),
                    new D.BlipFill(),
                    new D.PatternFill(),
                    new GroupFill()))
            { Name = "Office2" }),
        new ObjectDefaults(),
        new ExtraColorSchemeList());
    // </Snippet4>
}
// </Snippet0>