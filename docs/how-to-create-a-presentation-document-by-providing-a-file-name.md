---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 3d4a800e-64f0-4715-919f-a8f7d92a5c37
title: 'How to: Create a presentation document by providing a file name (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Create a presentation document by providing a file name (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 to
create a presentation document programmatically.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](http://www.microsoft.com/en-us/download/details.aspx?id=30425). You
must explicitly reference the following assemblies in your project:

-   WindowsBase

-   DocumentFormat.OpenXml (Installed by the Open XML SDK)

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

A presentation file, like all files defined by the Open XML standard,
consists of a package file container. This is the file that users see in
their file explorer; it usually has a .pptx extension. The package file
is represented in the Open XML SDK 2.5 by the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.PresentationDocument"><span
class="nolink">PresentationDocument</span></span> class. The
presentation document contains, among other parts, a presentation part.
The presentation part, represented in the Open XML SDK 2.5 by the <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.PresentationPart"><span
class="nolink">PresentationPart</span></span> class, contains the basic
*PresentationML* definition for the slide presentation. PresentationML
is the markup language used for creating presentations. Each package can
contain only one presentation part, and its root element must be
\<presentation\>.

The API calls used to create a new presentation document package are
relatively simple. The first step is to call the static <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.Create(System.String,DocumentFormat.OpenXml.PresentationDocumentType)"><span
class="nolink">Create(String, PresentationDocumentType)</span></span>
method of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.PresentationDocument"><span
class="nolink">PresentationDocument</span></span> class, as shown here
in the **CreatePresentation** procedure, which is the first part of the
complete code sample presented later in the article. The
**CreatePresentation** code calls the override of the <span
class="keyword">Create</span> method that takes as arguments the path to
the new document and the type of presentation document to be created.
The types of presentation documents available in that argument are
defined by a <span sdata="cer"
target="T:DocumentFormat.OpenXml.PresentationDocumentType"><span
class="nolink">PresentationDocumentType</span></span> enumerated value.

Next, the code calls <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.AddPresentationPart"><span
class="nolink">AddPresentationPart()</span></span>, which creates and
returns a **PresentationPart**. After the <span
class="keyword">PresentationPart</span> class instance is created, a new
root element for the presentation is added by setting the <span
sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.PresentationPart.Presentation"><span
class="nolink">Presentation</span></span> property equal to the instance
of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Presentation.Presentation"><span
class="nolink">Presentation</span></span> class returned from a call to
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

Using the Open XML SDK 2.5, you can create presentation structure and
content by using strongly-typed classes that correspond to
PresentationML elements. You can find these classes in the <span
sdata="cer" target="N:DocumentFormat.OpenXml.Presentation"><span
class="nolink">DocumentFormat.OpenXml.Presentation</span></span>
namespace. The following table lists the names of the classes that
correspond to the presentation, slide, slide master, slide layout, and
theme elements. The class that corresponds to the theme element is
actually part of the <span sdata="cer"
target="N:DocumentFormat.OpenXml.Drawing"><span
class="nolink">DocumentFormat.OpenXml.Drawing</span></span> namespace.
Themes are common to all Open XML markup languages.

| PresentationML Element | Open XML SDK 2.5 Class |
|---|---|
| &lt;presentation&gt; | Presentation |
| &lt;sld&gt; | Slide |
| &lt;sldMaster&gt; | SlideMaster |
| &lt;sldLayout&gt; | SlideLayout |
| &lt;theme&gt; | Theme |

The PresentationML code that follows is the XML in the presentation part
(in the file presentation.xml) for a simple presentation that contains
two slides.

```xml
    <p:presentation xmlns:p="…" … >
      <p:sldMasterIdLst>
        <p:sldMasterId xmlns:rel="http://…/relationships" rel:id="rId1"/>
      </p:sldMasterIdLst>
      <p:notesMasterIdLst>
        <p:notesMasterId xmlns:rel="http://…/relationships" rel:id="rId4"/>
      </p:notesMasterIdLst>
      <p:handoutMasterIdLst>
        <p:handoutMasterId xmlns:rel="http://…/relationships" rel:id="rId5"/>
      </p:handoutMasterIdLst>
      <p:sldIdLst>
        <p:sldId id="267" xmlns:rel="http://…/relationships" rel:id="rId2"/>
        <p:sldId id="256" xmlns:rel="http://…/relationships" rel:id="rId3"/>
      </p:sldIdLst>
      <p:sldSz cx="9144000" cy="6858000"/>
      <p:notesSz cx="6858000" cy="9144000"/>
    </p:presentation>
```

--------------------------------------------------------------------------------

Following is the complete sample C\# and VB code to create a
presentation, given a file path.

```csharp
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;
    using P = DocumentFormat.OpenXml.Presentation;
    using D = DocumentFormat.OpenXml.Drawing;

    namespace CreatePresentationDocument
    {
        class Program
        {
            static void Main(string[] args)
            {
                string filepath = @"C:\Users\username\Documents\PresentationFromFilename.pptx";
                CreatePresentation(filepath);
            } 

            public static void CreatePresentation(string filepath)
            {
                // Create a presentation at a specified file path. The presentation document type is pptx, by default.
                PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
                PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                CreatePresentationParts(presentationPart);            

                //Close the presentation handle
                presentationDoc.Close();
            } 

            private static void CreatePresentationParts(PresentationPart presentationPart)
            {
                SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
                SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
                SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
                NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
                DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

               presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

               SlidePart slidePart1;
               SlideLayoutPart slideLayoutPart1;
               SlideMasterPart slideMasterPart1;
               ThemePart themePart1;

                
                slidePart1 = CreateSlidePart(presentationPart);
                slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
                slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
                themePart1 = CreateTheme(slideMasterPart1); 
      
                slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
                presentationPart.AddPart(slideMasterPart1, "rId1");
                presentationPart.AddPart(themePart1, "rId5");            
            }

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

       private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
       {
           SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
           SlideMaster slideMaster = new SlideMaster(
           new CommonSlideData(new ShapeTree(
             new P.NonVisualGroupShapeProperties(
             new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
             new P.NonVisualGroupShapeDrawingProperties(),
             new ApplicationNonVisualDrawingProperties()),
             new GroupShapeProperties(new TransformGroup()),
             new P.Shape(
             new P.NonVisualShapeProperties(
               new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
               new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
               new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
             new P.ShapeProperties(),
             new P.TextBody(
               new BodyProperties(),
               new ListStyle(),
               new Paragraph())))),
           new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
           new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
           new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
           slideMasterPart1.SlideMaster = slideMaster;

           return slideMasterPart1;
        }

       private static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
       {
           ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
           D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

           D.ThemeElements themeElements1 = new D.ThemeElements(
           new D.ColorScheme(
             new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
             new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
             new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
             new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
             new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
             new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
             new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
             new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
             new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
             new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
             new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
             new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" })) { Name = "Office" },
             new D.FontScheme(
             new D.MajorFont(
             new D.LatinFont() { Typeface = "Calibri" },
             new D.EastAsianFont() { Typeface = "" },
             new D.ComplexScriptFont() { Typeface = "" }),
             new D.MinorFont(
             new D.LatinFont() { Typeface = "Calibri" },
             new D.EastAsianFont() { Typeface = "" },
             new D.ComplexScriptFont() { Typeface = "" })) { Name = "Office" },
             new D.FormatScheme(
             new D.FillStyleList(
             new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
             new D.GradientFill(
               new D.GradientStopList(
               new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                 new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
               new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 35000 },
               new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                new D.SaturationModulation() { Val = 350000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 100000 }
               ),
               new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
             new D.NoFill(),
             new D.PatternFill(),
             new D.GroupFill()),
             new D.LineStyleList(
             new D.Outline(
               new D.SolidFill(
               new D.SchemeColor(
                 new D.Shade() { Val = 95000 },
                 new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
               new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
             {
                 Width = 9525,
                 CapType = D.LineCapValues.Flat,
                 CompoundLineType = D.CompoundLineValues.Single,
                 Alignment = D.PenAlignmentValues.Center
             },
             new D.Outline(
               new D.SolidFill(
               new D.SchemeColor(
                 new D.Shade() { Val = 95000 },
                 new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
               new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
             {
                 Width = 9525,
                 CapType = D.LineCapValues.Flat,
                 CompoundLineType = D.CompoundLineValues.Single,
                 Alignment = D.PenAlignmentValues.Center
             },
             new D.Outline(
               new D.SolidFill(
               new D.SchemeColor(
                 new D.Shade() { Val = 95000 },
                 new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
               new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
             {
                 Width = 9525,
                 CapType = D.LineCapValues.Flat,
                 CompoundLineType = D.CompoundLineValues.Single,
                 Alignment = D.PenAlignmentValues.Center
             }),
             new D.EffectStyleList(
             new D.EffectStyle(
               new D.EffectList(
               new D.OuterShadow(
                 new D.RgbColorModelHex(
                 new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
             new D.EffectStyle(
               new D.EffectList(
               new D.OuterShadow(
                 new D.RgbColorModelHex(
                 new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
             new D.EffectStyle(
               new D.EffectList(
               new D.OuterShadow(
                 new D.RgbColorModelHex(
                 new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
             new D.BackgroundFillStyleList(
             new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
             new D.GradientFill(
               new D.GradientStopList(
               new D.GradientStop(
                 new D.SchemeColor(new D.Tint() { Val = 50000 },
                   new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
               new D.GradientStop(
                 new D.SchemeColor(new D.Tint() { Val = 50000 },
                   new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
               new D.GradientStop(
                 new D.SchemeColor(new D.Tint() { Val = 50000 },
                   new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 }),
               new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
             new D.GradientFill(
               new D.GradientStopList(
               new D.GradientStop(
                 new D.SchemeColor(new D.Tint() { Val = 50000 },
                   new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
               new D.GradientStop(
                 new D.SchemeColor(new D.Tint() { Val = 50000 },
                   new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 }),
               new D.LinearGradientFill() { Angle = 16200000, Scaled = true }))) { Name = "Office" });

           theme1.Append(themeElements1);
           theme1.Append(new D.ObjectDefaults());
           theme1.Append(new D.ExtraColorSchemeList());

           themePart1.Theme = theme1;
           return themePart1;

             }
        } 
    } 
```

```vb
    Imports System.Collections.Generic
    Imports System.Linq
    Imports System.Text
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Drawing
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Presentation
    Imports P = DocumentFormat.OpenXml.Presentation
    Imports D = DocumentFormat.OpenXml.Drawing


    Namespace CreatePresentationDocument
        Class Program
            Public Shared Sub Main(ByVal args As String())

                Dim filepath As String = "C:\Users\username\Documents\PresentationFromFilename.pptx"
                CreatePresentation(filepath)

            End Sub

            Public Shared Sub CreatePresentation(ByVal filepath As String)
                ' Create a presentation at a specified file path. The presentation document type is pptx, by default.
                Dim presentationDoc As PresentationDocument = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation)
                Dim presentationPart As PresentationPart = presentationDoc.AddPresentationPart()
                presentationPart.Presentation = New Presentation()

                CreatePresentationParts(presentationPart)

                'Close the presentation handle
                presentationDoc.Close()
            End Sub

            Private Shared Sub CreatePresentationParts(ByVal presentationPart As PresentationPart)
                Dim slideMasterIdList1 As New SlideMasterIdList(New SlideMasterId() With { _
                 .Id = CType(2147483648UI, UInt32Value), _
                 .RelationshipId = "rId1" _
                })
                Dim slideIdList1 As New SlideIdList(New SlideId() With { _
                 .Id = CType(256UI, UInt32Value), .RelationshipId = "rId2" _
                })
                Dim slideSize1 As New SlideSize() With { _
                 .Cx = 9144000, _
                 .Cy = 6858000, _
                 .Type = SlideSizeValues.Screen4x3 _
                }
                Dim notesSize1 As New NotesSize() With { _
                 .Cx = 6858000, _
                 .Cy = 9144000 _
                }
                Dim defaultTextStyle1 As New DefaultTextStyle()

                Dim slidePart1 As SlidePart
                Dim slideLayoutPart1 As SlideLayoutPart
                Dim slideMasterPart1 As SlideMasterPart
                Dim themePart1 As ThemePart

                presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1)

                slidePart1 = CreateSlidePart(presentationPart)
                slideLayoutPart1 = CreateSlideLayoutPart(slidePart1)
                slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1)
                themePart1 = CreateTheme(slideMasterPart1)

                slideMasterPart1.AddPart(slideLayoutPart1, "rId1")
                presentationPart.AddPart(slideMasterPart1, "rId1")
                presentationPart.AddPart(themePart1, "rId5")
            End Sub

            Private Shared Function CreateSlidePart(ByVal presentationPart As PresentationPart) As SlidePart
                Dim slidePart1 As SlidePart = presentationPart.AddNewPart(Of SlidePart)("rId2")
                slidePart1.Slide = New Slide(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                 .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "Title 1" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph(New EndParagraphRunProperties() With { _
                  .Language = "en-US" _
                }))))), New ColorMapOverride(New MasterColorMapping()))
                Return slidePart1
            End Function

            Private Shared Function CreateSlideLayoutPart(ByVal slidePart1 As SlidePart) As SlideLayoutPart
                Dim slideLayoutPart1 As SlideLayoutPart = slidePart1.AddNewPart(Of SlideLayoutPart)("rId1")
                Dim slideLayout As New SlideLayout(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), _
                    New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), _
                    New ListStyle(), New Paragraph(New EndParagraphRunProperties()))))), New ColorMapOverride(New MasterColorMapping()))
                slideLayoutPart1.SlideLayout = slideLayout
                Return slideLayoutPart1
            End Function

            Private Shared Function CreateSlideMasterPart(ByVal slideLayoutPart1 As SlideLayoutPart) As SlideMasterPart
                Dim slideMasterPart1 As SlideMasterPart = slideLayoutPart1.AddNewPart(Of SlideMasterPart)("rId1")
                Dim slideMaster As New SlideMaster(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), _
                    New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "Title Placeholder 1" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With { _
                  .Type = PlaceholderValues.Title _
                })), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph())))), New P.ColorMap() With { _
                  .Background1 = D.ColorSchemeIndexValues.Light1, _
                  .Text1 = D.ColorSchemeIndexValues.Dark1, _
                  .Background2 = D.ColorSchemeIndexValues.Light2, _
                  .Text2 = D.ColorSchemeIndexValues.Dark2, _
                  .Accent1 = D.ColorSchemeIndexValues.Accent1, _
                  .Accent2 = D.ColorSchemeIndexValues.Accent2, _
                  .Accent3 = D.ColorSchemeIndexValues.Accent3, _
                  .Accent4 = D.ColorSchemeIndexValues.Accent4, _
                  .Accent5 = D.ColorSchemeIndexValues.Accent5, _
                  .Accent6 = D.ColorSchemeIndexValues.Accent6, _
                  .Hyperlink = D.ColorSchemeIndexValues.Hyperlink, _
                  .FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink _
                }, New SlideLayoutIdList(New SlideLayoutId() With { _
                  .Id = CType(2147483649UI, UInt32Value), _
                  .RelationshipId = "rId1" _
                }), New TextStyles(New TitleStyle(), New BodyStyle(), New OtherStyle()))
                slideMasterPart1.SlideMaster = slideMaster

                Return slideMasterPart1
            End Function

            Private Shared Function CreateTheme(ByVal slideMasterPart1 As SlideMasterPart) As ThemePart
                Dim themePart1 As ThemePart = slideMasterPart1.AddNewPart(Of ThemePart)("rId5")
                Dim theme1 As New D.Theme() With { _
                  .Name = "Office Theme" _
                }

                Dim themeElements1 As New D.ThemeElements(New D.ColorScheme(New D.Dark1Color(New D.SystemColor() With { _
                  .Val = D.SystemColorValues.WindowText, _
                  .LastColor = "000000" _
                }), New D.Light1Color(New D.SystemColor() With { _
                  .Val = D.SystemColorValues.Window, _
                  .LastColor = "FFFFFF" _
                }), New D.Dark2Color(New D.RgbColorModelHex() With { _
                  .Val = "1F497D" _
                }), New D.Light2Color(New D.RgbColorModelHex() With { _
                  .Val = "EEECE1" _
                }), New D.Accent1Color(New D.RgbColorModelHex() With { _
                  .Val = "4F81BD" _
                }), New D.Accent2Color(New D.RgbColorModelHex() With { _
                  .Val = "C0504D" _
                }), _
                 New D.Accent3Color(New D.RgbColorModelHex() With { _
                  .Val = "9BBB59" _
                }), New D.Accent4Color(New D.RgbColorModelHex() With { _
                  .Val = "8064A2" _
                }), New D.Accent5Color(New D.RgbColorModelHex() With { _
                  .Val = "4BACC6" _
                }), New D.Accent6Color(New D.RgbColorModelHex() With { _
                  .Val = "F79646" _
                }), New D.Hyperlink(New D.RgbColorModelHex() With { _
                  .Val = "0000FF" _
                }), New D.FollowedHyperlinkColor(New D.RgbColorModelHex() With { _
                  .Val = "800080" _
                })) With { _
                  .Name = "Office" _
                }, New D.FontScheme(New D.MajorFont(New D.LatinFont() With { _
                  .Typeface = "Calibri" _
                }, New D.EastAsianFont() With { _
                  .Typeface = "" _
                }, New D.ComplexScriptFont() With { _
                  .Typeface = "" _
                }), New D.MinorFont(New D.LatinFont() With { _
                  .Typeface = "Calibri" _
                }, New D.EastAsianFont() With { _
                  .Typeface = "" _
                }, New D.ComplexScriptFont() With { _
                  .Typeface = "" _
                })) With { _
                  .Name = "Office" _
                }, New D.FormatScheme(New D.FillStyleList(New D.SolidFill(New D.SchemeColor() With { _
                  .Val = D.SchemeColorValues.PhColor _
                }), New D.GradientFill(New D.GradientStopList(New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }, New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 37000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 35000 _
                }, New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 15000 _
                }, New D.SaturationModulation() With { _
                  .Val = 350000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 100000 _
                }), New D.LinearGradientFill() With { _
                  .Angle = 16200000, _
                  .Scaled = True _
                }), New D.NoFill(), New D.PatternFill(), New D.GroupFill()), New D.LineStyleList(New D.Outline(New D.SolidFill(New D.SchemeColor(New D.Shade() With { _
                  .Val = 95000 _
                }, New D.SaturationModulation() With { _
                  .Val = 105000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }), New D.PresetDash() With { _
                  .Val = D.PresetLineDashValues.Solid _
                }) With { _
                  .Width = 9525, _
                  .CapType = D.LineCapValues.Flat, _
                  .CompoundLineType = D.CompoundLineValues.[Single], _
                  .Alignment = D.PenAlignmentValues.Center _
                }, New D.Outline(New D.SolidFill(New D.SchemeColor(New D.Shade() With { _
                  .Val = 95000 _
                }, New D.SaturationModulation() With { _
                  .Val = 105000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }), New D.PresetDash() With { _
                  .Val = D.PresetLineDashValues.Solid _
                }) With { _
                  .Width = 9525, _
                  .CapType = D.LineCapValues.Flat, _
                  .CompoundLineType = D.CompoundLineValues.[Single], _
                  .Alignment = D.PenAlignmentValues.Center _
                }, New D.Outline(New D.SolidFill(New D.SchemeColor(New D.Shade() With { _
                  .Val = 95000 _
                }, New D.SaturationModulation() With { _
                  .Val = 105000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }), New D.PresetDash() With { _
                  .Val = D.PresetLineDashValues.Solid _
                }) With { _
                  .Width = 9525, _
                  .CapType = D.LineCapValues.Flat, _
                  .CompoundLineType = D.CompoundLineValues.[Single], _
                  .Alignment = D.PenAlignmentValues.Center _
                }), New D.EffectStyleList(New D.EffectStyle(New D.EffectList(New D.OuterShadow(New D.RgbColorModelHex(New D.Alpha() With { _
                  .Val = 38000 _
                }) With { _
                  .Val = "000000" _
                }) With { _
                  .BlurRadius = 40000L, _
                  .Distance = 20000L, _
                  .Direction = 5400000, _
                  .RotateWithShape = False _
                })), New D.EffectStyle(New D.EffectList(New D.OuterShadow(New D.RgbColorModelHex(New D.Alpha() With { _
                  .Val = 38000 _
                }) With { _
                  .Val = "000000" _
                }) With { _
                  .BlurRadius = 40000L, _
                  .Distance = 20000L, _
                  .Direction = 5400000, _
                  .RotateWithShape = False _
                })), New D.EffectStyle(New D.EffectList(New D.OuterShadow(New D.RgbColorModelHex(New D.Alpha() With { _
                  .Val = 38000 _
                }) With { _
                  .Val = "000000" _
                }) With { _
                  .BlurRadius = 40000L, _
                  .Distance = 20000L, _
                  .Direction = 5400000, _
                  .RotateWithShape = False _
                }))), New D.BackgroundFillStyleList(New D.SolidFill(New D.SchemeColor() With { _
                  .Val = D.SchemeColorValues.PhColor _
                }), New D.GradientFill(New D.GradientStopList(New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }, New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }, New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }), New D.LinearGradientFill() With { _
                  .Angle = 16200000, _
                  .Scaled = True _
                }), New D.GradientFill(New D.GradientStopList(New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }, New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }), New D.LinearGradientFill() With { _
                  .Angle = 16200000, _
                  .Scaled = True _
                }))) With { _
                  .Name = "Office" _
                })

                theme1.Append(themeElements1)
                theme1.Append(New D.ObjectDefaults())
                theme1.Append(New D.ExtraColorSchemeList())

                themePart1.Theme = theme1
                Return themePart1

            End Function

        End Class

    End Namespace
```

--------------------------------------------------------------------------------

#### Concepts

[About the Open XML SDK 2.5 for Office](about-the-open-xml-sdk-2-5.md)  

[Structure of a PresentationML Document](structure-of-a-presentationml-document.md)  

[How to: Insert a new slide into a presentation (Open XML SDK)](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation (Open XML SDK)](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Retrieve the number of slides in a presentation document (Open XML SDK)](how-to-retrieve-the-number-of-slides-in-a-presentation-document.md)  

[How to: Apply a theme to a presentation (Open XML SDK)](how-to-apply-a-theme-to-a-presentation.md)  

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
