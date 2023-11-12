Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports P = DocumentFormat.OpenXml.Presentation

Module MyModule
    Function CreateNotesSlidePart(ByVal slidePart1 As SlidePart) As NotesSlidePart
        Dim notesSlidePart1 As NotesSlidePart = slidePart1.AddNewPart(Of NotesSlidePart)("rId6")
        Dim notesSlide As New NotesSlide(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With {
             .Id = 1UI,
             .Name = ""
            }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New _
                GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With {
             .Id = 2UI,
             .Name = ""
            }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With {
             .NoGrouping = True
            }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New _
                P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph(New EndParagraphRunProperties()))))),
            New ColorMapOverride(New MasterColorMapping()))
        notesSlidePart1.NotesSlide = notesSlide
        Return notesSlidePart1
    End Function
End Module
