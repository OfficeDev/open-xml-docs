

Module MyModule
Private Shared Function CreateNotesSlidePart(ByVal slidePart1 As SlidePart) As NotesSlidePart
            Dim notesSlidePart1 As NotesSlidePart = slidePart1.AddNewPart(Of NotesSlidePart)("rId6")
            Dim notesSlide As New NotesSlide(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
             .Id = DirectCast(1UI, UInt32Value), _
             .Name = "" _
            }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New _
                GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
             .Id = DirectCast(2UI, UInt32Value), _
             .Name = "" _
            }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
             .NoGrouping = True _
            }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New _
                P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph(New EndParagraphRunProperties()))))),
            New ColorMapOverride(New MasterColorMapping()))
            notesSlidePart1.NotesSlide = notesSlide
            Return notesSlidePart1
        End Function


    Private Shared Function CreateNotesSlidePart(ByVal slidePart1 As SlidePart) As NotesSlidePart
            Dim notesSlidePart1 As NotesSlidePart = slidePart1.AddNewPart(Of NotesSlidePart)("rId6")
            Dim notesSlide As New NotesSlide(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
             .Id = DirectCast(1UI, UInt32Value), _
             .Name = "" _
            }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New _
                GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
             .Id = DirectCast(2UI, UInt32Value), _
             .Name = "" _
            }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
             .NoGrouping = True _
            }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New _
                P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph(New EndParagraphRunProperties()))))),
            New ColorMapOverride(New MasterColorMapping()))
            notesSlidePart1.NotesSlide = notesSlide
            Return notesSlidePart1
        End Function
End Module
