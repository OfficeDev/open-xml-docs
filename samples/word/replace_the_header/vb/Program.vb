Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
    End Sub



    Public Sub AddHeaderFromTo(ByVal filepathFrom As String, ByVal filepathTo As String)
        ' Replace header in target document with header of source document.
        Using wdDoc As WordprocessingDocument = _
            WordprocessingDocument.Open(filepathTo, True)
            Dim mainPart As MainDocumentPart = wdDoc.MainDocumentPart

            ' Delete the existing header part.
            mainPart.DeleteParts(mainPart.HeaderParts)

            ' Create a new header part.
            Dim headerPart = mainPart.AddNewPart(Of HeaderPart)()

            ' Get Id of the headerPart.
            Dim rId As String = mainPart.GetIdOfPart(headerPart)

            ' Feed target headerPart with source headerPart.
            Using wdDocSource As WordprocessingDocument = _
                WordprocessingDocument.Open(filepathFrom, True)
                Dim firstHeader = wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault()

                If firstHeader IsNot Nothing Then
                    headerPart.FeedData(firstHeader.GetStream())
                End If
            End Using

            ' Get SectionProperties and Replace HeaderReference with new Id.
            Dim sectPrs = mainPart.Document.Body.Elements(Of SectionProperties)()
            For Each sectPr In sectPrs
                ' Delete existing references to headers.
                sectPr.RemoveAllChildren(Of HeaderReference)()

                ' Create the new header reference node.
                sectPr.PrependChild(Of HeaderReference)(New HeaderReference() With {.Id = rId})
            Next
        End Using
    End Sub
End Module