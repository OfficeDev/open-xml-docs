Imports System.IO
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
    End Sub



    ' This method can be used to replace a document part in a package.
    Public Sub ReplaceTheme(ByVal document As String, ByVal themeFile As String)
        Using wordDoc As WordprocessingDocument = _
            WordprocessingDocument.Open(document, True)
            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart

            ' Delete the old document part.
            mainPart.DeletePart(mainPart.ThemePart)

            ' Add a new document part and then add content.
            Dim themePart As ThemePart = mainPart.AddNewPart(Of ThemePart)()

            Using streamReader As New StreamReader(themeFile)
                Using streamWriter As _
                    New StreamWriter(themePart.GetStream(FileMode.Create))

                    streamWriter.Write(streamReader.ReadToEnd())
                End Using
            End Using
        End Using
    End Sub
End Module