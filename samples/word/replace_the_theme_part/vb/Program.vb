Imports DocumentFormat.OpenXml.Packaging
Imports System
Imports System.IO

Module Program
    Sub Main(args As String())
        ' <Snippet4>
        Dim document As String = args(0)
        Dim themeFile As String = args(1)

        ReplaceTheme(document, themeFile)
        ' </Snippet4>
    End Sub

    ' <Snippet0>
    ' This method can be used to replace the theme part in a package.
    Sub ReplaceTheme(document As String, themeFile As String)
        ' <Snippet2>
        ' <Snippet1>
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            ' </Snippet1>
            If wordDoc?.MainDocumentPart?.ThemePart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body and/or ThemePart is null.")
            End If

            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart

            ' Delete the old document part.
            mainPart.DeletePart(mainPart.ThemePart)
            ' </Snippet2>
            ' <Snippet3>
            ' Add a new document part and then add content.
            Dim themePart As ThemePart = mainPart.AddNewPart(Of ThemePart)()

            Using streamReader As New StreamReader(themeFile)
                Using streamWriter As New StreamWriter(themePart.GetStream(FileMode.Create))
                    streamWriter.Write(streamReader.ReadToEnd())
                End Using
            End Using
            ' </Snippet3>
        End Using
    End Sub
    ' </Snippet0>
End Module
