Imports DocumentFormat.OpenXml.Packaging
Imports System.IO

Module Program
    Sub Main(args As String())
        ' <Snippet4>
        Dim fromDocument1 As String = args(0)
        Dim toDocument2 As String = args(1)

        CopyThemeContent(fromDocument1, toDocument2)
        ' </Snippet4>
    End Sub

    ' To copy contents of one package part.
    ' <Snippet0>
    ' <Snippet2>
    Sub CopyThemeContent(fromDocument1 As String, toDocument2 As String)
        ' <Snippet1>
        Using wordDoc1 As WordprocessingDocument = WordprocessingDocument.Open(fromDocument1, False)
            Using wordDoc2 As WordprocessingDocument = WordprocessingDocument.Open(toDocument2, True)
                ' </Snippet1>
                Dim themePart1 As ThemePart = wordDoc1?.MainDocumentPart?.ThemePart
                Dim themePart2 As ThemePart = wordDoc2?.MainDocumentPart?.ThemePart
                ' </Snippet2>

                ' If the theme parts are null, then there is nothing to copy.
                If themePart1 Is Nothing OrElse themePart2 Is Nothing Then
                    Return
                End If
                ' <Snippet3>
                Using streamReader As New StreamReader(themePart1.GetStream())
                    Using streamWriter As New StreamWriter(themePart2.GetStream(FileMode.Create))
                        streamWriter.Write(streamReader.ReadToEnd())
                    End Using
                End Using
                ' </Snippet3>
            End Using
        End Using
    End Sub
    ' </Snippet0>
End Module
