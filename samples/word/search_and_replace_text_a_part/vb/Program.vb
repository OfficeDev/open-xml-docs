Imports DocumentFormat.OpenXml.Packaging
Imports System
Imports System.IO
Imports System.Text.RegularExpressions

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        SearchAndReplace(args(0))
        ' </Snippet2>
    End Sub

    ' To search and replace content in a document part.
    ' <Snippet0>
    Sub SearchAndReplace(document As String)
        ' <Snippet1>
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            ' </Snippet1>
            Dim docText As String = Nothing

            If wordDoc.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body is null.")
            End If

            Using sr As New StreamReader(wordDoc.MainDocumentPart.GetStream())
                docText = sr.ReadToEnd()
            End Using

            Dim regexText As New Regex("Hello World!")
            docText = regexText.Replace(docText, "Hi Everyone!")

            Using sw As New StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create))
                sw.Write(docText)
            End Using
        End Using
    End Sub
    ' </Snippet0>
End Module


