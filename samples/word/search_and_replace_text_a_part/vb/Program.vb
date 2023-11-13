Imports System.IO
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
    End Sub



    ' To search and replace content in a document part. 
    Public Sub SearchAndReplace(ByVal document As String)
        Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        Using (wordDoc)
            Dim docText As String = Nothing
            Dim sr As StreamReader = New StreamReader(wordDoc.MainDocumentPart.GetStream)

            Using (sr)
                docText = sr.ReadToEnd
            End Using

            Dim regexText As Regex = New Regex("Hello world!")
            docText = regexText.Replace(docText, "Hi Everyone!")
            Dim sw As StreamWriter = New StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create))

            Using (sw)
                sw.Write(docText)
            End Using
        End Using
    End Sub
End Module