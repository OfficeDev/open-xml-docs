Imports System.IO
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
    End Sub



    ' To search and replace content in a document part. 
    Public Sub SearchAndReplace(ByVal document As String)
        Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        using (wordDoc)
            Dim docText As String = Nothing
            Dim sr As StreamReader = New StreamReader(wordDoc.MainDocumentPart.GetStream)

            using (sr)
                docText = sr.ReadToEnd
            End using

            Dim regexText As Regex = New Regex("Hello world!")
            docText = regexText.Replace(docText, "Hi Everyone!")
            Dim sw As StreamWriter = New StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create))

            using (sw)
                sw.Write(docText)
            End using
        End using
    End Sub
End Module