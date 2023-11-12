Imports System.IO
Imports DocumentFormat.OpenXml.Packaging


Module MyModule
' To copy contents of one package part.
    Public Sub CopyThemeContent(ByVal fromDocument1 As String, ByVal toDocument2 As String)
       Dim wordDoc1 As WordprocessingDocument = WordprocessingDocument.Open(fromDocument1, False)
       Dim wordDoc2 As WordprocessingDocument = WordprocessingDocument.Open(toDocument2, True)
       Using (wordDoc2)
          Dim themePart1 As ThemePart = wordDoc1.MainDocumentPart.ThemePart
          Dim themePart2 As ThemePart = wordDoc2.MainDocumentPart.ThemePart
          Dim streamReader As StreamReader = New StreamReader(themePart1.GetStream())
          Dim streamWriter As StreamWriter = New StreamWriter(themePart2.GetStream(FileMode.Create))
          Using (streamWriter)
             streamWriter.Write(streamReader.ReadToEnd)
          End Using
       End Using
    End Sub
End Module
