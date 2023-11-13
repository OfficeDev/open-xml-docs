Imports DocumentFormat.OpenXml.Packaging


Module MyModule

    Sub Main(args As String())
    End Sub

    ' To remove a document part from a package.
    Public Sub RemovePart(ByVal document As String)
        Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
        If ((mainPart.DocumentSettingsPart) IsNot Nothing) Then
            mainPart.DeletePart(mainPart.DocumentSettingsPart)
        End If
    End Sub
End Module