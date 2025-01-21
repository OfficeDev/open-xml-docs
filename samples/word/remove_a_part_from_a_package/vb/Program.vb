Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
        RemovePart(args(0))
    End Sub

    ' <Snippet0>
    ' To remove a document part from a package.
    Sub RemovePart(document As String)
        ' <Snippet1>
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            ' </Snippet1>
            ' <Snippet2>
            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart

            If mainPart IsNot Nothing AndAlso mainPart.DocumentSettingsPart IsNot Nothing Then
                mainPart.DeletePart(mainPart.DocumentSettingsPart)
            End If
            ' </Snippet2>
        End Using
    End Sub
    ' </Snippet0>
End Module
