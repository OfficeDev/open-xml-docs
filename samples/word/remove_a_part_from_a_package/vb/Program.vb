Imports System
Imports DocumentFormat.OpenXml.Packaging


Module MyModule
' To remove a document part from a package.
    Public Sub RemovePart(ByVal document As String)
       Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, true)
       Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
       If (Not (mainPart.DocumentSettingsPart) Is Nothing) Then
          mainPart.DeletePart(mainPart.DocumentSettingsPart)
       End If
    End Sub
End Module
