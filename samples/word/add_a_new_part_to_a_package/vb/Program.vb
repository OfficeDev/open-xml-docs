Imports System.IO
Imports DocumentFormat.OpenXml.Packaging


Module MyModule
' To add a new document part to a package.
    Public Sub AddNewPart(ByVal document As String, ByVal fileName As String)
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
            
            Dim myXmlPart As CustomXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml)
            
            Using stream As New FileStream(fileName, FileMode.Open)
                myXmlPart.FeedData(stream)
            End Using
        End Using
    End Sub
End Module
