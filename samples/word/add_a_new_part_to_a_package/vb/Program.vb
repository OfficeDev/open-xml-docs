Imports DocumentFormat.OpenXml.Packaging
Imports System.IO

Module Program
    Sub Main(args As String())
        ' <Snippet3>
        Dim document As String = args(0)
        Dim fileName As String = args(1)

        AddNewPart(document, fileName)
        ' </Snippet3>
    End Sub

    ' To add a new document part to a package.
    ' <Snippet0>
    Sub AddNewPart(document As String, fileName As String)
        ' <Snippet1>
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            ' </Snippet1>
            ' <Snippet2>
            Dim mainPart As MainDocumentPart = If(wordDoc.MainDocumentPart, wordDoc.AddMainDocumentPart())

            Dim myXmlPart As CustomXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml)

            Using stream As New FileStream(fileName, FileMode.Open)
                myXmlPart.FeedData(stream)
            End Using
            ' </Snippet2>
        End Using
    End Sub
    ' </Snippet0>
End Module

