Imports System.IO
Imports System.IO.Packaging
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing


Module MyModule
    Public Sub OpenWordprocessingDocumentReadonly(ByVal filepath As String)
        ' Open a WordprocessingDocument based on a filepath.
        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, False)
            ' Assign a reference to the existing document body. 
            Dim body As Body = wordDocument.MainDocumentPart.Document.Body

            ' Attempt to add some text.
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"))

            ' Call Save to generate an exception and show that access is read-only.
            ' wordDocument.MainDocumentPart.Document.Save()
        End Using
    End Sub

    Public Sub OpenWordprocessingPackageReadonly(ByVal filepath As String)
        ' Open System.IO.Packaging.Package.
        Dim wordPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)

        ' Open a WordprocessingDocument based on a package.
        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(wordPackage)
            ' Assign a reference to the existing document body. 
            Dim body As Body = wordDocument.MainDocumentPart.Document.Body

            ' Attempt to add some text.
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingPackageReadonly"))

            ' Call Save to generate an exception and show that access is read-only.
            ' wordDocument.MainDocumentPart.Document.Save()
        End Using

        ' Close the package.
        wordPackage.Close()
    End Sub
End Module
