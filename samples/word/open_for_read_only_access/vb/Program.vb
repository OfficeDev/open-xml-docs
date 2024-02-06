' <Snippet0>
Imports System.IO
Imports System.IO.Packaging
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing


Module MyModule

    Sub Main(args As String())
        ' <Snippet6>
        OpenWordprocessingDocumentReadonly(args(0))
        ' </Snippet6>

        ' <Snippet7>
        OpenWordprocessingPackageReadonly(args(0))
        ' </Snippet7>

        ' <Snippet8>
        OpenWordprocessingStreamReadonly(args(0))
        ' </Snippet8>
    End Sub

    Public Sub OpenWordprocessingDocumentReadonly(ByVal filepath As String)

        ' <Snippet1>
        ' Open a WordprocessingDocument based on a filepath.
        Using wordProcessingDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, False)
            ' </Snippet1>

            ' <Snippet4>
            ' Assign a reference to the document body. 
            Dim mainDocumentPart As MainDocumentPart = If(wordProcessingDocument.MainDocumentPart, wordProcessingDocument.AddMainDocumentPart())

            If wordProcessingDocument.MainDocumentPart.Document Is Nothing Then
                wordProcessingDocument.MainDocumentPart.Document = New Document()
            End If

            If wordProcessingDocument.MainDocumentPart.Document.Body Is Nothing Then
                wordProcessingDocument.MainDocumentPart.Document.Body = New Body()
            End If

            Dim body As Body = wordProcessingDocument.MainDocumentPart.Document.Body
            ' </Snippet4>

            ' <Snippet5>
            ' Attempt to add some text.
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"))

            ' Call Save to generate an exception and show that access is read-only.
            'wordProcessingDocument.MainDocumentPart.Document.Save()
            ' </Snippet5>
        End Using
    End Sub

    Public Sub OpenWordprocessingPackageReadonly(ByVal filepath As String)
        ' <Snippet2>
        ' Open System.IO.Packaging.Package.
        Dim wordPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)

        ' Open a WordprocessingDocument based on a package.
        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(wordPackage)
            ' </Snippet2>

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

    Public Sub OpenWordprocessingStreamReadonly(ByVal filepath As String)
        ' <Snippet3>
        ' Get a stream of the wordprocessing document
        Using fileStream As FileStream = New FileStream(filepath, FileMode.Open)
            ' Open a WordprocessingDocument for read-only access based on a stream.
            Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(fileStream, False)
                ' </Snippet3>

                ' Assign a reference to the existing document body. 
                Dim body As Body = wordDocument.MainDocumentPart.Document.Body

                ' Attempt to add some text.
                Dim para As Paragraph = body.AppendChild(New Paragraph())
                Dim run As Run = para.AppendChild(New Run())
                run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingStreamReadonly"))

                ' Call Save to generate an exception and show that access is read-only.
                ' wordDocument.MainDocumentPart.Document.Save()
            End Using
        End Using
    End Sub
End Module
' </Snippet0>
