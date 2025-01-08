Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports System.IO
Imports System.Xml

Module Program
    Sub Main(args As String())
        If File.Exists(args(0)) Then
            File.Delete(args(0))
        End If

        AddNewPart(args(0))
    End Sub

    ' <Snippet0>
    ' <Snippet1>
    Sub AddNewPart(document As String)
        ' Create a new word processing document.
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
            ' </Snippet1>
            ' <Snippet2>
            ' Add the MainDocumentPart part in the new word processing document.
            Dim mainDocPart As MainDocumentPart = wordDoc.AddMainDocumentPart()
            mainDocPart.Document = New Document()

            ' Add the CustomFilePropertiesPart part in the new word processing document.
            Dim customFilePropPart = wordDoc.AddCustomFilePropertiesPart()
            customFilePropPart.Properties = New DocumentFormat.OpenXml.CustomProperties.Properties()

            ' Add the CoreFilePropertiesPart part in the new word processing document.
            Dim coreFilePropPart = wordDoc.AddCoreFilePropertiesPart()
            Using writer As New XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8)
                writer.WriteRaw("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" &
                                "<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" />")
                writer.Flush()
            End Using
            ' </Snippet2>
            ' <Snippet3>
            ' Add the DigitalSignatureOriginPart part in the new word processing document.
            wordDoc.AddNewPart(Of DigitalSignatureOriginPart)("rId4")

            ' Add the ExtendedFilePropertiesPart part in the new word processing document.
            Dim extendedFilePropPart = wordDoc.AddNewPart(Of ExtendedFilePropertiesPart)("rId5")
            extendedFilePropPart.Properties = New DocumentFormat.OpenXml.ExtendedProperties.Properties()

            ' Add the ThumbnailPart part in the new word processing document.
            wordDoc.AddNewPart(Of ThumbnailPart)("image/jpeg", "rId6")
            ' </Snippet3>
        End Using
    End Sub
    ' </Snippet0>
End Module
