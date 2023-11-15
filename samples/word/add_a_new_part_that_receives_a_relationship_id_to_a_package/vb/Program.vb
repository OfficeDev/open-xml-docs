Imports System.IO
Imports System.Xml
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing


Module MyModule

    Sub Main(args As String())
    End Sub

    Public Sub AddNewPart(ByVal document As String)
        ' Create a new word processing document.
        Dim wordDoc As WordprocessingDocument =
    WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)

        ' Add the MainDocumentPart part in the new word processing document.
        Dim mainDocPart = wordDoc.AddNewPart(Of MainDocumentPart) _
    ("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1")
        mainDocPart.Document = New Document()

        ' Add the CustomFilePropertiesPart part in the new word processing document.
        Dim customFilePropPart = wordDoc.AddCustomFilePropertiesPart()
        customFilePropPart.Properties = New DocumentFormat.OpenXml.CustomProperties.Properties()

        ' Add the CoreFilePropertiesPart part in the new word processing document.
        Dim coreFilePropPart = wordDoc.AddCoreFilePropertiesPart()
        Using writer As New XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create),
    System.Text.Encoding.UTF8)
            writer.WriteRaw(
    "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCr & vbLf &
    "<cp:coreProperties xmlns:cp=""https://schemas.openxmlformats.org/package/2006/metadata/core-properties""></cp:coreProperties>")
            writer.Flush()
        End Using

        ' Add the DigitalSignatureOriginPart part in the new word processing document.
        wordDoc.AddNewPart(Of DigitalSignatureOriginPart)("rId4")

        ' Add the ExtendedFilePropertiesPart part in the new word processing document.
        Dim extendedFilePropPart = wordDoc.AddNewPart(Of ExtendedFilePropertiesPart)("rId5")
        extendedFilePropPart.Properties =
    New DocumentFormat.OpenXml.ExtendedProperties.Properties()

        ' Add the ThumbnailPart part in the new word processing document.
        wordDoc.AddNewPart(Of ThumbnailPart)("image/jpeg", "rId6")

        wordDoc.Dispose()
    End Sub
End Module