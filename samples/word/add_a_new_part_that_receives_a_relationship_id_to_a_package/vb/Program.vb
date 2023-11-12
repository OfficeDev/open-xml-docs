Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports System.Xml


Module MyModule
Public Sub AddNewPart(ByVal document As String)
        ' Create a new word processing document.
        Dim wordDoc As WordprocessingDocument = _
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
        Using writer As New XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), _
    System.Text.Encoding.UTF8)
            writer.WriteRaw( _
    "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCr & vbLf & _
    "<cp:coreProperties xmlns:cp=""https://schemas.openxmlformats.org/package/2006/metadata/core-properties""></cp:coreProperties>")
            writer.Flush()
        End Using
        
        ' Add the DigitalSignatureOriginPart part in the new word processing document.
        wordDoc.AddNewPart(Of DigitalSignatureOriginPart)("rId4")
        
        ' Add the ExtendedFilePropertiesPart part in the new word processing document.
        Dim extendedFilePropPart = wordDoc.AddNewPart(Of ExtendedFilePropertiesPart)("rId5")
        extendedFilePropPart.Properties = _
    New DocumentFormat.OpenXml.ExtendedProperties.Properties()
        
        ' Add the ThumbnailPart part in the new word processing document.
        wordDoc.AddNewPart(Of ThumbnailPart)("image/jpeg", "rId6")
        
        wordDoc.Close()
    End Sub
End Module
