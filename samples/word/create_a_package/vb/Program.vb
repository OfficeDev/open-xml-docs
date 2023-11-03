
    Imports System.Text
    Imports System.IO
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging


    ' To create a new package as a Word document.
    Public Sub CreateNewWordDocument(ByVal document As String)
        Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
        Using (wordDoc)
            ' Set the content of the document so that Word can open it.
            Dim mainPart As MainDocumentPart = wordDoc.AddMainDocumentPart
            SetMainDocumentContent(mainPart)
        End Using
    End Sub

    Public Sub SetMainDocumentContent(ByVal part As MainDocumentPart)
        Const docXml As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & _
            "<w:document xmlns:w=""https://schemas.openxmlformats.org/wordprocessingml/2006/main"">" & _
                "<w:body>" & _
                    "<w:p>" & _
                        "<w:r>" & _
                            "<w:t>Hello world!</w:t>" & _
                        "</w:r>" & _
                    "</w:p>" & _
                "</w:body>" & _
            "</w:document>"
        Dim stream1 As Stream = part.GetStream
        Dim utf8encoder1 As UTF8Encoding = New UTF8Encoding()
        Dim buf() As Byte = utf8encoder1.GetBytes(docXml)
        stream1.Write(buf, 0, buf.Length)
    End Sub
