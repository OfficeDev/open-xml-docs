Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports System.IO
Imports System.Text

Module Program
    Sub Main(args As String())
        CreateNewWordDocument(args(0))
    End Sub

    ' <Snippet0>
    ' To create a new package as a Word document.
    Sub CreateNewWordDocument(document As String)
        ' <Snippet1>
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
            ' </Snippet1>
            ' Set the content of the document so that Word can open it.
            Dim mainPart As MainDocumentPart = wordDoc.AddMainDocumentPart()

            SetMainDocumentContent(mainPart)
        End Using
    End Sub

    ' Set the content of MainDocumentPart.
    Sub SetMainDocumentContent(part As MainDocumentPart)
        Const docXml As String = "<?xml version=""1.0"" encoding=""utf-8""?>" &
                                 "<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" &
                                 "<w:body>" &
                                 "<w:p>" &
                                 "<w:r>" &
                                 "<w:t>Hello World</w:t>" &
                                 "</w:r>" &
                                 "</w:p>" &
                                 "</w:body>" &
                                 "</w:document>"

        Using stream As Stream = part.GetStream()
            Dim buf As Byte() = (New UTF8Encoding()).GetBytes(docXml)
            stream.Write(buf, 0, buf.Length)
        End Using
    End Sub
    ' </Snippet0>
End Module
