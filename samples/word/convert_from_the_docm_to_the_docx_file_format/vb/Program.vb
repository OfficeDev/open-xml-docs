Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        ConvertDOCMtoDOCX(args(0))
        ' </Snippet2>
    End Sub

    ' Given a .docm file (with macro storage), remove the VBA 
    ' project, reset the document type, and save the document with a new name.
    ' <Snippet0>
    ' <Snippet1>
    Sub ConvertDOCMtoDOCX(fileName As String)
        ' </Snippet1>
        ' <Snippet3>
        Dim fileChanged As Boolean = False

        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            ' Access the main document part.
            If document Is Nothing Or document.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart is null.")
            End If

            Dim docPart = document.MainDocumentPart
            ' </Snippet3>

            ' <Snippet4>
            ' Look for the vbaProject part. If it is there, delete it.
            Dim vbaPart = docPart.VbaProjectPart
            If vbaPart IsNot Nothing Then
                ' Delete the vbaProject part.
                docPart.DeletePart(vbaPart)
                ' </Snippet4>

                ' <Snippet5>
                ' Change the document type to
                ' not macro-enabled.
                document.ChangeDocumentType(WordprocessingDocumentType.Document)

                ' Track that the document has been changed.
                fileChanged = True
                ' </Snippet5>
            End If
        End Using

        ' <Snippet6>
        ' If anything goes wrong in this file handling,
        ' the code will raise an exception back to the caller.
        If fileChanged Then
            ' Create the new .docx filename.
            Dim newFileName = Path.ChangeExtension(fileName, ".docx")

            ' If it already exists, it will be deleted!
            If File.Exists(newFileName) Then
                File.Delete(newFileName)
            End If

            ' Rename the file.
            File.Move(fileName, newFileName)
        End If
        ' </Snippet6>
    End Sub
    ' </Snippet0>
End Module
