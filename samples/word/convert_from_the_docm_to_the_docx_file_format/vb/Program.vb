Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
    End Sub



    ' Given a .docm file (with macro storage), remove the VBA 
    ' project, reset the document type, and save the document with a new name.
    Public Sub ConvertDOCMtoDOCX(ByVal fileName As String)
        Dim fileChanged As Boolean = False

        Using document As WordprocessingDocument =
            WordprocessingDocument.Open(fileName, True)

            ' Access the main document part.
            Dim docPart = document.MainDocumentPart

            ' Look for the vbaProject part. If it is there, delete it.
            Dim vbaPart = docPart.VbaProjectPart
            If vbaPart IsNot Nothing Then

                ' Delete the vbaProject part and then save the document.
                docPart.DeletePart(vbaPart)
                docPart.Document.Save()

                ' Change the document type to
                ' not macro-enabled.
                document.ChangeDocumentType(
                    WordprocessingDocumentType.Document)

                ' Track that the document has been changed.
                fileChanged = True
            End If
        End Using

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
    End Sub
End Module